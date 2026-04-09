import requests
import json
import pandas as pd
import logging
import re
import hashlib
import base64
import os
import csv
from io import StringIO
from datetime import datetime, timezone, timedelta

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Configuração
VERSAO_SCRIPT = "1.6.0"

# Exporta Dados - URL base
ED_URL = os.getenv("ED_URL", "https://ws1.soc.com.br/WebSoc/exportadados")

# Empresa Principal
ED_EMPRESA_PRINCIPAL = os.getenv("ED_EMPRESA_PRINCIPAL", " ")

# Cadastro de Funcionários por situação (FOL).
ED_FUNC_CODIGO = os.getenv("ED_FUNC_CODIGO", " ")
ED_FUNC_CHAVE  = os.getenv("ED_FUNC_CHAVE", " ")

# Cadastro de Empresas - Não suporta JSON, CSV.
ED_EMPRESAS_CODIGO = os.getenv("ED_EMPRESAS_CODIGO", " ")
ED_EMPRESAS_CHAVE  = os.getenv("ED_EMPRESAS_CHAVE", " ")

# WebService Modelo 2
WSDL_URL         = os.getenv("WSDL_URL", "https://ws1.soc.com.br/WSSoc/FuncionarioModelo2Ws?wsdl")
WS_USUARIO       = os.getenv("WS_USUARIO", " ")
WS_SENHA         = os.getenv("WS_SENHA", " ")
WS_CHAVE_ACESSO  = os.getenv("WS_CHAVE_ACESSO", " ")
WS_RESPONSAVEL   = os.getenv("WS_RESPONSAVEL", " ")
WS_EMPRESA_PRINC = os.getenv("WS_EMPRESA_PRINC", " ")

# Empresas Exceção
EMPRESAS_EXCECAO = [
    "EMPRESA 1",
    "EMPRESA 2",
]

# Modo teste
CPF_TESTE     = os.getenv("SOC_CPF_TESTE", None)          # None = todos ou um CPF específico.
DRY_RUN = False # True = Teste | False = Produção
DATA_DEMISSAO = datetime.now().strftime("%d/%m/%Y")

# Validação
def validar_configuracoes():
    obrigatorios = {
        "ED_EMPRESA_PRINCIPAL": ED_EMPRESA_PRINCIPAL,
        "ED_FUNC_CODIGO":       ED_FUNC_CODIGO,
        "ED_FUNC_CHAVE":        ED_FUNC_CHAVE,
        "ED_EMPRESAS_CODIGO":   ED_EMPRESAS_CODIGO,
        "ED_EMPRESAS_CHAVE":    ED_EMPRESAS_CHAVE,
        "WS_USUARIO":           WS_USUARIO,
        "WS_SENHA":             WS_SENHA,
        "WS_CHAVE_ACESSO":      WS_CHAVE_ACESSO,
        "WS_RESPONSAVEL":       WS_RESPONSAVEL,
        "WS_EMPRESA_PRINC":     WS_EMPRESA_PRINC,
    }
    faltantes = [k for k, v in obrigatorios.items() if not v]
    if faltantes:
        raise ValueError(f"Configurações obrigatórias não preenchidas: {', '.join(faltantes)}")

def validar_cpf(cpf: str) -> bool:
    cpf = re.sub(r"\D", "", cpf)
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    d1 = 0 if (soma * 10 % 11) >= 10 else (soma * 10 % 11)
    if d1 != int(cpf[9]):
        return False
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    d2 = 0 if (soma * 10 % 11) >= 10 else (soma * 10 % 11)
    return d2 == int(cpf[10])

# SOAP
def extrair_erro_soap(resposta_texto: str) -> str | None:
    padroes = [
        r"<faultstring[^>]*>(.*?)</faultstring>",
        r"<descricaoErro[^>]*>(.*?)</descricaoErro>",
        r"<mensagem[^>]*>(.*?)</mensagem>",
    ]
    for padrao in padroes:
        match = re.search(padrao, resposta_texto, re.S | re.I)
        if match:
            return match.group(1).strip()
    return None

def montar_envelope_soap(codigo_func: str, codigo_empresa: str, situacao: str, data_demissao: str) -> str:
    nonce_bytes = os.urandom(16)
    nonce_b64   = base64.b64encode(nonce_bytes).decode("utf-8")
    now         = datetime.now(timezone.utc)
    created     = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    expires     = (now + timedelta(minutes=5)).strftime("%Y-%m-%dT%H:%M:%SZ")

    sha1 = hashlib.sha1()
    sha1.update(nonce_bytes)
    sha1.update(created.encode("utf-8"))
    sha1.update(WS_SENHA.encode("utf-8"))
    password_digest = base64.b64encode(sha1.digest()).decode("utf-8")

    return f"""<soapenv:Envelope
    xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
    xmlns:ser="http://services.soc.age.com/">
  <soapenv:Header>
    <wsse:Security
        xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"
        xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"
        soapenv:mustUnderstand="1">
      <wsu:Timestamp wsu:Id="TS-1">
        <wsu:Created>{created}</wsu:Created>
        <wsu:Expires>{expires}</wsu:Expires>
      </wsu:Timestamp>
      <wsse:UsernameToken wsu:Id="UsernameToken-1">
        <wsse:Username>{WS_USUARIO}</wsse:Username>
        <wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordDigest">{password_digest}</wsse:Password>
        <wsse:Nonce EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary">{nonce_b64}</wsse:Nonce>
        <wsu:Created>{created}</wsu:Created>
      </wsse:UsernameToken>
    </wsse:Security>
  </soapenv:Header>
  <soapenv:Body>
    <ser:importacaoFuncionario>
      <Funcionario>
        <atualizarFuncionario>true</atualizarFuncionario>
        <funcionarioWsVo>
          <chaveProcuraFuncionario>CODIGO</chaveProcuraFuncionario>
          <codigo>{codigo_func}</codigo>
          <tipoBuscaEmpresa>CODIGO_SOC</tipoBuscaEmpresa>
          <codigoEmpresa>{codigo_empresa}</codigoEmpresa>
          <situacao>{situacao}</situacao>
          <dataDemissao>{data_demissao}</dataDemissao>
        </funcionarioWsVo>
        <identificacaoWsVo>
          <chaveAcesso>{WS_CHAVE_ACESSO}</chaveAcesso>
          <codigoEmpresaPrincipal>{WS_EMPRESA_PRINC}</codigoEmpresaPrincipal>
          <codigoResponsavel>{WS_RESPONSAVEL}</codigoResponsavel>
          <codigoUsuario>{WS_USUARIO}</codigoUsuario>
        </identificacaoWsVo>
      </Funcionario>
    </ser:importacaoFuncionario>
  </soapenv:Body>
</soapenv:Envelope>"""

# Busca Empresas
def buscar_empresas() -> list:
    """
    Consulta Cadastro de Empresas em formato CSV e retorna lista de empresas ativas. Não suporta JSON, apenas CSV/XML/HTML.
    """
    logging.info("Buscando lista de empresas no SOC...")

    params = {
        "empresa":   ED_EMPRESA_PRINCIPAL,
        "codigo":    ED_EMPRESAS_CODIGO,
        "chave":     ED_EMPRESAS_CHAVE,
        "tipoSaida": "csv",
    }
    url = f"{ED_URL}?parametro={json.dumps(params)}"

    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()

        csv_reader = csv.DictReader(StringIO(r.text), delimiter=";")
        empresas = []
        for row in csv_reader:
            if row.get("ATIVO", "").strip() == "1":
                codigo = row.get("CODIGO", "").strip()
                nome   = row.get("RAZAOSOCIAL", row.get("NOMEABREVIADO", "")).strip()
                if codigo:
                    empresas.append({"CODIGO": codigo, "NOME": nome})

        logging.info(f"  {len(empresas)} empresas ativas encontradas.")
        for i, emp in enumerate(empresas[:10]):
            logging.info(f"    {i+1:>2}. [{emp['CODIGO']}] {emp['NOME'][:60]}")

        return empresas

    except Exception as e:
        logging.error(f"Erro ao buscar empresas: {e}")
        return []

# Busca Funcionários
def buscar_funcionarios_por_empresa(codigo_empresa: str, nome_empresa: str) -> list:
    """
    Busca funcionários ativos de uma empresa específica. Usa codigo_empresa (cliente) no parâmetro 'empresa'.
    """
    params = {
        "empresa":   codigo_empresa,   # Empresa cliente
        "codigo":    ED_FUNC_CODIGO,
        "chave":     ED_FUNC_CHAVE,
        "tipoSaida": "json",
        "ativo":     "Sim",
    }
    url = f"{ED_URL}?parametro={json.dumps(params)}"

    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        dados = r.json()

        if isinstance(dados, dict):
            for k in ("dados", "retorno", "data", "funcionarios"):
                if k in dados:
                    dados = dados[k]
                    break

        if not isinstance(dados, list):
            dados = [dados] if dados else []

        return dados

    except Exception as e:
        logging.warning(f"  Erro ao buscar funcionários de {nome_empresa}: {e}")
        return []

def buscar_funcionarios_ativos() -> list:
    """
    Itera por todas as empresas ativas do SOC e agrega os funcionários ativos. Ignora empresas da lista EMPRESAS_EXCECAO.
    """
    empresas = buscar_empresas()
    if not empresas:
        logging.error("Nenhuma empresa encontrada. Encerrando.")
        return []

    excecoes = {e.upper().strip() for e in EMPRESAS_EXCECAO}

    todos_funcionarios   = []
    empresas_processadas = 0
    empresas_ignoradas   = 0

    for emp in empresas:
        codigo     = emp["CODIGO"]
        nome       = emp["NOME"]
        nome_upper = nome.upper()

        if not codigo or codigo == "0":
            continue

        if nome_upper in excecoes:
            logging.info(f"  IGNORADA (exceção): {nome}")
            empresas_ignoradas += 1
            continue

        logging.info(f"  Processando: {nome} [{codigo}]")
        funcionarios = buscar_funcionarios_por_empresa(codigo, nome)

        if funcionarios:
            for func in funcionarios:
                func["NOMEEMPRESA"]   = nome
                func["CODIGOEMPRESA"] = codigo
            logging.info(f"    {len(funcionarios)} funcionário(s) ativo(s)")
            todos_funcionarios.extend(funcionarios)
            empresas_processadas += 1
        else:
            logging.info("    Nenhum funcionário ativo")

    logging.info("")
    logging.info("📊 RESUMO DA BUSCA:")
    logging.info(f"  Empresas ativas     : {len(empresas)}")
    logging.info(f"  Processadas         : {empresas_processadas}")
    logging.info(f"  Ignoradas (exceção) : {empresas_ignoradas}")
    logging.info(f"  Total funcionários  : {len(todos_funcionarios)}")

    if CPF_TESTE:
        cpf_alvo = re.sub(r"\D", "", CPF_TESTE)
        todos_funcionarios = [
            f for f in todos_funcionarios
            if re.sub(r"\D", "", str(f.get("CPF", ""))) == cpf_alvo
        ]
        logging.info(f"  MODO TESTE: {len(todos_funcionarios)} funcionário(s) para CPF {CPF_TESTE}")

    return todos_funcionarios

# Inativacação SOAP
def inativar_funcionario(func: dict) -> tuple[bool, str]:
    nome        = func.get("NOME", "Desconhecido")
    codigo      = str(func.get("CODIGO", ""))
    cod_empresa = str(func.get("CODIGOEMPRESA", WS_EMPRESA_PRINC))
    nome_emp    = str(func.get("NOMEEMPRESA", "")).upper().strip()
    situacao    = str(func.get("SITUACAO", "")).upper()

    # Checagem Exceção
    if nome_emp in {e.upper() for e in EMPRESAS_EXCECAO}:
        return False, f"Empresa '{nome_emp}' na lista de exceção"

    # Já inativo
    if situacao in ("INATIVO", "DEMITIDO", "I", "D"):
        return False, "Funcionário já está inativo/demitido"

    # Código SOC
    if not codigo or codigo in ("", "None", "0"):
        return False, "Campo CODIGO ausente"

    # Simulação
    if DRY_RUN:
        return True, "[DRY-RUN] Simulação - nenhuma alteração enviada"

    envelope = montar_envelope_soap(codigo, cod_empresa, "INATIVO", DATA_DEMISSAO)
    headers  = {"Content-Type": "text/xml;charset=UTF-8", "SOAPAction": '""'}

    try:
        endpoint = WSDL_URL.replace("?wsdl", "")
        resp = requests.post(endpoint, data=envelope.encode("utf-8"), headers=headers, timeout=30)

        logging.debug(f"HTTP {resp.status_code} | {resp.text[:400]}")

        if resp.status_code != 200:
            msg = extrair_erro_soap(resp.text) or f"HTTP {resp.status_code}"
            return False, msg

        if "encontrouErro>true" in resp.text or "encontrouErro&gt;true" in resp.text:
            msg = extrair_erro_soap(resp.text) or "Erro de negócio não especificado"
            return False, msg

        if "atualizouFuncionario>true" in resp.text:
            return True, "Inativado com sucesso"

        return True, "Requisição aceita pelo servidor"

    except requests.exceptions.Timeout:
        return False, "Timeout - servidor demorou mais de 30s"
    except Exception as e:
        return False, f"Erro técnico: {e}"

# Log
def configurar_log():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[
            logging.FileHandler(f"inativacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log", encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )

# Principal
def main():
    configurar_log()

    try:
        validar_configuracoes()
    except ValueError as e:
        logging.error(str(e))
        return

    if CPF_TESTE and not validar_cpf(CPF_TESTE):
        logging.error(f"CPF de teste inválido: {CPF_TESTE}")
        return

    ambiente = f"TESTE (CPF: {CPF_TESTE})" if CPF_TESTE else "PRODUÇÃO"

    logging.info("=" * 70)
    logging.info("   INATIVAÇÃO DE FUNCIONÁRIOS - SOC WebService Modelo 2")
    logging.info(f"  Versão        : {VERSAO_SCRIPT}")
    logging.info(f"  Modo          : {ambiente}")
    logging.info(f"  DRY-RUN       : {' SIM (nenhuma alteração será enviada)' if DRY_RUN else ' NÃO'}")
    logging.info(f"  Data demissão : {DATA_DEMISSAO}")
    logging.info("=" * 70)

    if DRY_RUN:
        logging.warning("    MODO DRY-RUN ATIVO - nenhum registro será alterado no SOC")
        logging.warning("=" * 70)

    funcionarios = buscar_funcionarios_ativos()
    if not funcionarios:
        logging.error(" Nenhum funcionário encontrado para processamento.")
        return

    total = len(funcionarios)
    logging.info(f"\n Total a processar: {total}")
    logging.info("=" * 70)

    resultados = []
    n_sucesso  = 0
    n_ignorado = 0
    n_erro     = 0

    for i, func in enumerate(funcionarios, 1):
        nome = func.get("NOME", "Desconhecido")
        cpf  = func.get("CPF", "")
        logging.info(f"[{i:>4}/{total}]  {nome}  |  CPF: {cpf}")

        ok, msg = inativar_funcionario(func)

        ignorado_kw = ("exceção", "já está", "ausente", "dry-run", "inativo", "simulação")
        if ok:
            n_sucesso += 1
            status = "SUCESSO"
            logging.info(f"           {msg}")
        elif any(k in msg.lower() for k in ignorado_kw):
            n_ignorado += 1
            status = "IGNORADO"
            logging.info(f"           {msg}")
        else:
            n_erro += 1
            status = "ERRO"
            logging.warning(f"           {msg}")

        resultados.append({
            "Empresa":        func.get("NOMEEMPRESA", ""),
            "Codigo_Empresa": func.get("CODIGOEMPRESA", ""),
            "Codigo_SOC":     func.get("CODIGO", ""),
            "Matricula":      func.get("MATRICULAFUNCIONARIO", ""),
            "Funcionario":    nome,
            "CPF":            cpf,
            "Situacao_Antes": func.get("SITUACAO", ""),
            "Status":         status,
            "Detalhe":        msg,
            "Processado_Em":  datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Timestamp_Envio": datetime.now().isoformat(),
            "Ambiente":       "DRY-RUN" if DRY_RUN else ambiente,
            "Versao_Script":  VERSAO_SCRIPT,
        })

    arquivo_excel = f"Log_Inativacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    pd.DataFrame(resultados).to_excel(arquivo_excel, index=False)

    logging.info("")
    logging.info("=" * 70)
    logging.info("   Resultado")
    logging.info(f"  Total processado    : {total}")
    logging.info(f"  Inativados        : {n_sucesso}")
    logging.info(f"  Ignorados         : {n_ignorado}")
    logging.info(f"  Erros             : {n_erro}")
    logging.info(f"  Relatório gerado  : {arquivo_excel}")
    logging.info("=" * 70)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.info("\n Execução interrompida pelo usuário.")
    except Exception as e:
        logging.exception(f" Erro fatal: {e}")
    finally:
        input("\n Pressione Enter para fechar...")
