import requests
import json
import pandas as pd
import logging
import re
import hashlib
import base64
import os
from datetime import datetime, timezone, timedelta
from zeep.wsse import UsernameToken
from zeep.exceptions import Fault

# Exporta Dados (leitura)
ED_URL      = "https://ws1.soc.com.br/WebSoc/exportadados"
ED_EMPRESA  = "SEU_CODIGO_EMPRESA"
ED_CODIGO   = "SEU_CODIGO_EXPORTA_DADOS"
ED_CHAVE    = "SUA_CHAVE_EXPORTA_DADOS"

# WebService Modelo 2 (gravação)
WSDL_URL         = "https://ws1.soc.com.br/WSSoc/FuncionarioModelo2Ws?wsdl"
WS_USUARIO       = "SEU_USUARIO_SOC"
WS_SENHA         = "SUA_SENHA_SOC"
WS_CHAVE_ACESSO  = "SUA_CHAVE_ACESSO_WS"
WS_RESPONSAVEL   = "SEU_CODIGO_RESPONSAVEL"
WS_EMPRESA_PRINC = "SEU_CODIGO_EMPRESA"

# Regras de negócio
EMPRESAS_EXCECAO = [
    "NOME EXATO DA EMPRESA 1",
    "NOME EXATO DA EMPRESA 2",
]

CPF_TESTE = None  # ex: "12345678900"
DATA_DEMISSAO = datetime.now().strftime("%d/%m/%Y")

# WS-SECURITY — PasswordDigest
def montar_envelope_soap(
    codigo_func: str,
    codigo_empresa: str,
    situacao: str,
    data_demissao: str
) -> str:
    
    """
    Monta o envelope SOAP completo com cabeçalho WS-Security PasswordDigest. O SOC exige autenticação no padrão: PasswordDigest = Base64( SHA-1( Nonce_bytes + Created_utf8 + Password_utf8 ) )
    O funcionário é identificado pelo campo CODIGO (código interno SOC), usando chaveProcuraFuncionario=CODIGO.
    """

    # Nonce
    nonce_bytes = os.urandom(16)
    nonce_b64   = base64.b64encode(nonce_bytes).decode("utf-8")

    # Timestamps
    now     = datetime.now(timezone.utc)
    created = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    expires = (now + timedelta(minutes=5)).strftime("%Y-%m-%dT%H:%M:%SZ")

    # PasswordDigest
    sha1 = hashlib.sha1()
    sha1.update(nonce_bytes)
    sha1.update(created.encode("utf-8"))
    sha1.update(WS_SENHA.encode("utf-8"))
    password_digest = base64.b64encode(sha1.digest()).decode("utf-8")

    envelope = f"""<soapenv:Envelope
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

    return envelope

# Exporta Dados
def buscar_funcionarios_ativos() -> list:

    logging.info("Consultando Exporta Dados...")

    params = {
        "empresa":   ED_EMPRESA,
        "codigo":    ED_CODIGO,
        "chave":     ED_CHAVE,
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
            dados = [dados]

        logging.info(f"{len(dados)} registros retornados pelo Exporta Dados.")

        if CPF_TESTE:
            cpf_alvo = re.sub(r"\D", "", CPF_TESTE)
            dados = [
                f for f in dados
                if re.sub(r"\D", "", str(f.get("CPF", ""))) == cpf_alvo
            ]
            logging.info(
                f"MODO TESTE — {len(dados)} registro(s) para CPF {CPF_TESTE}."
            )

        return dados

    except Exception as e:
        logging.error(f"Erro ao consultar Exporta Dados: {e}")
        return []

# WebService Modelo 2 - SOAP
def inativar_funcionario(func: dict) -> tuple[bool, str]:

    """
    Monta e envia o envelope SOAP para inativar um funcionário no SOC. Usa requests.post diretamente (sem Zeep) para garantir controle total sobre o envelope WS-Security PasswordDigest.
    Retorna (True, mensagem) em caso de sucesso ou (False, motivo)  em caso de falha/ignorado.
    """

    nome        = func.get("NOME", "Desconhecido")
    codigo      = str(func.get("CODIGO", ""))
    cod_empresa = str(func.get("CODIGOEMPRESA", WS_EMPRESA_PRINC))
    nome_emp    = str(func.get("NOMEEMPRESA", "")).upper().strip()
    situacao    = str(func.get("SITUACAO", "")).upper()

    # Regra: Exceção
    if nome_emp in [e.upper() for e in EMPRESAS_EXCECAO]:
        return False, f"Empresa '{nome_emp}' em lista de exceção"

    # Regra: Já inativo
    if situacao in ("INATIVO", "DEMITIDO"):
        return False, "Funcionário já está inativo"

    # Código SOC
    if not codigo or codigo in ("", "None", "0"):
        return False, "Campo 'CODIGO' ausente — não é possível identificar o funcionário"

    envelope = montar_envelope_soap(codigo, cod_empresa, "INATIVO", DATA_DEMISSAO)

    headers = {
        "Content-Type": "text/xml;charset=UTF-8",
        "SOAPAction":   '""',
    }

    try:
        endpoint = WSDL_URL.replace("?wsdl", "")
        resp = requests.post(
            endpoint,
            data=envelope.encode("utf-8"),
            headers=headers,
            timeout=30,
        )

        logging.debug(f"HTTP {resp.status_code} | {resp.text[:400]}")

        if resp.status_code != 200:
            fault = re.search(r"<faultstring>(.*?)</faultstring>", resp.text, re.S)
            msg = fault.group(1).strip() if fault else f"HTTP {resp.status_code}"
            return False, msg

        # Verifica - Resposta 200
        if "encontrouErro>true" in resp.text or "encontrouErro&gt;true" in resp.text:
            erro = re.search(r"<descricaoErro>(.*?)</descricaoErro>", resp.text, re.S)
            return False, erro.group(1).strip() if erro else "Erro de negócio não especificado"

        if "atualizouFuncionario>true" in resp.text:
            return True, "Inativado com sucesso"

        # Resposta 200 sem erro — sucesso
        return True, "Requisição aceita pelo servidor"

    except requests.exceptions.Timeout:
        return False, "Timeout — servidor demorou mais de 30s"
    except Exception as e:
        return False, f"Erro técnico: {e}"

# Execução
def configurar_log():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        handlers=[
            logging.FileHandler(
                f"inativacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log",
                encoding="utf-8",
            ),
            logging.StreamHandler(),
        ],
    )

def main():
    configurar_log()

    logging.info("=" * 58)
    logging.info("  INATIVAÇÃO DE VIDAS — SOC WebService Modelo 2")
    logging.info(
        f"  Modo: {'TESTE (CPF: ' + CPF_TESTE + ')' if CPF_TESTE else 'PRODUÇÃO'}"
    )
    logging.info(f"  Data de demissão: {DATA_DEMISSAO}")
    logging.info("=" * 58)

    funcionarios = buscar_funcionarios_ativos()
    if not funcionarios:
        logging.error("Nenhum funcionário retornado. Encerrando.")
        return

    logging.info(f"Total a processar: {len(funcionarios)}")

    resultados = []
    n_sucesso = n_ignorado = n_erro = 0
    total = len(funcionarios)

    for i, func in enumerate(funcionarios, 1):
        nome    = func.get("NOME", "Desconhecido")
        cod     = func.get("CODIGO", "?")
        empresa = func.get("NOMEEMPRESA", "?")

        logging.info(f"[{i:>4}/{total}]  {nome}  |  cod:{cod}  |  {empresa}")

        ok, msg = inativar_funcionario(func)

        if ok:
            n_sucesso += 1
            status = "SUCESSO"
            logging.info(f"          ✓ {msg}")
        elif any(k in msg.lower() for k in ("exceção", "já está", "ausente")):
            n_ignorado += 1
            status = "IGNORADO"
            logging.info(f"          → {msg}")
        else:
            n_erro += 1
            status = "ERRO"
            logging.warning(f"          ✗ {msg}")

        resultados.append({
            "Empresa":        empresa,
            "Codigo_SOC":     cod,
            "Matricula":      func.get("MATRICULAFUNCIONARIO", ""),
            "Funcionario":    nome,
            "CPF":            func.get("CPF", ""),
            "Situacao_Antes": func.get("SITUACAO", ""),
            "Status":         status,
            "Detalhe":        msg,
            "Processado_Em":  datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        })

    # Gera relatório
    arq = f"Log_Inativacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    pd.DataFrame(resultados).to_excel(arq, index=False)

    logging.info("")
    logging.info("=" * 58)
    logging.info(f"  Total processado : {total}")
    logging.info(f"  ✓ Inativados     : {n_sucesso}")
    logging.info(f"  → Ignorados      : {n_ignorado}")
    logging.info(f"  ✗ Erros          : {n_erro}")
    logging.info(f"  Relatório        : {arq}")
    logging.info("=" * 58)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.info("Execução interrompida pelo usuário.")
    except Exception as e:
        logging.exception(f"Erro fatal: {e}")
    finally:
        input("\nPressione Enter para fechar...")
