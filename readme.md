# Automação de Inativação de Vidas - SOC

![Status](https://img.shields.io/badge/status-estável-brightgreen)
![Versão](https://img.shields.io/badge/versão-1.6.0-blue)
![Stack](https://img.shields.io/badge/stack-Python%203.10+-blue)
![Tipo](https://img.shields.io/badge/tipo-Integração%20SOAP%20%7C%20WebService-orange)

Script Python para automatizar a inativação em massa de funcionários no sistema SOC (Software Integrado de Gestão Ocupacional), iterando por todas as empresas ativas e utilizando o WebService SOAP com autenticação WS-Security PasswordDigest.

---

## Como funciona

```
[ED - Cadastro de Empresas] - CSV -► Lista de empresas ativas
                    ↓
    Para cada empresa (fora da lista de exceção):
                    ↓
    [ED - Cadastro de Funcionários por situação (FOL)] - JSON -► Funcionários ativos
                    ↓
                    ↓
    Para cada funcionário ativo:
                    ↓
    [WebService Modelo 2 - FuncionarioModelo2Ws] - SOAP -► INATIVO
                    ↓
                    ↓
      Log_Inativacao_YYYYMMDD_HHMMSS.xlsx
```

### Passos em sequência

1. **Busca empresas** - consulta ED Cadastro de Empresas (CSV) e lista todas as empresas com `ATIVO = 1`
2. **Filtra exceções** - ignora empresas da lista `EMPRESAS_EXCECAO`
3. **Busca funcionários** - para cada empresa, consulta ED Cadastro de Funcionários por situação (FOL) passando o **código da empresa cliente** no parâmetro `empresa`
4. **Inativa** - envia envelope SOAP com WS-Security PasswordDigest para cada funcionário ativo
5. **Relatório** - gera `.xlsx` com status de cada registro

---

## Pré-requisitos

- Python 3.10 ou superior
- Acesso ao sistema SOC com permissão de WebService.
- IP da máquina liberado no portal SOC para o serviço `FuncionarioModelo2Ws`

### Instalação das dependências

```bash
pip install requests pandas openpyxl python-dotenv
```

---

## Configuração

As variáveis podem ser definidas via arquivo `.env` ou diretamente no script.

### Arquivo `.env`

```env
# URL base do Exporta Dados
ED_URL=https://ws1.soc.com.br/WebSoc/exportadados

# Empresa principal do SOC
ED_EMPRESA_PRINCIPAL= Código da empresa principal no SOC.

# Exporta Dados Cadastro de Funcionários por situação (FOL).
ED_FUNC_CODIGO= Código do Exporta Dados (tela 733 do SOC).
ED_FUNC_CHAVE= Chave gerada na tela 733 do SOC.

# Exporta Dados Cadastro de Empresas
ED_EMPRESAS_CODIGO= Código do Exporta Dados (tela 733 do SOC).
ED_EMPRESAS_CHAVE= Chave gerada na tela 733 do SOC.

# WebService SOAP
WSDL_URL=https://ws1.soc.com.br/WSSoc/FuncionarioModelo2Ws?wsdl
WS_USUARIO= Login de usuário do SOC Integração (ex: U1234567)
WS_SENHA= Chave de acesso do WebService - gere em: SOC › Pessoas - Usuário (189) › Usuário Integração › Inclui ou Consulta Usuário › Configuração de Integração do Usuário Webservice.
WS_CHAVE_ACESSO= WS_SENHA
WS_RESPONSAVEL= Código do responsável cadastrado na configuração do WebService - tela 337 › Configuração de Integração.
WS_EMPRESA_PRINC= Código empresa principal.

# Modo de execução
SOC_DRY_RUN=true        # true = simulação | false = execução real
SOC_CPF_TESTE=          # vazio = todos | CPF = teste unitário
```

---

## Como executar

### 1. Simulação completa (DRY-RUN)

Processa todas as empresas e funcionários **sem alterar nada no SOC**. Gera o relatório Excel normalmente.

```bash
# .env: SOC_DRY_RUN=true
python inativacao_soc.py
```

Verifique o log - confirme o total de empresas, funcionários encontrados e quais seriam ignorados.

### 2. Teste com um CPF específico

```bash
# .env: SOC_CPF_TESTE=12345678900 | SOC_DRY_RUN=false
python inativacao_soc.py
```

Valide no SOC se o funcionário foi inativado corretamente.

### 3. Execução em produção

```bash
# .env: SOC_DRY_RUN=false | SOC_CPF_TESTE= (vazio)
python inativacao_soc.py
```

---

## Saída gerada

| Arquivo | Conteúdo |
|---|---|
| `inativacao_YYYYMMDD_HHMMSS.log` | Log completo com cada etapa |
| `Log_Inativacao_YYYYMMDD_HHMMSS.xlsx` | Relatório com status por funcionário |

### Colunas do relatório Excel

`Empresa` · `Codigo_Empresa` · `Codigo_SOC` · `Matricula` · `Funcionario` · `CPF` · `Situacao_Antes` · `Status` · `Detalhe` · `Processado_Em` · `Timestamp_Envio` · `Ambiente` · `Versao_Script`

### Status possíveis

| Status | Significado |
|---|---|
| `SUCESSO` | Funcionário inativado com sucesso |
| `IGNORADO` | Empresa em exceção, funcionário já inativo, código ausente ou DRY-RUN |
| `ERRO` | Falha na comunicação ou erro de negócio retornado pelo SOC |

---

## Detalhes técnicos

### Por que CSV para o ED Cadastro de Empresas?

O Exporta Dados Cadastro de Empresas **não suporta retorno JSON** (confirmado na tela 733 do SOC). O script usa `tipoSaida=csv` com `csv.DictReader` e separador `;`.

### Por que o código da empresa cliente no parâmetro `empresa`?

O ED Cadastro de Funcionários por situação (FOL) filtra funcionários pelo código informado em `empresa`. Passar apenas a empresa principal retorna `[]`. É necessário passar o código de cada empresa cliente individualmente - esse é o fix central desta versão.

### Autenticação WS-Security PasswordDigest

```
PasswordDigest = Base64( SHA-1( Nonce_bytes + Created_utf8 + Password_utf8 ) )
```

Nonce aleatório de 16 bytes e Timestamp com janela de 5 minutos gerados a cada requisição.

### Validação de CPF

Validação completa com dígitos verificadores - rejeita CPFs com formato inválido, sequências repetidas ou dígitos incorretos.

---

## Lista de exceção

Empresas em `EMPRESAS_EXCECAO` não terão funcionários processados. A verificação é feita em dois momentos: durante a iteração de empresas e novamente dentro de `inativar_funcionario()` como dupla checagem.

Para adicionar ou remover empresas, edite a lista diretamente no script.

---

## Recomendações de Segurança

> **Nunca commite credenciais no repositório.**

Use o arquivo `.env` para todas as credenciais e confirme que ele está no `.gitignore`.

### `.gitignore` recomendado

```
.env
*.log
*.xlsx
__pycache__/
```

---

## Endpoint do WebService

```
https://ws1.soc.com.br/WSSoc/FuncionarioModelo2Ws
```

WSDL:
```
https://ws1.soc.com.br/WSSoc/FuncionarioModelo2Ws?wsdl
```

---

## 👤 Autor

**Vinícius Medeiros**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/viniciusrmedeiros)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=flat&logo=github&logoColor=white)](https://github.com/vini-medeiros-cmd)
