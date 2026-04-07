# Automação de Inativação de Vidas - SOC

![Status](https://img.shields.io/badge/status-estável-brightgreen)
![Stack](https://img.shields.io/badge/stack-Python%203.10+-blue)
![Tipo](https://img.shields.io/badge/tipo-Integração%20SOAP%20%7C%20WebService-orange)

Script Python para automatizar o processo de inativação de funcionários no sistema **SOC** (Software Integrado de Gestão Ocupacional), substituindo o fluxo manual de relatório → planilha → importação Modelo I por uma integração direta via API.

---

## Como funciona

O script realiza dois passos em sequência:

1. **Leitura** - consulta o *Exporta Dados* do SOC e obtém a lista de funcionários ativos em formato JSON.
2. **Gravação** - para cada funcionário, monta um envelope SOAP com autenticação WS-Security (*PasswordDigest*) e envia ao WebService *FuncionarioModelo2Ws*, alterando a situação para `INATIVO` e preenchendo a data de demissão.

Ao final, gera um relatório `.xlsx` com o resultado de cada registro (sucesso, ignorado ou erro).

```
Exporta Dados  ──►  Script Python  ──►  WS FuncionarioModelo2Ws
   (leitura)                                 (gravação)
                                │
                                ▼
                     Log_Inativacao_YYYYMMDD.xlsx
```

---

## Pré-requisitos

- Python 3.10 ou superior
- Acesso ao sistema SOC com permissão de WebService
- IP da máquina liberado no portal SOC para o serviço `FuncionarioModelo2Ws`

### Instalação das dependências

```bash
pip install requests pandas openpyxl zeep
```

---

## Configuração

Abra `inativacao_soc.py` e preencha as variáveis no bloco **CONFIGURAÇÕES**:

### Exporta Dados (leitura)

| Variável | O que preencher |
|---|---|
| `ED_EMPRESA` | Código da empresa principal no SOC |
| `ED_CODIGO` | Código do Exporta Dados (tela 733 do SOC) |
| `ED_CHAVE` | Chave gerada na tela 733 do SOC |

### WebService (gravação)

| Variável | O que preencher |
|---|---|
| `WS_USUARIO` | Login de usuário do SOC Integração (ex: `U1234567`) |
| `WS_SENHA` | Senha do usuário no SOC Integração |
| `WS_CHAVE_ACESSO` | Chave de acesso do WebService - gere em: SOC › Pessoas - Usuário (189) › Usuário Integração › Inclui ou Consulta Usuário › Configuração de Integração do Usuário Webservice  |
| `WS_RESPONSAVEL` | Código do responsável cadastrado na configuração do WebService - tela 337 › Configuração de Integração |
| `WS_EMPRESA_PRINC` | Mesmo valor de `ED_EMPRESA` |

### Regras de negócio

| Variável | O que preencher |
|---|---|
| `EMPRESAS_EXCECAO` | Lista com os nomes **exatos** das empresas que **não** devem ter funcionários inativados (copie o nome como aparece no SOC) |
| `CPF_TESTE` | CPF de um funcionário para teste unitário. Deixe `None` para rodar em produção |
| `DATA_DEMISSAO` | Data de demissão a gravar (padrão: data atual) |

---

## Como executar

### 1. Teste com um funcionário

Antes de rodar em produção, valide com um único registro:

```python
# Em inativacao_soc.py, defina:
CPF_TESTE = "12345678900"  # CPF do funcionário de teste (somente números)
```

```bash
python inativacao_soc.py
```

Verifique no SOC se o funcionário foi inativado corretamente e se a data de demissão foi preenchida.

### 2. Execução em produção

Após validar o teste:

```python
# Em inativacao_soc.py, defina:
CPF_TESTE = None
```

```bash
python inativacao_soc.py
```

---

## Saída gerada

Ao final de cada execução são gerados dois arquivos na pasta do script:

| Arquivo | Conteúdo |
|---|---|
| `inativacao_YYYYMMDD_HHMMSS.log` | Log completo de cada etapa da execução |
| `Log_Inativacao_YYYYMMDD_HHMMSS.xlsx` | Relatório com status de cada funcionário |

### Colunas do relatório Excel

`Empresa` · `Codigo_SOC` · `Matricula` · `Funcionario` · `CPF` · `Situacao_Antes` · `Status` · `Detalhe` · `Processado_Em`

### Status possíveis

| Status | Significado |
|---|---|
| `SUCESSO` | Funcionário inativado com sucesso no SOC |
| `IGNORADO` | Empresa em lista de exceção, funcionário já inativo, ou código ausente |
| `ERRO` | Falha na comunicação com o WebService ou erro de negócio retornado pelo SOC |

---

## Detalhes técnicos

### Autenticação WS-Security PasswordDigest

O SOC exige o padrão **WS-Security UsernameToken Profile 1.0** com `PasswordDigest`. O script calcula o digest conforme a especificação:

```
PasswordDigest = Base64( SHA-1( Nonce_bytes + Created_utf8 + Password_utf8 ) )
```

Um `Nonce` aleatório de 16 bytes e um `Timestamp` com janela de 5 minutos são gerados a cada requisição. O script usa `requests.post` diretamente (sem delegar ao Zeep) para ter controle total sobre o envelope.

### Identificação do funcionário

O WebService identifica o funcionário pelo campo `CODIGO` (código interno SOC), usando `chaveProcuraFuncionario=CODIGO`. Esse campo é retornado diretamente pelo Exporta Dados - Cadastro de Funcionários por situação (FOL)

### Campos enviados na inativação

```xml
<funcionarioWsVo>
  <chaveProcuraFuncionario>CODIGO</chaveProcuraFuncionario>
  <codigo>{CODIGO_SOC}</codigo>
  <tipoBuscaEmpresa>CODIGO_SOC</tipoBuscaEmpresa>
  <codigoEmpresa>{CODIGO_EMPRESA}</codigoEmpresa>
  <situacao>INATIVO</situacao>
  <dataDemissao>{DATA_DEMISSAO}</dataDemissao>
</funcionarioWsVo>
```

---

## Segurança

> **Nunca commite credenciais no repositório.**

As variáveis de credenciais (`WS_USUARIO`, `WS_SENHA`, `WS_CHAVE_ACESSO`, etc.) estão declaradas diretamente no script para facilidade de uso. Para ambientes compartilhados ou pipelines CI/CD, mova-as para variáveis de ambiente ou um arquivo `.env` (adicionado ao `.gitignore`):

```bash
# .env (não versionar)
WS_USUARIO=U1234567
WS_SENHA=sua_senha
WS_CHAVE_ACESSO=sua_chave
```

```python
# No script, substitua as variáveis por:
import os
WS_USUARIO      = os.environ["WS_USUARIO"]
WS_SENHA        = os.environ["WS_SENHA"]
WS_CHAVE_ACESSO = os.environ["WS_CHAVE_ACESSO"]
```

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

WSDL disponível em:
```
https://ws1.soc.com.br/WSSoc/FuncionarioModelo2Ws?wsdl
```

---

## 👤 Autor

**Vinícius Medeiros**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/viniciusrmedeiros)
[![GitHub](https://img.shields.io/badge/GitHub-181717?style=flat&logo=github&logoColor=white)](https://github.com/vini-medeiros-cmd)
