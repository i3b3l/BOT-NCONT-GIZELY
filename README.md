# Robô SIGEF — Regularização de Convênios

Automação de etapas de **certificação, liquidação de despesas e preparação de pagamentos** no SIGEF, utilizando uma planilha Excel como base de trabalho.

Desenvolvido e adaptado por **Izabel Galvão**.

## Componentes

- `Certificar_Liquidar_Convenios.py` — valida os dados e a estrutura da planilha, processa certificações e liquidações pendentes e registra os identificadores de CE e NL.
- `Preparar_Pagamento_Convenios.py` — consulta notas de liquidação pendentes, confere o favorecido e o domicílio bancário e registra o número da preparação de pagamento.
- `Iniciar_Certificar_Liquidar_Convenios.bat` — iniciador do fluxo de certificação e liquidação.
- `Iniciar_Preparar_Pagamento_Convenios.bat` — iniciador do fluxo de preparação de pagamentos.
- `Depurador do Google.bat` — arquivo auxiliar incluído no pacote.
- `LEIA-ME.txt` — manual detalhado de instalação, execução, recuperação e solução de problemas.

## Requisitos

- Windows
- Python 3.10 ou superior
- Google Chrome
- Bibliotecas Python: `openpyxl`, `playwright`, `keyboard` e `pyautogui`
- Acesso autorizado ao SIGEF
- Planilha operacional `Pagamentos_Tratado.xlsx`, mantida na mesma pasta dos scripts durante a execução

Instale as bibliotecas com:

```powershell
py -3 -m pip install --upgrade pip
py -3 -m pip install openpyxl playwright keyboard pyautogui
```

Se o comando `py` não estiver disponível, tente substituir `py -3` por `python`.

## Como executar

1. Mantenha os arquivos do pacote na mesma pasta.
2. Prepare a planilha `Pagamentos_Tratado.xlsx` conforme o manual `LEIA-ME.txt`.
3. Execute **um único robô por vez**, abrindo o `.bat` correspondente.
4. Faça manualmente o login no SIGEF e resolva o CAPTCHA, se solicitado.
5. Acompanhe o terminal e confira os resultados no SIGEF e na planilha.

## Cuidados importantes

- O pacote não realiza login nem resolve CAPTCHA automaticamente.
- Feche a planilha no Excel antes de iniciar o robô e mantenha-a fechada durante a execução.
- Não repita automaticamente operações interrompidas ou com status `PESQUISAR_CE`, `PESQUISAR_NL` ou `PESQUISAR_PP`. Confira primeiro o documento no SIGEF.
- Preserve os arquivos de recuperação enquanto houver uma recuperação pendente.
- Não compartilhe planilhas com dados pessoais ou financeiros sem autorização.
- Utilize somente o endereço oficial do SIGEF e credenciais próprias autorizadas.

Consulte `LEIA-ME.txt` para instruções completas e procedimentos de recuperação.

## Observação sobre dados

A planilha real de pagamentos não está incluída neste repositório. Ela deve ser disponibilizada localmente apenas a usuários autorizados e protegida de acordo com as regras aplicáveis aos dados tratados.
