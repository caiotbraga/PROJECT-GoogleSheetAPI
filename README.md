# Google Sheets API Integration
Este projeto tem como objetivo ler uma planilha no Google Sheets e fazer alterações com base nos dados encontrados nela. Nesse caso está sendo calculado
o status

## Requisitos
Antes de utilizar este código, siga os passos abaixo para configurar suas credenciais e ativar a API do Google Sheets:

- **Passo 1**: Acesse Google Cloud Console, Link: https://www.google.com/aclk?sa=l&ai=DChcSEwjp38z90byEAxU2YUgAHd9mB3UYABAAGgJjZQ&gclid=EAIaIQobChMI6d_M_dG8hAMVNmFIAB3fZgd1EAAYASAAEgLBg_D_BwE&sig=AOD64_00JYK1PVOod-wWc5tqJ_uYVTyoag&q&adurl&ved=2ahUKEwix3Mb90byEAxWHppUCHacwBWgQ0Qx6BAgFEAE
- **Passo 2**: Procure por **"Google Sheets API"** e ative-a para o seu projeto.
- **Passo 3**: Procure por **"Credenciais"** e crie ou selecione uma credencial de serviço. Faça o download do arquivo JSON contendo suas credenciais.

## Configuração
- Clone este repositório em sua máquina local.
- Certifique-se de ter o SDK .NET Core instalado em sua máquina.
- Adicione o arquivo JSON de credenciais baixado do Console do Google Cloud ao diretório Credentials do projeto.
- No código, **atualize o valor da variável credPath** para apontar para o caminho do arquivo de credenciais em sua máquina local.
- Atualize o valor da variável spreadsheetId para o ID da planilha do Google Sheets que deseja modificar.

## Como usar
Após configurar suas credenciais e ajustar as variáveis no código, você pode executar o projeto. Ele lerá a planilha especificada, calculará as notas dos alunos e atualizará a planilha com os resultados.

## Observações
Certifique-se de que a planilha do Google Sheets tenha a estrutura correta e que os intervalos especificados no código correspondam aos dados corretos na planilha.
Link para a planilha utilizada no projeto: https://docs.google.com/spreadsheets/d/1-iz4w7WxPP5Dj6I-UjL5t1Y9WDrXqgg_g7B2csJVRbo/edit#gid=0


