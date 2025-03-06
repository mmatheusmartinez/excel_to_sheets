# **Integração entre Excel e Google Sheets via API**

## **Descrição do Projeto**

Este projeto permite a sincronização automática de planilhas do Excel com o Google Sheets, possibilitando atualizações dinâmicas através de scripts em Python. A solução é ideal para automação de processos e compartilhamento de dados em tempo real entre plataformas locais e a nuvem.

A integração com o Google Sheets utilizando código próprio é pouco documentada e muitas vezes complexa. Este projeto simplifica esse processo ao fornecer uma abordagem estruturada para conectar e manipular planilhas do Google diretamente a partir do Excel.

## **Tecnologias Utilizadas**

- **Python**: Linguagem principal para manipulação e integração de dados.
- **Google Sheets API**: Interface de comunicação para interagir com planilhas na nuvem.
- **gspread**: Biblioteca Python para conexão e manipulação do Google Sheets.
- **pandas**: Para processamento e análise de dados.
- **OpenPyXL**: Biblioteca para manipulação de arquivos Excel.

## **Requisitos**

Para utilizar o projeto, é necessário configurar uma chave de API no **Google Cloud Platform (GCP)** e gerar credenciais para acesso à API do Google Sheets. O processo inclui:

1. Criar um projeto no Google Cloud.
2. Ativar a API do Google Sheets.
3. Criar credenciais de acesso (OAuth 2.0 ou chave de serviço).
4. Compartilhar a planilha com o e-mail do serviço gerado.

## **Como Funciona**

1. O script extrai dados de uma planilha do Excel.
2. Conecta-se ao Google Sheets via API.
3. Atualiza ou insere novos dados na planilha online.
4. Permite sincronização programada para manter os dados atualizados.

Este projeto é essencial para automação e compartilhamento de informações entre equipes, garantindo maior agilidade e precisão na gestão de dados.
