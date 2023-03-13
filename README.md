# Automatização de envio de e-mail com Python

O Python é uma linguagem de programação amplamente usada em aplicações da Web, desenvolvimento de software, ciência de dados, machine learning (ML), automações, entre outras coisas.

Nesse projeto, foi feita uma **análise de dados** a partir de uma base de vendas com mais de 100.000 linhas para que fosse possível enviar um **relatório de vendas** por e-mail de forma automatizada.

<br>

## Pré-requisitos

Para implementar os códigos feitos no projeto, é necessário ter instalada uma versão do Python. Além disso, nesse projeto foi utilizada a versão gratuita do Pycharm.

<img src="https://img.shields.io/static/v1?label=&message=Python&color=9118c1&style=for-the-badge&logo=ghost"/>
<img src="https://img.shields.io/static/v1?label=&message=PyCharm&color=656c1&style=for-the-badge&logo=ghost"/>
<img src="https://img.shields.io/static/v1?label=&message=Outlook&color=9118c1&style=for-the-badge&logo=ghost"/>

<br>

## Link para download do PyCharm:
```
https://www.jetbrains.com/pycharm/download/#section=windows
```

Com a IDE instalada corretamente, é necessário instalar a biblioteca Pandas e a Openpyxl que lerão nossa base de dados:

```
pip install pandas
```
```
pip install openpyxl
```

<br>

## Configuração Importante no PowerShell

Para quem utiliza Windows, pode ser necessário habilitar a execução de scripts para que o PyCharm funcione corretamente.

Assim, para habilitar a conexão com o ambiente virtual do PyCharm no PowerShell, digite:

```
Get-ExecutionPolicy -List
```
```
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned 
```

No terminal do PyCharm, ativar a conexão com VENV:
```
venv\Scripts\activate 
```

Para desabilitar a conexão:
```
deactivate
```

## ⚙️ Executando 

Depois das instalações necessárias, o código disponível nesse repositório já pode ser executado. O Pyhton enviará o e-mail com as informações abaixo em formato de tabela: 
- Faturamento por loja;
- Quantidade de produtos vendidos por loja;
- Ticket médio (faturamento dividido pela quantidadee de produtos vendidos);

<br>

### Onde encontrar o conteúdo desse projeto?
<strong>Hashtag Treinamentos:</strong> <br>
https://www.hashtagtreinamentos.com/hashtag-free?origemurl=hashtag_yt_org_vid_ivlA5b2sC9g


---
