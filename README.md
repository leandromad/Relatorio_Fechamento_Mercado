# Relatorio de Fechamento de Mercado com Python
Criando um relatório de fechamento de mercado por e-mail utilizando Pyhton. AULA 1 do BOTCAMP VAROS - Programação

## Notas
O arquivo `requirements.txt`  contem todas as bibliotecas Python que foram utilizadas no notebook, para instalar essas bibliotecas basta abrir o terminal e indicar a pasta do projeto, e então no terminal inserir o comando:
```
pip install -r requirements.txt
```

### O Projeto
Este projeto tem como objetivo final gerar um relatório como o da Figura 1, que pode ser enviado diariamente por email com os resultados de fechamento do ibovespa e do dolar, dados obtidos através do Yahoo Finance.

<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/216992220-71b213af-bf74-4277-b644-7abe47c6b835.png" />
  <p> Figura 1 - Relatório Email. </p>
</div>

Junto do email, também é anexado dois arquivos onde mostra um gráfico com o histórico do dolar e do ibovespa durante 1 ano.

<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/216992942-95bad51e-b7e1-4162-bcee-d242d28d8c11.png" />
  <p> Figura 2 - Histórico Ibovespa. </p>
</div>

<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/216993134-e00226e0-7b15-4005-9c74-02b7643caf05.png" />
  <p> Figura 3 - Histórico Dólar. </p>
</div>

### Passo a Passo
Inicialmente deve-se importar as bibliotecas que serão utilizadas no projeto.
```
import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import win32com.client as win32
```
