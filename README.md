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

Utilizando a biblioteca do Yahoo Finance podemos obter os dados de fechamento ajustados para o Ibovespa e para o dólar.
```
df = yf.download(['^BVSP','BRL=X'],datetime.datetime.now() - datetime.timedelta(days = 365), datetime.datetime.now())['Adj Close']
df
```

E o dataframe com os dados de fechamento ajustado:
<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/217103097-2f1bfe97-dc19-4c5e-94c1-c25880c283d3.png" />
  <p> Figura 4 - DataFrame fechamento ajustado dólar e ibovespa . </p>
</div>

Para o melhor entendimento do dataframe alterou-se o nome das colunas para *dolar* e *ibovespa*
```
df.columns = ['dolar','ibovespa']
df.head(50)
```
A linha *df.head(50)* apresenta as 50 primeiras linhas do nosso dataframe, é possível observar a presença de valores *NaNs*
<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/217107921-d7591638-7997-4191-9e79-4052fa9526bb.png" />
  <p> Figura 5 - DataFrame com valores NaNs. </p>
</div>

Para eliminar linhas com valores *NaNs*:

```
df = df.dropna()
```

Agora temos todos os dados corretos. 
O próximo passo é obter os dados mensais e anuais. Podemos obter com as seguintes linhas:

```
df_anual = df.resample('Y').last()
df_mensal = df.resample('M').last()
```

E assim podemos calcular os retornos diários, mensais e anuais do ibovespa e dólar:
```
retorno_anual = df_anual.pct_change().dropna()
retorno_mensal = df_mensal.pct_change().dropna()
retorno_diario = df.pct_change().dropna()
```

Pegando a ultima linha com os valores de retornos obtidos e salvando em variáveis:
```
retorno_diario_dolar = retorno_diario.iloc[-1,0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]

retorno_mensal_dolar = retorno_mensal.iloc[-1,0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]

retorno_anual_dolar = retorno_anual.iloc[-1,0]
retorno_anual_ibov = retorno_anual.iloc[-1, 1]
```

Transformando os dados em porcentagem, e arredondado com duas casas decimais para melhor visualização posteriormente:
```
etorno_diario_dolar = round(retorno_diario_dolar * 100, 2)
retorno_diario_ibov = round(retorno_diario_ibov * 100, 2)

retorno_mensal_dolar = round(retorno_mensal_dolar * 100, 2)
retorno_mensal_ibov = round(retorno_mensal_ibov * 100, 2)

retorno_anual_dolar = round(retorno_anual_dolar * 100, 2)
retorno_anual_ibov = round(retorno_anual_ibov * 100, 2)
```

Plotando o gráfico para retorno do ibovespa no range de um ano:
```
plt.style.use('cyberpunk')

df.plot(y = 'ibovespa', use_index = True, legend = False)

plt.title('Ibovespa')

plt.savefig('ibovespa.png', dpi = 300)

plt.show()
```

<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/217109087-8190502f-c590-41df-954d-d287f6fc4f11.png" />
  <p> Figura 6 - Gráfico Ibovespa. </p>
</div>

E para o gráfico de retorno do dólar no range de um ano:
```
plt.style.use('cyberpunk')

df.plot(y = 'dolar', use_index = True, legend = False)

plt.title('Dolar')

plt.savefig('dolar.png', dpi = 300)

plt.show()
```

<div align="center">
  <img src="https://user-images.githubusercontent.com/82683162/217109299-45f41b72-f306-4534-8ec3-354bdc1afd81.png" />
  <p> Figura 7 - Gráfico Dólar. </p>
</div>

Agora precisamos conectar com o outlook para automatizar o envio de email:

```
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
```

E podemos automatizar o email:
```
email.To = 'leandro.madureira@aluno.ifsp.edu.br; brenno@varos.com.br'
email.Subject = 'Relatório Diário'
email.body = f'''Prezado diretor, segue o relatório diário:

Bolsa:

No ano o ibovespa está tendo uma rentabilidade de {retorno_anual_ibov}%,
enquanto no mês a rentabilidade é de {retorno_mensal_ibov}%.

No último dia útil, o fechamento do Ibovespa foi de {retorno_diario_ibov}%.

Dólar:

No ano o Dólar está tendo uma rentabilidade de {retorno_anual_dolar}%,
enquanto no mês a rentabilidade é de {retorno_mensal_dolar}%.

No último dia útil, o fechamento do Dólar foi de {retorno_diario_dolar}%.

Att,
Leandro Madureira
'''

anexo_ibovespa = r'C:\Users\pichau\Relatorio_Fechamento_Mercado\ibovespa.png'
anexo_dolar = r'C:\Users\pichau\Relatorio_Fechamento_Mercado\dolar.png'

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_dolar)

email.Send()
```

`email.To` podemos fornecer uma lista de emails dento das aspas, seprando o email por ";".
`email.Subject` é onde colocamos o assunto do email.
`email.Body` escrevemos o corpo do email.
`email.Attachments.Add(arquivo)` anexa o arquivo passando o local.
`email.Send()` envia o email.

