#!/usr/bin/env python
# coding: utf-8

# # Automação de Indicadores
# 
# ### Objetivo: Treinar e criar um Projeto Completo que envolva a automatização de um processo feito no computador
# 
# ### Descrição:
# 
# Imagine que você trabalha em uma grande rede de lojas de roupa com 25 lojas espalhadas por todo o Brasil.
# 
# Todo dia, pela manhã, a equipe de análise de dados calcula os chamados One Pages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.
# 
# Um One Page é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir em 1 página (daí o nome OnePage) tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.
# 
# Exemplo de OnePage:

# ![title](onepage.png)

# O seu papel, como Analista de Dados, é conseguir criar um processo da forma mais automática possível para calcular o OnePage de cada loja e enviar um email para o gerente de cada loja com o seu OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo.
# 
# Ex: O e-mail a ser enviado para o Gerente da Loja A deve ser como exemplo

# ![exemplo_email](Exemplo.JPG)

# ### Arquivos e Informações Importantes
# 
# - Arquivo Emails.xlsx com o nome, a loja e o e-mail de cada gerente. Obs: Sugiro substituir a coluna de e-mail de cada gerente por um e-mail seu, para você poder testar o resultado
# 
# - Arquivo Vendas.xlsx com as vendas de todas as lojas. Obs: Cada gerente só deve receber o OnePage e um arquivo em excel em anexo com as vendas da sua loja. As informações de outra loja não devem ser enviados ao gerente que não é daquela loja.
# 
# - Arquivo Lojas.csv com o nome de cada Loja
# 
# - Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (informações também estão no arquivo Emails.xlsx) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. Além disso, no corpo do e-mail, deve ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento da loja.
# 
# - As planilhas de cada loja devem ser salvas dentro da pasta da loja com a data da planilha, a fim de criar um histórico de backup
# 
# ### Indicadores do OnePage
# 
# - Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
# - Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
# - Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500
# 
# Obs: Cada indicador deve ser calculado no dia e no ano. O indicador do dia deve ser o do último dia disponível na planilha de Vendas (a data mais recente)
# 
# Obs2: Dica para o caracter do sinal verde e vermelho: pegue o caracter desse site (https://fsymbols.com/keyboard/windows/alt-codes/list/) e formate com html

# Passo 1: Importar arquivos e bibliotecas 

# In[1]:


#importar bibliotecas
import pandas as pd 
import pathlib
import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# In[2]:


#importar base de dados 
emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding = 'latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')
display(emails)
display(lojas)
display(vendas)


# Passo 2 - Definir Criar uma tabela para cada loja e Definir o dia do indicador 

# In[3]:


#incluir nome da loja em vendas 
vendas = vendas.merge(lojas, on='ID Loja')


# In[4]:


dic_loja = {}
for loja in lojas['Loja']:
    dic_loja[loja] = vendas.loc[vendas['Loja'] == loja, :]
    
    


# In[5]:


dia_indicador = vendas['Data'].max()
print(dia_indicador)


# Passo 3 - Salvar a Planilha na pasta de backup

# In[6]:


#identificar se a pasta já existe
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()
lista_nome_backup = []
for arquivo in arquivos_pasta_backup:
    lista_nome_backup.append(arquivo.name)

for loja in dic_loja:
    if loja not in lista_nome_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    #salvar dentro da pasta
    nome_arquivo = (f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx')
    local_arquivo = caminho_backup / loja / nome_arquivo

    dic_loja[loja].to_excel(local_arquivo)


# Passo 4 - Calcular o indicador pra uma loja

# In[7]:


#definição de metas
meta_faturamento_dia = 1000
metafaturamentoano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500


# In[8]:


for loja1 in dic_loja:


    vendas_loja = dic_loja[loja1]
    vendas_loja1_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]
    
    
    
    #faturamento 
    faturamento_ano = vendas_loja['Valor Final'].sum(numeric_only=True)
    faturamento_dia = vendas_loja1_dia['Valor Final'].sum(numeric_only=True)
    #print(faturamento_dia)
    #print(faturamento_ano)
    #diversidade de produtos
    qtde_produto_ano = len(vendas_loja['Produto'].unique())
    qtde_produto_dia = len(vendas_loja1_dia['Produto'].unique())
    #print(qtde_produto_ano)
    #print(qtde_produto_dia)
    
    
    
    #ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True) #groupby agrupa pelos valores que são iguais de acordo com a coluna que passar pra ele
    ticket_medio_ano = valor_venda['Valor Final'].mean(numeric_only=True)
    #print(ticket_medio_ano)
    valor_venda_dia = vendas_loja1_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean(numeric_only=True)
    #print(ticket_medio_dia)
    #enviar email
    nome = emails.loc[emails['Loja']==loja1, 'Gerente'].values[0]
    fromaddr = "caio.toscano345@gmail.com"
    toaddr = emails.loc[emails['Loja']==loja1, 'E-mail'].values[0]
    msg = MIMEMultipart()
    
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja1}'
    
    
    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= metafaturamentoano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produto_dia >= meta_qtde_produtos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produto_ano >= meta_qtde_produtos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticket_medio_dia:
        cor_ticket_dia = 'green'
    else: cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticket_medio_ano:
        cor_ticket_ano = 'green'
    else: cor_ticket_ano = 'red'
    
    body = f'''<p> Bom Dia, {nome}</p> 
    
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja1}</strong> foi: </p>
    
    <html>
    <head>
    
    </head>
    <body>
    
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td> R$ {faturamento_dia:.2f}</td>
        <td>R$ {meta_faturamento_dia:.2f}</td>
        <td><font color = "{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td>{qtde_produto_dia}</td>
        <td>{meta_qtde_produtos_dia}</td>
        <td><font color = "{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td>R$ {ticket_medio_dia:.2f}</td>
        <td>R$ {meta_ticket_medio_dia:.2f}</td>
        <td><font color = "{cor_ticket_dia}">◙</font></td>
      </tr>
      
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td>R$ {faturamento_ano:.2f}</td>
        <td>R$ {metafaturamentoano:.2f}</td>
        <td><font color = "{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td>{qtde_produto_ano}</td>
        <td>{meta_qtde_produtos_ano}</td>
        <td><font color = "{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td>R$ {ticket_medio_ano:.2f}</td>
        <td>R$ {meta_ticket_medio_ano:.2f}</td>
        <td><font color = "{cor_ticket_ano}">◙</font></td>
      
    </table>
    
    </body>
    </html>
    </body>
    </html>
    
    
    <p>Segue em anexo planilha de dados para mais detalhes.</p>
    
    <p>Qualquer dúvida estou a disposição</p>
    <p>Atenciosamente, Caio Toscano</p>
    '''
    #msg.
    #Aqui começa a parte do Anexo
    
    msg.attach(MIMEText(body, 'html')) #mudando o formato do corpo para html
    filename = pathlib.Path.cwd() / caminho_backup / loja1 / f'{dia_indicador.month}_{dia_indicador.day}_{loja1}.xlsx'
    attachment = open(str(filename),'rb') #filename precisa ser convertido em str
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)
    attachment.close()
    
    #Aqui termina a parte do Anexo
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "tnox bfyx mpdd pzuf")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()
    
    print('\nEmail enviado com sucesso!')
    print('E-mail da Loja {} foi enviado'.format(loja1))




# Passo 5 - Enviar E-mails para os gerentes

# In[9]:


nome = emails.loc[emails['Loja']==loja1, 'Gerente'].values[0]
fromaddr = "caio.toscano345@gmail.com"
toaddr = emails.loc[emails['Loja']==loja1, 'E-mail'].values[0]
msg = MIMEMultipart()

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja1}'


if faturamento_dia >= meta_faturamento_dia:
    cor_fat_dia = 'green'
else:
    cor_fat_dia = 'red'
if faturamento_ano >= metafaturamentoano:
    cor_fat_ano = 'green'
else:
    cor_fat_ano = 'red'
if qtde_produto_dia >= meta_qtde_produtos_dia:
    cor_qtde_dia = 'green'
else:
    cor_qtde_dia = 'red'
if qtde_produto_ano >= meta_qtde_produtos_ano:
    cor_qtde_ano = 'green'
else:
    cor_qtde_ano = 'red'
if ticket_medio_dia >= meta_ticket_medio_dia:
    cor_ticket_dia = 'green'
else: cor_ticket_dia = 'red'
if ticket_medio_ano >= meta_ticket_medio_ano:
    cor_ticket_ano = 'green'
else: cor_ticket_ano = 'red'

body = f'''<p> Bom Dia, {nome}</p> 

<p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja1}</strong> foi: </p>

<html>
<head>

</head>
<body>

<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Dia</th>
    <th>Meta Dia</th>
    <th>Cenário Dia</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td> R$ {faturamento_dia:.2f}</td>
    <td>R$ {meta_faturamento_dia:.2f}</td>
    <td><font color = "{cor_fat_dia}">◙</font></td>
  </tr>
  <tr>
    <td>Diversidade de Produtos</td>
    <td>{qtde_produto_dia}</td>
    <td>{meta_qtde_produtos_dia}</td>
    <td><font color = "{cor_qtde_dia}">◙</font></td>
  </tr>
  <tr>
    <td>Ticket Médio</td>
    <td>R$ {ticket_medio_dia:.2f}</td>
    <td>R$ {meta_ticket_medio_dia:.2f}</td>
    <td><font color = "{cor_ticket_dia}">◙</font></td>
  </tr>
  
</table>
<br>
<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Ano</th>
    <th>Meta Ano</th>
    <th>Cenário Ano</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td>R$ {faturamento_ano:.2f}</td>
    <td>R$ {metafaturamentoano:.2f}</td>
    <td><font color = "{cor_fat_ano}">◙</font></td>
  </tr>
  <tr>
    <td>Diversidade de Produtos</td>
    <td>{qtde_produto_ano}</td>
    <td>{meta_qtde_produtos_ano}</td>
    <td><font color = "{cor_qtde_ano}">◙</font></td>
  </tr>
  <tr>
    <td>Ticket Médio</td>
    <td>R$ {ticket_medio_ano:.2f}</td>
    <td>R$ {meta_ticket_medio_ano:.2f}</td>
    <td><font color = "{cor_ticket_ano}">◙</font></td>
  
</table>

</body>
</html>
</body>
</html>


<p>Segue em anexo planilha de dados para mais detalhes.</p>

<p>Qualquer dúvida estou a disposição</p>
<p>Atenciosamente, Caio Toscano</p>
'''
#msg.
#Aqui começa a parte do Anexo

msg.attach(MIMEText(body, 'html')) #mudando o formato do corpo para html
filename = pathlib.Path.cwd() / caminho_backup / loja1 / f'{dia_indicador.month}_{dia_indicador.day}_{loja1}.xlsx'
attachment = open(str(filename),'rb') #filename precisa ser convertido em str
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msg.attach(part)
attachment.close()

#Aqui termina a parte do Anexo

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "tnox bfyx mpdd pzuf")
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
server.quit()

print('\nEmail enviado com sucesso!')




# Passo 6 - Automatizar todas as lojas 

# Passo 7 - Criar Ranking para Diretoria

# In[10]:


#tabela faturamento ano

faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_ano = faturamento_lojas.sort_values(by = 'Valor Final', ascending = False)
display(faturamento_lojas_ano)

#exportando para excel

nome_arquivo = (f'{dia_indicador.month}_{dia_indicador.day}_RankingAnual.xlsx')
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


#faturamento dos dias 

vendas_dia = vendas.loc[vendas['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum(numeric_only=True)
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by= 'Valor Final', ascending = False)
display(faturamento_lojas_dia)


#exportando para excel


nome_arquivo = (f'{dia_indicador.month}_{dia_indicador.day}_RankingDiário.xlsx')
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


# Passo 8 - Enviando Arquivo para a diretoria 

# In[11]:


fromaddr = "caio.toscano345@gmail.com"
toaddr = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
msg = MIMEMultipart()

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = f'RankingDia {dia_indicador.day}/{dia_indicador.month}'

body = f'''Prezados, Bom Dia.

Melhor Loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R$ {faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior Loja do Dia em faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R$ {faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor Loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R$ {faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior Loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R$ {faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Atenciosamente, Caio Toscano (mestre do Python)

'''



#msg.
#Aqui começa a parte do Anexo

msg.attach(MIMEText(body, 'html')) #mudando o formato do corpo para html
filename = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_RankingAnual.xlsx'
attachment = open(str(filename),'rb') #filename precisa ser convertido em str
filename = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_RankingDiário.xlsx'
attachment = open(str(filename),'rb') #filename precisa ser convertido em str
part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msg.attach(part)
attachment.close()

#Aqui termina a parte do Anexo

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "tnox bfyx mpdd pzuf")
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
server.quit()

print('\nEmail enviado com sucesso!')

