#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
import re


# In[2]:


s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login('xandao.labs@gmail.com', 'xand3d3v')


# In[3]:


alunos = pd.read_excel("alunos.xlsx")
alunos = alunos.sort_values('MÉDIA', ascending = False)
alunos['Classificação'] = list(range(1, 128))
alunos.index = alunos.Classificação


# In[4]:


tabela = alunos.dropna(axis=1, how="all")
tabela = tabela.drop(["index", "EMAIL", "NOME", "POSTO/QUADRO", "ANT"], axis=1)
tabela.index = tabela["Classificação"]


# In[5]:


for i in alunos["Classificação"]:
    msg = EmailMessage()
    msg['Subject'] = 'Atualização de notas e classificação'
    msg['From'] = "xandao.labs@gmail.com"
    msg['To'] = [alunos["EMAIL"][i].lower()]
    texto = 'Fala, {0}!'.format(alunos["NOME DE GUERRA"][i])

    HTML = """  <html>
                    <head>
                        <meta charset="utf-8">
                    </head>
                    <body>
                        <h3>Aqui vão as estatísticas da última rodada!</h3>"""

    tabela_aluno = tabela[tabela.index == i].to_html()

    tabela_aluno = re.sub(r"<th>Classificação</th", "", tabela_aluno)
    tabela_aluno = re.sub(r"<th>NOME DE GUERRA</th>", "<th>Classificação</th><th>NOME DE GUERRA</th>", tabela_aluno)
    tabela_aluno = re.sub(r"<th></th>", "", tabela_aluno)
    tabela_aluno = re.sub(r"<td", "<td style='text-align: center;'", tabela_aluno)
    tabela_aluno = re.sub(r"\n", "", tabela_aluno)
    tabela_aluno = re.sub(r"      >    </tr>    <tr>      >                                                                </tr>  ", "", tabela_aluno)
    tabela_aluno = re.sub(r"<td style=\'text-align: center;\'>{0}</td>".format(i), "", tabela_aluno)
    HTML += tabela_aluno

    HTML += """ <h3><b>Observações:</b></h3>
                <ul>
                    <li>Em caso de dúvidas sobre notas, procurar o <b>Departamento de Ensino.</b></li>
                    <li> TAF ainda não entrou na classificação.</li>
                    <li> Para utilizar no excel, basta copiar e colar a tabela acima.</li>
                </ul>
                <h4>1T(IM) Alexandre</h4>
                <p> <b>Acesse o código fonte em:</b> www.github.com/im-alexandre</p>
            </body>
            </html>"""

    part1 = MIMEText(texto,'plain')
    part2 = MIMEText(HTML, 'html')

    msg.set_content(part1)
    msg.set_content(part2)

    try:
        s.sendmail(msg['From'], msg['To'], msg.as_string())
    except:
        print(alunos['NOME DE GUERRA'][i] + "Não recebeu o email")
        s.close()
