#!/usr/bin/env python
# coding: utf-8

#Importação dos módulos necessários
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
import re

# Conexão com o servidor smtp do google (outros provedores também possuem servidor) 
s = smtplib.SMTP('smtp.gmail.com', 587)
s.starttls()
s.login('email', 'senha')

# Importação da tabela com os dados dos alunos
alunos = pd.read_excel("alunos.xlsx")
# Ordena pela média decrescente
alunos = alunos.sort_values('MÉDIA', ascending = False)
#Atribui a clasificação de 1 até o comprimento da tabela
alunos['Classificação'] = list(range(1, alunos.shape[0]))
# Indexa a tabela pela classificação
alunos.index = alunos.Classificação

# Gera uma segunda tabela somente com os dados que serão encaminhados aos alunos
tabela = alunos.dropna(axis=1, how="all")
tabela = tabela.drop(["index", "EMAIL", "NOME", "POSTO/QUADRO", "ANT"], axis=1)

#indexa a tabela de notas pela classificação (Para poder "cruzar" com a tabela de dados dos alunos)
tabela.index = tabela["Classificação"]


# Loop que percorre o índice de classificação e envia os e-mails
for i in alunos["Classificação"]:
    #cria uma mensagem com os atributos: assunto e origem
    msg = EmailMessage()
    msg['Subject'] = 'Atualização de notas e classificação'
    msg['From'] = "xandao.labs@gmail.com"
    # pega o email da primeira tabela com base na clasificação (variável "i")
    msg['To'] = [alunos["EMAIL"][i].lower()]
    
    # Cria o corpo do e-mail
    texto = 'Fala, {0}!'.format(alunos["NOME DE GUERRA"][i])
    
    # Gera um html para anexar ao corpo do e-mail (pode ser que alguns alunos recebam um e-mail binário
    # Nesse caso, veja com os alunos para fornecer um novo e-mail ou você mesmo pode verificar os e-mails enviados e
    # mostrar o e-mail, imprimir em pdf e mandar, etc.
    HTML = """  <html>
                    <head>
                        <meta charset="utf-8">
                    </head>
                    <body>
                        <h3>Aqui vão as estatísticas da última rodada!</h3>"""
    
    #"Filtra" a tabela de notas pela classificação do indivíduo (variável "i")
    tabela_aluno = tabela[tabela.index == i]
    # gera uma tabela em html (armazenar o resultado na mesma variável é uma má prática)
    tabela_aluno = tabela_aluno.to_html()
    
    # a tabela é uma string com tags html, portanto podemos utilizar a função re.sub.
    # na verdade, poderia utilizar str.replace também. Vale à pena testar para firmar o aprendizado
    
    # Corrige algumas imperfeições na tabela html (só vi depois que enviei, por isso é importante enviar pra si 
    # e corrigir os erros antes de executar pra todo mundo
    tabela_aluno = re.sub(r"<th>Classificação</th", "", tabela_aluno)
    tabela_aluno = re.sub(r"<th>NOME DE GUERRA</th>", "<th>Classificação</th><th>NOME DE GUERRA</th>", tabela_aluno)
    tabela_aluno = re.sub(r"<th></th>", "", tabela_aluno)
    # Inclui um estilo na primeira linha (títulos das colunas)
    tabela_aluno = re.sub(r"<td", "<td style='text-align: center;'", tabela_aluno)
    
    # mais erros corrigidos
    tabela_aluno = re.sub(r"\n", "", tabela_aluno)
    
    # Quando escrevi isso, só eu e Deus sabamos o que era. Agora, só Deus. Mas funciona, segue o baile
    tabela_aluno = re.sub(r"      >    </tr>    <tr>      >                                                                </tr>  ", "", tabela_aluno)
    
    #Mais estilo
    tabela_aluno = re.sub(r"<td style=\'text-align: center;\'>{0}</td>".format(i), "", tabela_aluno)
    
    # Soma a tabela à váriável HTML criada acima
    HTML += tabela_aluno
    
    # Inclui o texto que vem após a tabela
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
    
    # Transforma conteúdo para inclusão na mensagem
    part1 = MIMEText(texto,'plain')
    part2 = MIMEText(HTML, 'html')
    
    # Inclui o corpo da mensagem
    msg.set_content(part1)
    msg.set_content(part2)
    
    # Tenta encaminhar o e-mail. Caso falhe, imprime o nome de quem não recebeu.
    # Como estava em ordem de classificação, pode-se fazer um slice em alunos["Classificação"] com base no primeiro
    try:
        s.sendmail(msg['From'], msg['To'], msg.as_string())
    except:
        print( alunos['NOME DE GUERRA'][i] + '\tClassificação: ' + i  + "\tNão recebeu o email")
        #fecha a conexão
        s.close()
        #encerra o programa
        exit()
