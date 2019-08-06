from flask import Flask, url_for
import pandas as pd
import sqlite3 as db


app = Flask(__name__)


@app.route('/<name>')
def index(name):
    c = db.connect('alunos.db')

    tabela_aluno = pd.read_sql_query("""SELECT * FROM alunos WHERE NOME_GUERRA = '{}'
                                    """.format(name.upper()), c)

    tabela_aluno = tabela_aluno.drop(["index", "EMAIL", "NOME", "ANT"], axis=1)

    return tabela_aluno.to_html()

# with app.test_request_context():

#     print(url_for('index', name='diogofreitas'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
