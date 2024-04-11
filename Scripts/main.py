import os
from processo import Processo

caminho = os.getcwd()
url = os.path.join(caminho, 'Paginas HTML','index.html')
base_de_dados = os.path.join(caminho, 'Base de Dados','Processos.xlsx')
nova_pasta = os.path.join(caminho, 'Arquivos Gerados')
nome_arquivo_gerado = 'Processos Finalizados.xlsx'

if __name__ == '__main__':
    bot = Processo(url, base_de_dados)
    bot.preencher_formulario()
    bot.salvar_dataframe(nova_pasta, nome_arquivo_gerado)