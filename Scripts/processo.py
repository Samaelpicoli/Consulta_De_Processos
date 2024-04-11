from botcity.web import WebBot, Browser, By
from webdriver_manager.chrome import ChromeDriverManager
import os
import pandas as pd
import time

class Processo:
    """
    Classe responsável por automatizar o preenchimento de formulários em
    um site e salvar os dados em um arquivo Excel.
    """
    def __init__(self, url, base_de_dados):
        """
        Inicializa a classe Processo.

        Args:
            url (str): A URL do site onde os formulários serão preenchidos.
            base_de_dados (str): O caminho para o arquivo Excel contendo os dados a 
            serem preenchidos nos formulários.
        """
        self._url = url
        self._base_de_dados = base_de_dados
        self._driver = WebBot()
        self._dados = None
        self._inicializar_driver()
        self._ler_dados()

    def _inicializar_driver(self):
        """
        Inicializa o WebDriver.

        Raises:
            Exception: Se ocorrer um erro durante a inicialização do WebDriver.
        """
        try:
            self._driver.headless = False
            self._driver.browser = Browser.CHROME
            self._driver.driver_path = ChromeDriverManager().install()
            self._driver.browse(self._url)
            self._driver.maximize_window()
        except Exception as e:
            print(e)
            raise Exception(f'Erro durante a inicialização do webdriver. Detalhes {e}')

    def _ler_dados(self):
        """
        Realiza a leitura dos dados do arquivo Excel.

        Raises:
            Exception: Se ocorrer um erro durante a leitura dos dados.
        """
        extensao = os.path.splitext(self._base_de_dados)
        if 'xlsx' in extensao[1]:
            self._dados = pd.read_excel(self._base_de_dados)
            self._dados['Status'] = self._dados['Status'].astype(str)
        else:
            raise Exception('Arquivo Não Permitido')

    def preencher_formulario(self):
        """
        Preenche os formulários do site com os dados fornecidos no arquivo Excel.
        """
        for index, linha in self._dados.iterrows():
            self._clicar_botao()
            self._selecionar_cidade(linha['Cidade'])
            self._preencher_campos(linha['Nome'], linha['Advogado'], linha['Processo'])
            self._registrar()
            resultado = self._aguardar_resultado()
            self._atualizar_status(index, resultado)
            self._fechar_aba()
        self._driver.stop_browser()

    def _clicar_botao(self):
        """
        Clica em um botão na página.

        Raises:
            Exception: Se ocorrer um erro ao clicar no botão.
        """
        try:
            self._driver.find_element('button', By.TAG_NAME, waiting_time=40000, ensure_visible=True).click()
        except Exception as e:
            print(e)
            raise Exception(f'Erro ao clicar no botão dropdown. Detalhes {e}')

    def _selecionar_cidade(self, cidade):
        """
        Seleciona a cidade na página.

        Args:
            cidade (str): O nome da cidade a ser selecionada.

        Raises:
            Exception: Se ocorrer um erro ao selecionar a cidade.
        """
        try:
            self._driver.find_element(cidade, By.PARTIAL_LINK_TEXT).click()
        except Exception as e:
            print(e)
            raise Exception(f'Erro ao selecionar a cidade no menu dropdown. Detalhes {e}')
        
    def _preencher_campos(self, nome, advogado, numero):
        """
        Preenche os campos do formulário na página.

        Args:
            nome (str): O nome a ser preenchido no campo 'Nome'.
            advogado (str): O advogado a ser preenchido no campo 'Advogado'.
            numero (str): O número do processo a ser preenchido no campo 'Número do Processo'.

        Raises:
            Exception: Se ocorrer um erro ao preencher os campos.
        """
        try:
            lista_abas = self._driver.get_tabs()
            self._driver.activate_tab(lista_abas[-1])
            self._driver.find_element('nome', By.ID, waiting_time=3000).send_keys(nome)
            self._driver.find_element('advogado', By.ID).send_keys(advogado)
            self._driver.find_element('numero', By.ID).send_keys(numero)
        except Exception as e:
            print(e)
            raise Exception(f'Erro durante o preenchimento dos campos. Detalhes {e}')

    def _registrar(self):
        """
        Registra as informações preenchidas no formulário.

        Raises:
            Exception: Se ocorrer um erro ao registrar as informações.
        """
        try:
            self._driver.find_element('registerbtn', By.CLASS_NAME).click()
            dialogo = self._driver.get_js_dialog()
            dialogo.accept()
        except Exception as e:
            print(e)
            raise Exception(f'Erro ao clicar no botão registrar. Detalhes {e}')

    def _aguardar_resultado(self):
        """
        Aguarda o resultado do preenchimento do formulário.

        Returns:
            str: O resultado do preenchimento do formulário.

        Raises:
            Exception: Se ocorrer um erro ao aguardar o resultado.
        """
        try:
            while True:
                alerta = self._driver.get_js_dialog()
                if alerta is not None:
                    print(alerta.text)
                    return alerta.text
                time.sleep(1)
        except Exception as e:
            print(e)
            raise Exception(f'Erro durante a espera do alerta. Detalhes {e}')

    def _atualizar_status(self, index, resultado):
        """
        Atualiza o status do processo no DataFrame com base no resultado do preenchimento do formulário.

        Args:
            index (int): O índice do processo no DataFrame.
            resultado (str): O resultado do preenchimento do formulário.

        Raises:
            Exception: Se ocorrer um erro ao atualizar o status.
        """
        try:
            alerta = self._driver.get_js_dialog()
            alerta.accept()
            if "Processo encontrado com sucesso" in resultado:
                self._dados.loc[index, 'Status'] = "Encontrado"
            else:
                self._dados.loc[index, 'Status'] = "Não encontrado"
        except Exception as e:
            print(e)
            raise Exception(f'Erro ao alterar o status no DataFrame. Detalhes {e}')

    def _fechar_aba(self):
        """
        Fecha a aba do navegador.

        Raises:
            Exception: Se ocorrer um erro ao fechar a aba do navegador.
        """
        try:
            self._driver.close_page()
        except Exception as e:
            print(e)
            raise Exception(f'Erro ao fechar a aba. Detalhes {e}')

    def salvar_dataframe(self, diretorio, nome_arquivo): 
        """
        Salva o DataFrame em um arquivo Excel.

        Args:
            diretorio (str): O diretório onde o arquivo Excel será salvo.
            nome_arquivo (str): O nome do arquivo Excel.

        Returns:
            str: Uma mensagem indicando o sucesso da criação do arquivo Excel.

        Raises:
            Exception: Se ocorrer um erro ao criar o arquivo Excel.
        """
        try:
            os.makedirs(diretorio, exist_ok=True)
            self._dados.to_excel(os.path.join(diretorio, nome_arquivo), index=False)
            return 'Arquivo Criado com Sucesso'
        except Exception as e:
            print(e)
            raise Exception(f'Erro ao salvar o arquivo. Detalhes {e}')
