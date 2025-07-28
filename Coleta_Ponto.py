import os
import time
import shutil
import sys
import threading
import psutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import logging
import json
from datetime import datetime, timedelta
from tkinter import Tk, scrolledtext, BOTH, WORD, END, Label, Entry, Button, StringVar, IntVar, Frame, X
from tkinter import ttk

# Caminho para o diretório de downloads
download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

# Função para configurar o driver com opções headless
def configurar_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Executa o Chrome em modo oculto
    chrome_options.add_argument("--no-sandbox")  # Segurança extra para o headless
    chrome_options.add_argument("--disable-gpu")  # Desativa a GPU para melhorar a estabilidade
    chrome_options.add_argument("--disable-dev-shm-usage")  # Evita problemas de memória compartilhada em headless
    chrome_options.add_argument("--disable-webgl")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--silent")

    # Silencia os logs do WebDriver
    service = Service(log_path=os.devnull)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def limpar_diretorio_exceto_txt(diretorio):
    """Remove todos os arquivos no diretório e subdiretórios, exceto arquivos .txt."""
    try:
        for raiz, _, arquivos in os.walk(diretorio):
            for arquivo in arquivos:
                caminho_completo = os.path.join(raiz, arquivo)
                if not arquivo.endswith(".txt"):
                    os.remove(caminho_completo)
    except Exception as e:
        print(f"Erro ao limpar o diretório {diretorio}: {e}")


def limpar_txt_marcacoes(diretorio_downloads):
    """Remove arquivos .txt que começam com 'marcacoes' na pasta de downloads do usuário."""
    try:
        for arquivo in os.listdir(diretorio_downloads):
            caminho_completo = os.path.join(diretorio_downloads, arquivo)
            if arquivo.startswith("marcacoes") and arquivo.endswith(".txt"):
                os.remove(caminho_completo)
    except Exception as e:
        print(f"Erro ao limpar arquivos marcacoes.txt na pasta de downloads: {e}")

def finalizar_processos_chrome():
    """Finaliza todos os processos do Chrome e ChromeDriver silenciosamente."""
    for processo in psutil.process_iter(attrs=['pid', 'name']):
        if processo.info['name'] in ['chrome', 'chromedriver']:
            try:
                processo.kill()
            except psutil.NoSuchProcess:
                pass
            except Exception:
                pass

import psutil

def finalizar_processos_chrome():
    """Finaliza todos os processos relacionados ao Google Chrome e ChromeDriver."""
    processos_alvo = ['chrome', 'chromedriver', 'Google Chrome']  # Adicione mais nomes conforme necessário

    for processo in psutil.process_iter(attrs=['pid', 'name']):
        try:
            if any(nome in processo.info['name'].lower() for nome in processos_alvo):
                processo.kill()
        except psutil.NoSuchProcess:
            pass  # O processo já foi encerrado
        except Exception:
            pass  # Ignora outros erros

def main():
    try:
        # Limpa o diretório de arquivos na rede
        diretorio_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto"
        logging.info(f"Limpando o diretório: {diretorio_rede}")
        limpar_diretorio_exceto_txt(diretorio_rede)

        # Limpa arquivos "marcacoes*.txt" na pasta de downloads do usuário
        diretorio_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        logging.info(f"Limpando arquivos 'marcacoes*.txt' no diretório: {diretorio_downloads}")
        limpar_txt_marcacoes(diretorio_downloads)

        # Lista de funções de download para execução sequencial
        logging.info("Iniciando a execução das funções de download.")
        funcoes_download = [
            download_bbh,
            download_BFLimpeza,
            download_BFPortaria4,
            download_BIPortaria2,
            download_BIPortaria3,
            download_BIPortaria,
            download_BIAdministrativo,
            download_BIPortariaBBH,
        ]

        for funcao in funcoes_download:
            logging.info(f"Executando função: {funcao.__name__}")
            if not tentar_novamente(funcao):
                logging.warning(f"Erro ao tentar executar a função {funcao.__name__}. Continuando para a próxima.")
    except Exception as e:
        logging.error(f"Erro durante a execução da função principal (main): {e}", exc_info=True)
    finally:
        logging.info("Processo concluído.")


def login(driver, username, password):
    """Realiza login no site com o usuário e senha especificados."""
    usuario = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/form/div[1]/header/div[1]/div[2]/div/div/div[1]/div/input'))
    )
    usuario.send_keys(username)

    senha = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/form/div[1]/header/div[1]/div[2]/div/div/div[2]/div/input'))
    )
    senha.send_keys(password)

    botao_login = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[1]/header/div[1]/div[2]/div/div/button'))
    )
    botao_login.click()

    try:
        elemento_opcional = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/div/div/div/div[2]/a[2]'))
        )
        elemento_opcional.click()
    except Exception as e:
        pass


def esperar_download(driver, download_original, download_novo, tempo_espera=15):
    """Espera pelo download e renomeia o arquivo."""
    inicio = time.time()
    while not os.path.exists(download_original):
        if time.time() - inicio > tempo_espera:
            return False
        time.sleep(1)

    # Renomeia o arquivo se ele foi baixado com sucesso
    try:
        os.rename(download_original, download_novo)
        time.sleep(3)
        return True
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return False


def tentar_novamente(func, *args, tentativas=3, **kwargs):
    for tentativa in range(1, tentativas + 1):
        try:
            print(f"Tentativa {tentativa} para executar {func.__name__}...")
            resultado = func(*args, **kwargs)
            if resultado is False:
                raise ValueError(f"Erro na lógica da função {func.__name__}.")
            print(f"Função {func.__name__} concluída com sucesso na tentativa {tentativa}.")
            print()  # Adiciona uma linha em branco
            return True
        except Exception as e:
            print(f"Erro na tentativa {tentativa} da função {func.__name__}: {e}")
            print()  # Adiciona uma linha em branco
    print(f"Função {func.__name__} falhou após {tentativas} tentativas.")
    print()  # Adiciona uma linha em branco
    return False

def download_bbh():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "rotinas@bruno.com.br", "bbh37014")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div/label'))
        ).click()  # Clicar em BBHCorredor

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[2]'))
        ).click()  # Clicar no botão de download portaria 1510

        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BBH\KAIROS 1510\BBHCorredor"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BBHCorredor.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BBHCorredor.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome()


def download_BFLimpeza():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "rh3@bruno.com.br", "bf37324")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[1]'))
        ).click()  # Clicar em BFLimpeza

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[2]/span'))
        ).click()  # Clicar no botão de download portaria 1510


        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BF\KAIROS 1510\BFLimpeza"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BFLimpeza.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BFLimpeza.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome()# Sempre fecha o navegador


def download_BFPortaria4():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "rh3@bruno.com.br", "bf37324")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[2]'))
        ).click()  # Clicar em BFPortaria4

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)


        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[3]/span'))
        ).click()  # Clicar no botão de download portaria 671

        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BF\KAIROS 671\BFPortaria4"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BFPortaria4.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BFPortaria4.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome()# Sempre fecha o navegador


def download_BIPortaria2():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "joseane.daros@bruno.com.br", "bi05145")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[1]'))
        ).click()  # Clicar em BIPortaria2

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[2]/span'))
        ).click()  # Clicar no botão de download portaria 1510

        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BI\KAIROS 1510\BIPortaria2"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BIPortaria2.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BIPortaria2.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome() # Sempre fecha o navegador

def download_BIPortaria3():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "joseane.daros@bruno.com.br", "bi05145")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[2]'))
        ).click()  # Clicar em BIPortaria3

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[2]/span'))
        ).click()  # Clicar no botão de download portaria 1510

        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BI\KAIROS 671\BIPortaria3"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BIPortaria3.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BIPortaria3.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome() # Sempre fecha o navegador


def download_BIPortaria():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "joseane.daros@bruno.com.br", "bi05145")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[3]'))
        ).click()  # Clicar em BIPortaria

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[3]/span'))
        ).click()  # Clicar no botão de download portaria 671


        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BI\KAIROS 671\BIPortaria"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BIPortaria.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BIPortaria.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome() # Sempre fecha o navegador


def download_BIAdministrativo():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "joseane.daros@bruno.com.br", "bi05145")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[4]'))
        ).click()  # Clicar em BIAdministrativo

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[3]/span'))
        ).click()  # Clicar no botão de download portaria 671


        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BI\KAIROS 671\BIAdministrativo"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BIAdministrativo.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BIAdministrativo.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome()# Sempre fecha o navegador


def download_BIPortariaBBH():
    driver = None
    try:
        driver = configurar_driver()
        driver.get("https://www.dimepkairos.com.br/")  # Recarrega a página

        login(driver, "joseane.daros@bruno.com.br", "bi05145")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[5]/label'))
        ).click()  # Clicar em todos os Relógios
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[1]/label'))
        ).click()  # Clicar em Matriz
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/div/div[6]/div[2]/div[5]'))
        ).click()  # Clicar em BIPortariaBBH

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[1]/div[2]/div/i'))
        ).click()  # Clicar na seta da data
        time.sleep(1)
        # Calcula as datas
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=4)).strftime("%d/%m/%Y")
        time.sleep(1)
        # Insere a data de ontem
        data_ontem_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[1]/div/input')
            )
        )
        data_ontem_input.clear()
        data_ontem_input.send_keys(ontem)
        time.sleep(1)
        # Insere a data de hoje
        data_hoje_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[1]/div[2]/div/input')
            )
        )
        data_hoje_input.clear()
        data_hoje_input.send_keys(hoje)
        time.sleep(5)
        # Clicar em pesquisar
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/form/div[2]/input')
            )
        ).click()

        time.sleep(3)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[5]/div[1]/div/div[2]/div[2]/div/input[1]'))
        ).click()  # Clicar em Exportar
        time.sleep(2)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[6]/div/div/div[3]/span'))
        ).click()  # Clicar no botão de download portaria 671

        # Define o caminho de destino
        destino_rede = r"\\192.168.100.230\Arquiv0s\SC-CNV  BI-BF\RH\Rotinas\Controle de Coleta de Ponto\Coleta BI\KAIROS 671\BIPortariaBBH"

        # Espera e renomeia o arquivo
        download_original = os.path.join(download_dir, "marcacoes.txt")
        download_novo = os.path.join(download_dir, "BIPortariaBBH.txt")
        if esperar_download(driver, download_original, download_novo):
            shutil.move(download_novo, os.path.join(destino_rede, "BIPortariaBBH.txt"))
            print(f"O download foi concluído, renomeado para {os.path.basename(download_novo)} e movido.")
            return True
        else:
            print("O download não foi concluído com sucesso.")
            return False
    except Exception as e:
        print("Ocorreu um erro durante o download:", e)
        return False
    finally:
        if driver:
            driver.quit()
        finalizar_processos_chrome() # Sempre fecha o navegador

########################################################################################################################

# Caminho do arquivo de configuração
config_file = os.path.join(os.path.dirname(sys.argv[0]), "config.json")

# Estado inicial do agendamento
estado_agendamento = "agendado"

# Função para carregar a configuração
def carregar_config():
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                logging.warning("Arquivo de configuração inválido. Usando valores padrão.")
    return {"hora_inicial": "12:00", "frequencia": 30}

# Função para salvar a configuração
def salvar_config(hora_inicial, frequencia):
    with open(config_file, 'w') as f:
        json.dump({"hora_inicial": hora_inicial, "frequencia": frequencia}, f)
    logging.info(f"Configuração salva: Hora Inicial = {hora_inicial}, Frequência = {frequencia} minutos.")

# Configurar o logger
def configurar_logger():
    pasta_base = os.path.dirname(os.path.abspath(__file__))
    nome_arquivo_log = datetime.now().strftime("Coleta_Pontos_%Y-%m-%d.log")
    caminho_log = os.path.join(pasta_base, nome_arquivo_log)

    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=caminho_log,
        filemode='a',
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )
    return caminho_log

# Redirecionar saída padrão e de erro para o log
def redirecionar_saidas():
    class LoggerWrapper:
        def __init__(self, stream, nivel):
            self.stream = stream
            self.nivel = nivel

        def write(self, mensagem):
            if mensagem.strip():
                logging.log(self.nivel, mensagem.strip())

        def flush(self):
            pass

    sys.stdout = LoggerWrapper(sys.stdout, logging.INFO)
    sys.stderr = LoggerWrapper(sys.stderr, logging.ERROR)

# Atualizar o logger
def atualizar_logger():
    global caminho_log
    caminho_log = configurar_logger()
    redirecionar_saidas()
    return caminho_log

# Limpar logs antigos
def limpar_logs_antigos(diretorio, dias_reter, arquivo_em_uso):
    limite_tempo = datetime.now() - timedelta(days=dias_reter)
    try:
        for arquivo in os.listdir(diretorio):
            if arquivo.startswith("Coleta_Pontos_") and arquivo.endswith(".log") and arquivo != arquivo_em_uso:
                caminho_arquivo = os.path.join(diretorio, arquivo)
                try:
                    data_arquivo = datetime.strptime(arquivo.replace("Coleta_Pontos_", "").replace(".log", ""), "%Y-%m-%d")
                    if data_arquivo < limite_tempo:
                        os.remove(caminho_arquivo)
                        logging.info(f"Arquivo antigo removido: {arquivo}")
                except ValueError as e:
                    logging.warning(f"Formato inesperado para o arquivo {arquivo}: {str(e)}")
    except Exception as e:
        logging.error(f"Erro ao tentar limpar logs antigos: {str(e)}")


# Configurar agendamento
def configurar_agendamento(hora_inicial, frequencia_minutos):
    global estado_agendamento, agendamento_timer

    # Cancela qualquer agendamento existente
    if 'agendamento_timer' in globals() and agendamento_timer.is_alive():
        agendamento_timer.cancel()

    agora = datetime.now()
    hora_inicial_dt = datetime.strptime(hora_inicial, "%H:%M").replace(year=agora.year, month=agora.month, day=agora.day)

    if hora_inicial_dt <= agora:
        hora_inicial_dt += timedelta(days=1)

    intervalo = hora_inicial_dt - agora
    minutos, segundos = divmod(intervalo.total_seconds(), 60)

    logging.info(f"Tarefa agendada para começar às {hora_inicial_dt.strftime('%Y-%m-%d %H:%M:%S')}. "
                 f"Tempo restante até o primeiro agendamento: {int(minutos)} minutos e {int(segundos)} segundos.")

    # Armazena o timer globalmente para cancelamento futuro
    agendamento_timer = threading.Timer(intervalo.total_seconds(), agendar_repeticoes, [hora_inicial_dt, frequencia_minutos])
    agendamento_timer.start()

def agendar_repeticoes(hora_inicial_dt, frequencia_minutos):
    global estado_agendamento
    while estado_agendamento != "encerrado":
        agora = datetime.now()
        if agora >= hora_inicial_dt:
            rodar_tarefa()
            hora_inicial_dt += timedelta(minutes=frequencia_minutos)
            logging.info(f"Próxima execução agendada para {hora_inicial_dt.strftime('%Y-%m-%d %H:%M:%S')}.")
            logging.info("\n")
        time.sleep(1)


# Rodar tarefa
def rodar_tarefa():
    global estado_agendamento
    try:
        estado_agendamento = "executando"

        # Atualizar o logger para um novo arquivo, se necessário
        global caminho_log
        caminho_log = atualizar_logger()

        # Limpar logs antigos
        limpar_logs_antigos(
            os.path.dirname(caminho_log),  # Usa o mesmo diretório do arquivo de log
            dias_reter=4,
            arquivo_em_uso=os.path.basename(caminho_log)
        )

        executar_script()
    except Exception as e:
        pass
    finally:
        estado_agendamento = "agendado"

# Função de execução principal (exemplo de script)
def executar_script():
    logging.info("Iniciando o script.")
    try:
        main()  # Substitua por sua função principal
        logging.info("Script executado com sucesso.")
    except Exception as e:
        logging.error("Erro durante a execução do script.", exc_info=True)
    finally:
        pass  # Bloco 'finally' presente para garantir limpeza, mas sem ações específicas

# Validar horário
def validar_horario(horario):
    try:
        datetime.strptime(horario, "%H:%M")
        return True
    except ValueError:
        return False

# Exibir interface gráfica
def exibir_interface_grafica(caminho_log):
    config = carregar_config()
    ultimo_conteudo = []  # Variável para armazenar o conteúdo anterior do log

    def carregar_log():
        nonlocal ultimo_conteudo
        global caminho_log  # Torna o caminho do log acessível

        # Atualiza o caminho do log se houver um novo arquivo
        novo_caminho_log = configurar_logger()
        if caminho_log != novo_caminho_log:
            caminho_log = novo_caminho_log
            logging.info(f"Caminho do log atualizado para: {caminho_log}")

        try:
            if os.path.exists(caminho_log):
                with open(caminho_log, 'r', encoding='latin1', errors='replace') as log_file:
                    conteudo = log_file.readlines()

                if conteudo != ultimo_conteudo:
                    text_area.delete(1.0, END)
                    for linha in conteudo:
                        if "INFO" in linha:
                            text_area.insert(END, linha, "info")
                        elif "ERROR" in linha:
                            text_area.insert(END, linha, "error")
                        elif "WARNING" in linha:
                            text_area.insert(END, linha, "warning")
                        else:
                            text_area.insert(END, linha)

                    text_area.yview_moveto(1.0)
                    ultimo_conteudo = conteudo
            else:
                text_area.delete(1.0, END)
                text_area.insert(END, "Arquivo de log ainda não criado.\n")
        except Exception as e:
            text_area.delete(1.0, END)
            text_area.insert(END, f"Erro ao carregar o log: {str(e)}\n")

        text_area.after(1000, carregar_log)

    def atualizar_agendamento():
        global estado_agendamento
        if estado_agendamento != "agendado":
            label_status.config(text="Erro: Não é possível atualizar enquanto a tarefa está executando!", fg="red")
            return

        nova_hora = entrada_hora.get()
        nova_frequencia = entrada_frequencia.get()
        if not validar_horario(nova_hora) or not nova_frequencia.isdigit():
            label_status.config(text="Erro: Hora ou frequência inválida!", fg="red")
            return

        nova_frequencia = int(nova_frequencia)
        salvar_config(nova_hora, nova_frequencia)  # Salva a configuração no arquivo
        configurar_agendamento(nova_hora, nova_frequencia)
        label_status.config(text="Agendamento atualizado com sucesso!", fg="green")

    def fechar_programa():
        global estado_agendamento
        estado_agendamento = "encerrado"
        logging.info("Encerrando o programa...")
        os._exit(0)

    # Configuração da janela principal
    janela = Tk()
    janela.title("Visualizador de Log e Configuração")
    janela.geometry("1300x700")

    # Definir ícone da aplicação
    icon_path = os.path.join(os.path.dirname(caminho_log), 'icone.ico')
    janela.iconbitmap(icon_path)

    # Estilo darkly
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TFrame", background="#222222")
    style.configure("TLabel", background="#222222", foreground="#ffffff")
    style.configure("TButton", background="#333333", foreground="#ffffff")

    # Configuração - Fundo escuro
    frame_config = Frame(janela, bg="#222222")
    frame_config.pack(fill=X, pady=0)  # Remova o espaçamento vertical ajustando pady=0

    Label(frame_config, text="Hora Inicial (HH:MM):", bg="#222222", fg="#ffffff").grid(row=1, column=0, padx=10)
    entrada_hora = Entry(frame_config, bg="#333333", fg="#ffffff", width=10)
    entrada_hora.grid(row=1, column=1, padx=10)
    entrada_hora.insert(0, config["hora_inicial"])  # Carregar hora inicial do arquivo

    Label(frame_config, text="Frequência (min):", bg="#222222", fg="#ffffff").grid(row=1, column=2, padx=10)
    entrada_frequencia = Entry(frame_config, bg="#333333", fg="#ffffff", width=5)
    entrada_frequencia.grid(row=1, column=3, padx=10)
    entrada_frequencia.insert(0, config["frequencia"])  # Carregar frequência do arquivo

    Button(frame_config, text="Atualizar Agendamento", command=atualizar_agendamento, bg="#444444", fg="#ffffff").grid(
        row=1, column=4, padx=10)

    label_status = Label(frame_config, text="", bg="#222222", fg="#ffffff")
    label_status.grid(row=1, column=5, pady=20)  # Ajuste pady aqui também se necessário

    # Área de log - Fundo preto e texto colorido
    text_area = scrolledtext.ScrolledText(
        janela,
        wrap=WORD,
        font=("Courier", 10),
        bg="black",
        fg="#90EE90",
        insertbackground="#90EE90",
        selectbackground="gray",
        selectforeground="black",
        borderwidth=0,
    )
    text_area.pack(fill=BOTH, expand=True, padx=0, pady=0)  # Remova qualquer padding extra

    text_area.tag_config("info", foreground="#90EE90")
    text_area.tag_config("error", foreground="red")
    text_area.tag_config("warning", foreground="yellow")

    carregar_log()
    janela.protocol("WM_DELETE_WINDOW", fechar_programa)
    janela.mainloop()

# Configuração inicial
if __name__ == "__main__":
    caminho_log = configurar_logger()
    redirecionar_saidas()

    # Carrega configuração inicial
    config = carregar_config()
    configurar_agendamento(config["hora_inicial"], config["frequencia"])

    # Exibir interface
    exibir_interface_grafica(caminho_log)

####################################################################

# Configuração inicial
if __name__ == "__main__":
    # Configura o logger e redireciona saídas
    caminho_log = configurar_logger()
    redirecionar_saidas()

    # Carrega configuração inicial
    config = carregar_config()

    # Configura o agendamento
    configurar_agendamento(config["hora_inicial"], config["frequencia"])

    # Exibe a interface gráfica (executa no mesmo thread principal)
    exibir_interface_grafica(caminho_log)





