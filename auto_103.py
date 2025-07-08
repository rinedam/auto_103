# Importação das bibliotecas necessárias para automação web e manipulação de datas
try:
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service
    from selenium.webdriver.edge.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from datetime import datetime, timedelta
    import time
    import os
    from dotenv import load_dotenv
except ImportError as e:
    print(f"Erro ao importar módulos: {e}")
    raise

# Variável global para armazenar a função de callback que será usada para logging
adicionar_log_callback = None

def set_log_callback(callback):
    # Configura a função de callback para logging
    global adicionar_log_callback
    adicionar_log_callback = callback

def log(mensagem):
    # Função para registrar logs tanto no callback quanto no console
    if adicionar_log_callback:
        adicionar_log_callback(mensagem)
    print(mensagem)

# Pasta onde os arquivos baixados serão salvos
download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\DB_COMUM\\DB_103')

# Carrega as credenciais do arquivo .env
load_dotenv("credenciais.env")

def realizar_login(driver):
    # Realiza o login no sistema SSW usando as credenciais do arquivo .env
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))
    driver.find_element(By.NAME, "f1").send_keys(os.getenv("SSW_EMPRESA"))
    driver.find_element(By.NAME, "f2").send_keys(os.getenv("SSW_CNPJ"))
    driver.find_element(By.NAME, "f3").send_keys(os.getenv("SSW_USUARIO"))
    driver.find_element(By.NAME, "f4").send_keys(os.getenv("SSW_SENHA"))
    login_button = driver.find_element(By.ID, "5")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)

def formatar_nome_arquivo(data):
    # Formata o nome do arquivo usando o mês em português e o ano
    meses_pt = {1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"}
    mes_str = meses_pt[data.month]
    ano_str = data.strftime("%Y")
    return f"{mes_str}{ano_str}"

def esperar_download_e_renomear(pasta_download, novo_nome_arquivo):
    # Espera o download ser concluído e renomeia o arquivo para o formato desejado
    # Lista todos os arquivos no diretório exceto desktop.ini
    arquivos = [f for f in os.listdir(pasta_download) if f != 'desktop.ini']
    
    if not arquivos:
        print("Nenhum arquivo encontrado na pasta de download.")
        return False
    
    # Encontra o arquivo mais recente
    arquivo_mais_recente = max(
        arquivos,
        key=lambda f: os.path.getmtime(os.path.join(pasta_download, f))
    )
    
    caminho_antigo = os.path.join(pasta_download, arquivo_mais_recente)
    extensao = os.path.splitext(caminho_antigo)[1]
    caminho_novo = os.path.join(pasta_download, novo_nome_arquivo + extensao)
    
    if os.path.exists(caminho_novo):
        os.remove(caminho_novo)
    
    os.rename(caminho_antigo, caminho_novo)
    print(f"Sucesso! Arquivo renomeado para: {os.path.basename(caminho_novo)}")
    return True

def baixar_relatorio_por_data(driver, data, is_mes_atual=False):
    # Baixa o relatório para um período específico
    log(f"Iniciando processamento para {data.strftime('%B/%Y')}")
    
    # Aguarda e preenche o campo de opção
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f2").clear()
    driver.find_element(By.NAME, "f2").send_keys("CTA")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f3").send_keys("103")
    time.sleep(5)

    # Troca para a última aba aberta
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])
    
    # Preenche os campos do formulário
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "17")))

    driver.find_element(By.ID, "17").clear()
    driver.find_element(By.ID, "17").send_keys("e")
    time.sleep(0.3)
    
    # Configura o primeiro dia do mês
    primeiro_dia = data.replace(day=1)
    
    # Define a data final com base no parâmetro is_mes_atual
    if is_mes_atual:
        data_final = datetime.now()
        log(f"Usando data atual ({data_final.strftime('%d/%m/%y')}) para o mês corrente")
    else:
        if data.month == 12:
            ultimo_dia = data.replace(year=data.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            ultimo_dia = data.replace(month=data.month + 1, day=1) - timedelta(days=1)
        data_final = ultimo_dia
        log(f"Usando último dia do mês ({data_final.strftime('%d/%m/%y')})")
    
    # Preenche as datas no formulário
    driver.find_element(By.ID, "14").clear()
    driver.find_element(By.ID, "14").send_keys(primeiro_dia.strftime("%d%m%y"))
    time.sleep(0.3)
    driver.execute_script("document.getElementById('15').value = '';")
    driver.find_element(By.ID, "15").send_keys(data_final.strftime("%d%m%y"))
    time.sleep(0.3)
    driver.find_element(By.ID, "20").click()
    
    time.sleep(25)

def main(callback=None):
    # Função principal que coordena todo o processo de automação
    # Configura o callback para logs
    set_log_callback(callback)
    
    log("Iniciando processo de automação...")
    
    # Configura as opções do navegador Edge
    edge_options = Options()
    edge_prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safeBrowse.enabled": True
    }
    edge_options.add_experimental_option("prefs", edge_prefs)
    
    # Calcula as datas para processamento (mês atual, anterior e retrasado)
    data_atual = datetime.now()
    primeiro_dia_mes_atual = data_atual.replace(day=1)
    
    # Calcula o primeiro dia do mês anterior
    data_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    primeiro_dia_mes_anterior = data_mes_anterior.replace(day=1)
    
    # Calcula o primeiro dia do mês retrasado
    data_mes_retrasado = primeiro_dia_mes_anterior - timedelta(days=1)
    primeiro_dia_mes_retrasado = data_mes_retrasado.replace(day=1)

    # Lista de datas a serem processadas
    datas_a_processar = [
        (data_atual, True),  # Mês atual
        (primeiro_dia_mes_anterior, False),  # Mês anterior
        (primeiro_dia_mes_retrasado, False)  # Mês retrasado
    ]

    try:
        # Processa cada mês na lista
        for data, is_mes_atual in datas_a_processar:
            log(f"\nProcessando mês: {data.strftime('%B/%Y')}")
            # Cria uma nova instância do driver para cada mês
            service = Service()
            driver = webdriver.Edge(service=service, options=edge_options)
            
            try:
                log("Realizando login no sistema...")
                realizar_login(driver)
                baixar_relatorio_por_data(driver, data, is_mes_atual)
                nome_arquivo = formatar_nome_arquivo(data)
                log(f"Renomeando arquivo para: {nome_arquivo}")
                esperar_download_e_renomear(download_folder, nome_arquivo)
                time.sleep(5)
            finally:
                # Fecha o driver após processar cada mês
                driver.quit()
                log(f"Processamento do mês {data.strftime('%B/%Y')} finalizado.")

    except Exception as e:
        log(f"Erro: {str(e)}")
        raise
    finally:
        log("Processo finalizado.")

# Ponto de entrada do script quando executado diretamente
if __name__ == "__main__":
    main()