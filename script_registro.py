# %%
import pandas as pd
import os
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
from selenium.common.exceptions import TimeoutException
load_dotenv()



# %%
folderFile = os.path.join(os.getcwd(), 'planilha_registro.xlsx')

# %%
df = pd.read_excel(folderFile,dtype={'Documento do cooperado': str})


# %%
df_filtrado = df[df['Protocolo Visão'].isna()].copy()

# %%
df_filtrado.head()

# %%
# Bloco 1 – Acessar a URL do Visão 360 (com webdriver-manager)
 
options = Options()
options.add_argument("--start-maximized")  # Abre o navegador em tela cheia
 
# Inicializa o driver automaticamente com webdriver-manager
service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
 
# Acessa o Visão 360
url = "https://portal.sisbr.coop.br/auth/realms/sisbr/protocol/openid-connect/auth?&scope=openid&client_id=visao360-sisbr&response_type=code&redirect_uri=https://portal.sisbr.coop.br/visao360/consult"
driver.get(url)
 
time.sleep(5)

# %%
login = 'usuário'
senha = 'senha'

 
# Espera e preenche o campo de login
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(login)
 
# Preenche o campo de senha
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(senha)
 
# Clica no botão "Logar"
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "kc-login"))).click()
 
# Espera o QR Code aparecer (tempo ajustável, ex: 30s para você escanear e digitar o código)
print("Aguardando escaneamento do QR Code e inserção do código...")
WebDriverWait(driver, 60).until(lambda d: "visao360/consult" in d.current_url)
print("QR Code validado com sucesso!")

# %%
def preencher_campo_com_validacao(campo_id, texto, tentativas=3):

    for tentativa in range(1, tentativas + 1):

        try:

            campo = WebDriverWait(driver, 10).until(

                EC.element_to_be_clickable((By.ID, campo_id))

            )
            
            # Limpar o campo antes de preencher
            
            campo.clear()
            time.sleep(0.3)  # Pequena pausa para garantir que o campo está limpo

            actions = ActionChains(driver)
            
            actions.click(campo).pause(0.3)
 
            for letra in texto:

                actions.send_keys(letra).pause(0.1)
 
            actions.send_keys(Keys.ARROW_DOWN).pause(0.3).send_keys(Keys.ENTER).perform()
 
            time.sleep(1.5)  # tempo para Angular aplicar a classe "ng-valid"

            classe = campo.get_attribute("class")

            if "ng-valid" in classe:

                print(f"✅ Campo '{campo_id}' preenchido e validado na tentativa {tentativa}.")

                return True

            else:

                print(f"⚠️ Tentativa {tentativa} falhou: '{campo_id}' ainda não está válido.")

        except Exception as e:

            print(f"❌ Erro na tentativa {tentativa} ao preencher o campo '{campo_id}': {e}")
 
    print(f"❌ Todas as tentativas falharam ao preencher o campo '{campo_id}'.")

    return False

# %%
def esperando_elemento_spinner():
    WebDriverWait(driver, 40).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "ngx-spinner")))
 

# %%
def preencher_documento_com_validacao_ng_valid(driver, cpf, tentativas=3):

    for tentativa in range(1, tentativas + 1):

        try:

            campo = WebDriverWait(driver, 10).until(

                EC.element_to_be_clickable((By.NAME, "document"))

            )
            
            campo.click()

            campo.clear()

            campo.send_keys(cpf)
 
            # Aguarda o Angular processar e aplicar a classe

            time.sleep(1.2)
 
            classe = campo.get_attribute("class")

            if "ng-valid" in classe:

                print(f"✅ Campo 'Documento' preenchido e validado na tentativa {tentativa}.")
                
                botao_consultar = WebDriverWait(driver, 10).until(

                    EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[3]/sc-button/button'))

                )

                
                botao_consultar.click()

                print("📤 Botão 'Consultar' clicado com sucesso.")
 
                return True

            else:

                print(f"⚠️ Tentativa {tentativa}: Documento preenchido mas ainda inválido.")

        except Exception as e:

            print(f"❌ Erro ao tentar preencher o campo 'Documento' na tentativa {tentativa}: {e}")
 
    print("❌ Todas as tentativas de validação do campo 'Documento' falharam.")

    return False

 

# %%
def formatar_documento(documento):
    """Formata CPF ou CNPJ, preenchendo zeros à esquerda quando necessário."""
    numeros = ''.join(filter(str.isdigit, str(documento)))

    if 0 < len(numeros) <= 11:
        numeros = numeros.zfill(11)
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    
    elif len(numeros) <= 14:
        numeros = numeros.zfill(14)
        return f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"
    
    else:
        print(f"⚠️ Documento inválido: {documento}")
        return documento


# %%
def aguardar_spinner_sumir():
    """Enquanto o spinner estiver visível, aguarda 1 segundo por ciclo, até sumir ou atingir o timeout."""
    timeout=30
    xpath_spinner = "ngx-spinner-overlay ng-tns-c59-0 ng-trigger ng-trigger-fadeIn ng-star-inserted ng-animating"
    tempo_inicial = time.time()
 
    while True:
        try:
            driver.find_element(By.CLASS_NAME, xpath_spinner)
            print("⏳ Aguardando carregamento...")
            time.sleep(1)
 
            # Interrompe se o tempo máximo for ultrapassado
            if time.time() - tempo_inicial > timeout:
                print(f"⚠️ Timeout: carregamento demorou mais que {timeout}s.")
                break
        except:
            print("✅ Spinner sumiu. Continuando o fluxo.")
            break

# %%
for linha, coluna in df_filtrado.iterrows():
    print(linha, coluna.iloc[0], coluna.iloc[1], coluna.iloc[2], coluna.iloc[3], coluna.iloc[4])
    
    documento_formatado = formatar_documento(coluna.iloc[0])
    cooperativa = coluna.iloc[1]
    categoria = coluna.iloc[3]
    subcategoria = "Api Sicoob"
    servico = coluna.iloc[4]
    protocolo_plad = coluna.iloc[2]
    descricao = coluna.iloc[6]    
    
    
    
        # Bloco 3 – Preencher CPF/CNPJ e clicar em "Consultar"
    
    if not preencher_documento_com_validacao_ng_valid(driver, documento_formatado):
        print("Documento inválido. Pulando para o próximo registro.")
        continue
    
    
    #-----------------------------------------------------------------------------------
    
    
    # Bloco 4 – Clicar no botão "Abrir"
    
    # Aguarda o botão "Abrir" ficar visível e clicável

    botao_abrir = WebDriverWait(driver, 10).until(

        EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/section/sc-content/sc-consult/div/div[2]/div/sc-card-content/div/main/form/div/div[4]/sc-card/div/sc-card-content/div/div/div[2]/sc-button/button'))

    )
    
    # Clica no botão "Abrir"
    
    botao_abrir.click()
    esperando_elemento_spinner()
    
    print("Botão 'Abrir' clicado com sucesso.")
    
    
    
    
    
        # Bloco 5 – Selecionar conta conforme a cooperativa
    
    # Aguarda o <select> de contas estar disponível
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "accounts"))
    )
    
    # Captura todas as opções da conta
    options = select_element.find_elements(By.TAG_NAME, "option")
    
    # Percorre as opções e seleciona a que contém a cooperativa
    conta_encontrada = False
    for option in options:
        if f"Coop: {cooperativa}" in option.text:
            option.click()
            conta_encontrada = True
            print(f"Conta da cooperativa {cooperativa} selecionada com sucesso.")
            break
    
    # Se nenhuma conta compatível for encontrada
    if not conta_encontrada:
        print(f"⚠️ Nenhuma conta com cooperativa {cooperativa} encontrada.")
        
        
        
        
            # Bloco 6 – Selecionar o produto "Cobrança"
    
    # XPath do elemento que representa o card do produto "Cobrança"
    xpath_cobranca = '//*[@id="products"]/div[10]/sc-card/div/div/div/div'
    
    # Aguarda o produto "Cobrança" ficar clicável
    produto_cobranca = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, xpath_cobranca))
    )
    
    
    produto_cobranca.click()
    esperando_elemento_spinner()
    time.sleep(1.5)  # tempo curto para a classe atualizar visualmente
    
    # Verifica se o produto foi selecionado (classe mudou para 'selected-product')
   
    if "selected-product" in produto_cobranca.get_attribute("class"):
        print("Produto 'Cobrança' selecionado com sucesso.")
    else:
        print("⚠️ Produto 'Cobrança' não foi selecionado corretamente.")
        
        
        
        
        
            # Bloco 7 – Abrir o formulário (clicar no botão com ícone de Registro de chamado)
    
    # Aguarda o botão de "Registro de chamado" estar visível e clicável
    botao_formulario = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@tooltip='Registro de chamado']"))
    )
    
    
    # Clica no botão
    botao_formulario.click()
    esperando_elemento_spinner()
    
    print("Formulário de chamado aberto com sucesso.")
    
    
    
    
    # Bloco 8 – Preencher corretamente o campo "Tipo de Atendimento"
    
    preencher_campo_com_validacao("serviceTypeId", "Chat Receptivo")
    esperando_elemento_spinner()
        
        
        
        
        
    # Bloco 9 – Preencher o campo "Categoria"
    
    preencher_campo_com_validacao("categoryId", categoria)
    esperando_elemento_spinner()
        
        
        
        
    # Bloco 10 – Preencher o campo "Subcategoria"
    
    preencher_campo_com_validacao("subCategoryId", subcategoria)
    esperando_elemento_spinner()





    # Bloco 11 – Preencher o campo "Serviço"
    
    preencher_campo_com_validacao("serviceId", servico)
    esperando_elemento_spinner()

    
    
    
    
    # Bloco 12 – Preencher o campo "Canal de autoatendimento"
    
    # Aguarda o select ficar presente

    try:
        # Tenta localizar o elemento dentro de 10 segundos
        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "Canal De Autoatendimento"))
        )

        # Usa Select do Selenium para selecionar pelo texto visível
        select = Select(select_element)
        select.select_by_visible_text("Não Se Aplica")

        # Aguarda e valida se o campo foi aceito (classe ng-valid)
        time.sleep(0.5)
        classe_atual = select_element.get_attribute("class")

        if "ng-valid" in classe_atual:
            print("✅ Campo 'Canal de autoatendimento' selecionado com sucesso.")
        else:
            print("⚠️ O campo 'Canal de autoatendimento' foi preenchido, mas não foi validado pela aplicação.")

    except TimeoutException:
        # Caso o campo não esteja presente na página
        print("ℹ️ Campo 'Canal de autoatendimento' não encontrado. Pulando este passo.")


    
    
    
    
    
    # Bloco 13 – Preencher o campo "Protocolo PLAD"
    
    # Aguarda o campo estar clicável

    campo_protocolo = WebDriverWait(driver, 10).until(

        EC.element_to_be_clickable((By.ID, "Protocolo Plad"))

    )
    
    # Clica e digita o valor da variável protocolo_plad
    
    campo_protocolo.click()
    esperando_elemento_spinner()

    campo_protocolo.clear()

    campo_protocolo.send_keys(protocolo_plad)
    
    # Aguarda e valida o preenchimento

    time.sleep(1)

    classe_atual = campo_protocolo.get_attribute("class")

    if "ng-valid" in classe_atual:

        print("✅ Campo 'Protocolo PLAD' preenchido e validado com sucesso.")

    else:

        print("⚠️ O campo 'Protocolo PLAD' foi preenchido, mas não foi validado pela aplicação.")





    # Bloco 14 – Preencher o campo "Descrição"
    
    # Verifica se a descrição é válida ou precisa usar fallback

    # Trata o NaN e verifica se tem conteúdo válido
    
    if pd.isna(descricao) or not isinstance(descricao, str) or len(descricao.strip()) < 10:
        descricao_final = "Chamado registrado automaticamente via RPA"
    else:
        descricao_final = descricao.strip()

        
    # Aguarda o textarea de descrição ficar clicável

    campo_descricao = WebDriverWait(driver, 10).until(

    EC.element_to_be_clickable((By.ID, "description"))

        )
        
    # Clica, limpa e digita a descrição

    
    campo_descricao.click()
    esperando_elemento_spinner()

    campo_descricao.clear()

    campo_descricao.send_keys(descricao_final)
        
        # Aguarda e valida o preenchimento

    time.sleep(1)

    classe_atual = campo_descricao.get_attribute("class")

    if "ng-valid" in classe_atual:

        print("✅ Campo 'Descrição' preenchido e validado com sucesso.")

    else:

        print("⚠️ O campo 'Descrição' foi preenchido, mas não foi validado pela aplicação.")
    
    
    
    
    # Bloco 15 – Clicar no botão "Registrar"
    
    # Aguarda o botão "Registrar" ficar clicável

    botao_registrar = WebDriverWait(driver, 20).until(

        EC.element_to_be_clickable((By.XPATH, '//*[@id="actionbar hide"]/div/div[2]/form/div/div[20]/sc-button/button'))

    )
    
    # Clica no botão
    
    botao_registrar.click()
    
    print("✅ Primeiro clique no botão 'Registrar' realizado com sucesso.")




    # Bloco 16 – Clicar no segundo botão "Registrar" (confirmação do modal)
    
    # Aguarda o botão "Registrar" dentro do modal ficar clicável

    botao_confirmar_registro = WebDriverWait(driver, 20).until(

        EC.element_to_be_clickable((By.XPATH, '//*[@id="modal"]/div/sc-modal-footer/div/div/div[2]/sc-button/button'))

    )
    
    # Clica no botão de confirmação
    #000
    
    botao_confirmar_registro.click()
    
    print("✅ Segundo clique no botão 'Registrar' (confirmação) realizado com sucesso.")




    # Aguarda o elemento com o protocolo ficar visível
    elemento_protocolo = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, '//*[@id="actionbar hide"]/div/div[2]/form/div/div[2]/sc-card/div/sc-card-content/div/div/div[1]/h5'))
    )

    # Extrai o texto (número do protocolo)
    numero_protocolo = elemento_protocolo.text.strip()
    print(f"✅ Protocolo capturado: {numero_protocolo}")


    # Atualiza o valor na linha correspondente (índice i dentro do loop for)
    df.loc[linha, 'Protocolo Visão'] = numero_protocolo
    print(f"✅ Protocolo {numero_protocolo} colado na planilha com sucesso.")
    
        # Salva de volta
    df.to_excel(folderFile, index=False)



    # Bloco – Clicar em "Finalizar Atendimento"

    try:

        # Aguarda o botão "Finalizar Atendimento" estar clicável

        btn_finalizar = WebDriverWait(driver, 20).until(

            EC.element_to_be_clickable((By.XPATH, '//*[@id="actionbar hide"]/div/div[2]/form/div/div[5]/sc-button/button'))

        )
        #000
        
        btn_finalizar.click()
        esperando_elemento_spinner()

        print("✅ Botão 'Finalizar Atendimento' clicado com sucesso.")

    except Exception as e:

        print("❌ Erro ao clicar no botão 'Finalizar Atendimento':", e)

    
    
    
    
    
        # Bloco – Clicar em "Confirmar" para finalizar o atendimento

        # Aguarda o botão 'Confirmar' estar clicável E sem overlay por cima

    try:

        WebDriverWait(driver, 10).until(

            EC.element_to_be_clickable((By.XPATH, '//*[@id="modal"]/div/main/div/div[4]/button'))

        )
    
        # Usa JavaScript para clicar se houver interceptação visual
        #000
        
        botao_confirmar = driver.find_element(By.XPATH, '//*[@id="modal"]/div/main/div/div[4]/button')

        driver.execute_script("arguments[0].click();", botao_confirmar)
    
        print("✅ Botão 'Confirmar' clicado com sucesso (via JS).")
    
    except Exception as e:

        print("❌ Erro ao clicar no botão 'Confirmar':", e)

 
 
 
    
    print("Fim  loop")
 

    
print("Todos chamados registrados") 

# %%
#aguardar_spinner_sumir("ngx-spinner-overlay ng-tns-c59-0 ng-trigger ng-trigger-fadeIn ng-star-inserted ng-animating")


