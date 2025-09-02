from playwright.sync_api import sync_playwright, expect
import time
import pandas as pd
import openpyxl


df = pd.read_excel(r"G:/Meu Drive/COMUNICAÇÃO DE EVENTO.xlsx")
df['DATA E HORA DO OCORRIDO'] = pd.to_datetime(df['DATA E HORA DO OCORRIDO'])
df['data'] = df['DATA E HORA DO OCORRIDO'].dt.strftime('%Y-%m-%d')
df['hora'] = df['DATA E HORA DO OCORRIDO'].dt.strftime('%H:%M')
from datetime import datetime


agora = datetime.now()

dataHoje = agora.strftime('%Y-%m-%d')
horaHoje = agora.strftime('%H:%M')

with sync_playwright() as pw:
    nav = pw.chromium.launch(headless=False)
    page = nav.new_page()
    page.goto("https://sga.hinova.com.br/sga/sgav4_alfa_cb/v5/login.php")

    page.wait_for_selector('div.modal')
    page.get_by_role("button", name="Continuar e Fechar").click()
    
    page.get_by_role("textbox", name="Usuário").fill("Romulofernando")
    page.get_by_role("textbox", name="Senha").fill("R0mul0")

    # iframe_element = page.wait_for_selector('iframe[title="reCAPTCHA"]')
    # frame = iframe_element.content_frame()
    # frame.get_by_role("checkbox", name="Não sou um robô").click()

    # page.get_by_role("button", name="Entrar").click()
    page.wait_for_selector('div.modal',timeout=60000)
    page.get_by_role("button", name="Fechar").click()
    

    for i, linha in df.iterrows():
        page.get_by_role("link", name="Evento").click()
        page.get_by_role("link", name="9.1) Cadastrar Evento").click()
        page.get_by_role('textbox', name='Placa').fill(linha[r'PLACA DO VEÍCULO'])        
        
        
        try:
            page.get_by_text(f"{linha[r'PLACA DO VEÍCULO']} - ATIVO").click(timeout=3000)
            page.get_by_role('link', name='Pesquisar').click()
            


        
        except Exception as e:
            print(f"Placa {linha[r'PLACA DO VEÍCULO']} não encontrada!")
            # Se quiser, pode continuar, pular para próxima linha ou fazer outra ação
            continue

        page.select_option("select#eventomotivo", label=f"{linha['TIPO DE EVENTO']}")
        page.select_option("select#envolvimento_evento", label=f"{linha['ENVOLVIMENTO']}")
        page.get_by_role("textbox", name="Data / Hora Evento (Quando").fill(f"{linha['data']}")
        page.locator("#novoevento_hora_evento").fill(f"{linha['hora']}")
        page.locator("#novoevento_data_comunicado_evento").fill(dataHoje)
        page.locator("#novoevento_hora_comunicado_evento").fill(horaHoje)
        page.get_by_role("button", name="Cadastrar Informações Gerais").click()
        page.get_by_role("button", name="Fechar").click()
        page.get_by_role("tab", name="Dados Ocorrência").click()
        page.get_by_role("tab", name="Dados", exact=True).click()
        page.locator("#numero_bo_novoevento").fill("1")
        page.get_by_label("Houve Vítima? *").select_option("N")
        page.get_by_role("tab", name="Local do Evento").click()
        page.get_by_role("textbox", name="Logradouro   *").fill(f"{linha['[INFORME O LOCAL DO OCORRIDO] Endereço']}")
        page.get_by_role("spinbutton", name="Numero   *").fill(f"{linha['[INFORME O LOCAL DO OCORRIDO] Nº']}")
        page.get_by_role("textbox", name="Bairro  *").fill(f"{linha['[INFORME O LOCAL DO OCORRIDO] Cidade']}")
        page.get_by_role("textbox", name="Cidade   *").fill(f"{linha['[INFORME O LOCAL DO OCORRIDO] Cidade']}")
        page.get_by_label("Estado *").select_option(f"{linha['[INFORME O LOCAL DO OCORRIDO] Estado']}")
        page.get_by_role("tab", name="Descrição do Evento para Área").click()
        page.locator("#novoevento_descricao_evento_area_ifr").content_frame.locator("#tinymce").fill(linha['DESCREVA O OCORRIDO'])
        page.get_by_role("tab", name="Reparo do Veículo", exact=True).click()
        page.get_by_role("tab", name="Reparo", exact=True).click()
        page.locator("#fornecedor_veiculo").fill("alfa")
        page.get_by_text("ALFA - - 36.117.103/0002-").click()
        page.get_by_role("button", name="Cadastrar Reparo do Veículo").click()
        page.get_by_role("button", name="Fechar").click()

    
        
    

    
    nav.close()