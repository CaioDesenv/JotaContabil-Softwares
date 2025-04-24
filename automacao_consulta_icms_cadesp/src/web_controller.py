from playwright.sync_api import sync_playwright
from utils import process_image
from processamento_planilha import ProcessamentoPlanilha
from smtp_config import SMTP_CONFIGURACAO
from time import sleep

controla_tempo = 1

def consulta_empresa(page, excel_cnpj):
    download_img_path = r"C:\Caminho\Para\download_img\imagem_capturada.png"
    
    dropdown_selector = "select#ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_tipoFiltroDropDownList"
    page.select_option(dropdown_selector, value="1")
    sleep(controla_tempo)
    
    image_selector = "img#ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_imagemDinamica"
    sleep(controla_tempo)
    image_locator = page.locator(image_selector)
    image_locator.screenshot(path=download_img_path)
    
    captcha_text = process_image(download_img_path).strip()
    sleep(controla_tempo)
    
    captcha_input = "input#ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_imagemDinamicaTextBox"
    page.fill(captcha_input, captcha_text)
    sleep(controla_tempo)
    
    cnpj_input = "input#ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_valorFiltroTextBox"
    page.fill(cnpj_input, str(excel_cnpj).strip())
    sleep(controla_tempo)
    
    consult_button = "input#ctl00_conteudoPaginaPlaceHolder_filtroTabContainer_filtroEmitirCertidaoTabPanel_consultaPublicaButton"
    page.click(consult_button)
    sleep(controla_tempo)
    
    try:
        situacao_xpath = "//td[contains(text(), 'Situação')]/following-sibling::td[1]"
        page.wait_for_selector(f"xpath={situacao_xpath}", timeout=10000)
        situacao_text = page.eval_on_selector(f"xpath={situacao_xpath}", "el => el.innerText").strip()
        print("Status capturado:", situacao_text)
        if situacao_text not in {"Ativo", "Inativo", "Suspenso", "Baixado"}:
            situacao_text = None
        
        try:
            ocorrencia_xpath = "//td[b[contains(text(), 'Ocorrência')]]/following-sibling::td[1]"
            page.wait_for_selector(f"xpath={ocorrencia_xpath}", timeout=10000)
            ocorrencia_text = page.eval_on_selector(f"xpath={ocorrencia_xpath}", "el => el.innerText").strip()
            print("Ocorrência Fiscal capturada:", ocorrencia_text)
        except Exception as e:
            print("Falha na extração da Ocorrência Fiscal:", e)
            ocorrencia_text = "Erro na extração"
        
        return situacao_text, ocorrencia_text
    
    except Exception as e:
        print("Falha na extração do status:", e)
        return None, None

def realizar_consultas_em_lote():
    excel_path = r"C:\Caminho\Para\Relatorio-de-consulta-CADESP.xlsx"
    
    processador = ProcessamentoPlanilha(excel_path, SMTP_CONFIGURACAO)
    url = "https://www.cadesp.fazenda.sp.gov.br/Pages/Cadastro/Consultas/ConsultaPublica/ConsultaPublica.aspx"
    total_empresas = 128
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url)
        sleep(controla_tempo)
        
        for i in range(83, total_empresas + 1):
            cnpj_cell = processador.sheet.cell(row=i, column=3)
            cnpj_value = cnpj_cell.value
            print(f"Processando empresa na linha {i} com CNPJ: {cnpj_value}")
            
            resultado, ocorrencia = None, None
            attempts = 0
            while attempts < 10 and resultado is None:
                attempts += 1
                print(f"Tentativa {attempts} para o CNPJ {cnpj_value}")
                page.goto(url)
                sleep(controla_tempo)
                try:
                    resultado, ocorrencia = consulta_empresa(page, cnpj_value)
                except Exception as e:
                    print(f"Erro na tentativa {attempts} para CNPJ {cnpj_value}: {e}")
                    resultado = None
                if resultado is None:
                    print(f"Tentativa {attempts} falhou para o CNPJ {cnpj_value}.")
            
            if resultado is None:
                resultado = "Nao Consultado"
            
            processador.atualizar_resultado(i, resultado, ocorrencia)
            print(f"Resultado para o CNPJ {cnpj_value}: {resultado}, Ocorrência Fiscal: {ocorrencia}")
            sleep(controla_tempo)
        browser.close()
        
    print("Processamento concluído. Planilha atualizada com os resultados das consultas.")

if __name__ == "__main__":
    realizar_consultas_em_lote()
