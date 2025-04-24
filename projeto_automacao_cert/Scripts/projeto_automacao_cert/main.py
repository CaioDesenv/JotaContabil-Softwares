from controller.gui_controller import clicar_na_imagem
from model.planilha_model import PlanilhaModel
from utils.gui_windows import focar_janela_por_titulo, minimizar_todas_as_janelas, exibir_mensagem_ok, aguardar_janela_por_titulo
from utils.gui_teclas import seta_baixo
from utils.gui_inputs import pressionar_tab, pressionar_enter, solicitar_input_usuario
from utils.gui_elementos import aguardar_elemento, obter_posicao_elemento, clicar_em_elemento_img_tratada
from utils.ocr_tesseract import extrair_texto_janela_ativa
from utils.gui_mouse import clicar_em_posicao
import pyautogui
import time

'''
Aqui temos inicio da nossa automação para que ela funcione corretamente o navegador deverá estar aberto
na pagina do FGTS Digital e já estar logado com o certificado da empresa Jota Contabil CNPJ: 05369682000128

time.sleep(3)
obter_posicao()
time.sleep(60)
'''

exibir_mensagem_ok('Atenção!', 'Verifique se a tela de execução é 1920 x 1080\nVerifique tambem se já iniciou o Google Chrome\n com o login no site FGTS efetuado.')

valor_digitado = int(solicitar_input_usuario())

print("🚀 Iniciando automação GUI do FGTS Digital")

#trecho inicial para ler dados da plnilha Excel que vai conter os CNPJ a serem analisado.

planilha = PlanilhaModel("model\LISTA_FGTS.xlsx")
cnpjs = planilha.obter_cnpjs(valor_digitado)
razoes = planilha.obter_razao_social(valor_digitado)
img_lbl_obrigatorio = "assets\lbl_obrigatorio.PNG"
img_lbl_tipo_debito_obrigatorio = 'assets\lbl_tipo_debito_obrigatorio.PNG'
img_lbl_nao_ha_debitos = 'assets\lbl_nao_ha_debitos.PNG'
img_lbl_empregador = 'assets\lbl_empregador.PNG'
img_lbl_emissao_guia_rapida = 'assets\lbl_emissao_guia_rapida.PNG'

img_btn_trocal_perfil = 'assets\\btn_TrocarPerfil.PNG'
img_btn_procurador = 'assets\\btn_procurador.PNG'
img_btn_selecionar = 'assets\\btn_Selecionar.PNG'
img_btn_gestao_guias = 'assets\\btn_GestaoGuias.PNG'
img_btn_filtro_recisao = 'assets\\btn_filtro_recisao.PNG'
img_btn_emitir_guia = 'assets\\btn_emitir_guia.PNG'
img_btn_data = 'assets\\btn_data.PNG'
img_btn_filtro_recisorio = 'assets\\btn_filtro_recisorio.PNG'
img_btn_pesquisar_guia = 'assets\\btn_pesquisar_guia.PNG'
img_btn_emissao_guia_rapida = 'assets\\btn_emissao_guia_rapida.PNG'
img_btn_emissao_detalhe_guia = 'assets\\btn_detalhe_guia.PNG'
img_btn_retornar_paginal_inicial = 'assets\\btn_pagina_inicial_fgts_digital.PNG'

#Variavel criada para poder comparar com a extração de dados da pagina para verificar se tem Guia ou não
emissao_guia_rapida = '''
Existem vínculos desligados com cálculo da Indenização Compensatória pendente
'''

'''
# Simulando resultados da automação:
        status_gfd = "Baixado"
        status_detalhe = "Erro"
        planilha.atualizar_downloads(i, status_gfd, status_detalhe)
        planilha.atualizar_situacao(i, "Detalhe da guia não encontrada")
'''

minimizar_todas_as_janelas()
time.sleep(2)
focar_janela_por_titulo('FGTS Digital')

#Como eu criei errado: for i, cnpj, razao in enumerate(cnpjs, razao, start=2):

for i, (cnpj, razao_social) in enumerate(zip(cnpjs, razoes), start=valor_digitado):
    print(f"Linha {i} | CNPJ: {cnpj} | Razão Social: {razao_social}")
    try:
        #input("📌 Pressione ENTER assim que o site carregar e estiver visível para clicar no botão...")
        clicar_em_posicao(x = 151, y = 151)
        aguardar_elemento(img_btn_trocal_perfil)
        clicar_na_imagem(img_btn_trocal_perfil)
        clicar_na_imagem(img_lbl_obrigatorio)
        seta_baixo(2)
        pressionar_enter()
        pressionar_tab()
        time.sleep(1)
        pyautogui.write(f'{cnpj}', interval=0.02)
        clicar_na_imagem(img_btn_selecionar)
        aguardar_elemento(img_btn_gestao_guias)
        clicar_na_imagem(img_btn_gestao_guias)
        aguardar_elemento(img_btn_emissao_guia_rapida)
        clicar_na_imagem(img_btn_emissao_guia_rapida)
        aguardar_elemento(img_lbl_empregador)
        
        armazenar_texto_janela_ativa = extrair_texto_janela_ativa('por')

        if 'Não há débitos de interesse' in armazenar_texto_janela_ativa:
            planilha.atualizar_situacao(i, 'Não há débitos de interesse.')
            planilha.atualizar_downloads(i, 'Sem arquivo para download', 'Sem arquivo para download')
            clicar_em_posicao(x = 151, y = 151)
            time.sleep(2)
            
        elif 'Q.03/2025' in armazenar_texto_janela_ativa:
             seta_baixo(3)
             planilha.atualizar_situacao(i, 'Contem debitos!')
             if aguardar_elemento(img_btn_filtro_recisorio):
                 clicar_na_imagem(img_btn_filtro_recisorio)
             aguardar_elemento(img_btn_pesquisar_guia)
             clicar_na_imagem(img_btn_pesquisar_guia)
             seta_baixo(11)
             aguardar_elemento(img_btn_emitir_guia)
             clicar_na_imagem(img_btn_emitir_guia)
             aguardar_janela_por_titulo('Salvar como')
             pyautogui.write(f'{razao_social}-GFD', interval=0.05)
             pressionar_enter()
             planilha.atualizar_downloads_GFD(i,'Download Efetuado!')
             seta_baixo(18)
             aguardar_elemento(img_btn_emissao_detalhe_guia)
             clicar_na_imagem(img_btn_emissao_detalhe_guia)
             aguardar_janela_por_titulo('Salvar como')
             pyautogui.write(f'{razao_social}-Detalhe-Guia', interval=0.05)
             planilha.atualizar_downloads_detalhe_da_guia(i,'Download Efetuado!')
             pressionar_enter()
             time.sleep(2)
             #Clica na imagem para retornar a pagina inicial
             clicar_em_posicao(x = 151, y = 151)
             aguardar_elemento(img_lbl_empregador)

        else:
            planilha.atualizar_situacao(i, 'Data de pesquisa incorreta!')
            clicar_em_posicao(x = 151, y = 151)
            aguardar_elemento(img_lbl_empregador)

    except Exception as e:
        planilha.atualizar_situacao(i, f"Erro: {str(e)}")

#abrir_chrome_com_perfil()
#clicar_botao_certificado(img_lbl_descricao)
#planilha.atualizar_situacao(i, 'Contem Débitos')
print("✅ Processo concluído.")
