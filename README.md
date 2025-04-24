Este repositorio foi criado para adicionar todas as automações e softwares que estão sendo desenvolvidas para Jota Contabil.

1) CADESP
   -Esta automação foi desenvolvida para acessar um site da prefeitura de Itatiba e realizar a consulta da situação das empresa com o municipio, caso seja identificado durante a consulta
   problemas o codigo extrai da web a notificação e armazena em uma planilha, depois é so acionar um script que possui a configuração de um SMPT e automaticamente o
   mesmo envia um e-mail para o rh@jotacontabil informando os problemas encontrados.

2) RPA NFSE
   -Esta automação realiza a extração de NFSe no site da prefeitura de Itatiba e armazena separadamente cada arquivo em um modelo de pasta especifico para importar no sistema Dominio.
   Apos o donwload pode-se acionar outro script que realiza a transcrição dos arquivos XML para um planilha Excel coletando as informações para serem analisadas.

3) CONCILIADOR
   Foi criado um softaware que pegar duas fontes de informações e cruzam os dados para identificar inconsistencias no sistema ERP da Dominio.
   1 é preciso ter o excel ja extraido dos XMLs, depois é preciso gerar um excel que conterá a extração dos dados diretamente do banco de dados do sistema Dominio.
   Então é executar o sistema e efetuar as conciliações.

4) DOMÍNIO NFSE
   Esta automação realiza a extração de informações via ODBC no banco de dados da Dominio para gerar um relatorio para o confronto do conciliador.

5) AUTOMAÇÃO FGTS
   Esta automação foi criada com pyautogui inicialmente não tem muitas validaçõe porem esta funcionando e contem uma biblioteca criada manualmente para lidar com
   qualquer tipo de situação que poça surrgir. É rapida, le e atualiza situações em uma planilha.
