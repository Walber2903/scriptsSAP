from modulos.sap import Sap

sapEnv = 'ERP IU - Produção'

s = Sap(sapEnv, language='PT',connectByDescr=True)

acao_orcamentaria = ["12OR.CHSF", "15X4.CHSF", "146A.CHSF", "3390.CHSF", "4101.CHSF", "4102.CHSF", "4103.CHSF", "4476.CHSF", "5107.CHSF", "2D61.CHSF", "2D63.CHSF"]

saldo_nome = ["12OR.CHSF 2022 - Saldo.XLSX", "15X4.CHSF 2022 - Saldo.XLSX", "146A.CHSF 2022 - Saldo.XLSX", "3390.CHSF 2022 - Saldo.XLSX", "4101.CHSF 2022 - Saldo.XLSX",
            "4102.CHSF 2022 - Saldo.XLSX", "4103.CHSF 2022 - Saldo.XLSX", "4476.CHSF 2022 - Saldo.XLSX", "5107.CHSF 2022 - Saldo.XLSX", "2D61.CHSF 2022 - Saldo.XLSX", "2D63.CHSF 2022 - Saldo.XLSX"]

session = s.session

session.findById("wnd[0]").maximize()

for item in range(acao_orcamentaria.__len__()):
    #1 - iniciado com a transação abaixo no campo de atalho de transacoes
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nz4fm2301"
    session.findById("wnd[0]").sendVKey(0)
    #2 - Preenchendo os dados do relatorio da transacao, sera um loop a partir daqui
    session.findById("wnd[0]/usr/ctxt$4FVERSN").text = "0" # campo versao
    session.findById("wnd[0]/usr/txt$ZANO").text = "2022" # campo ano
    session.findById("wnd[0]/usr/ctxt$ZPOINV").text = acao_orcamentaria[item] #campo Grupo Programa orcamento
    #3 - Rodar o relatorio
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    #4 - Filtrar para achar o Saldo
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "000001"
    session.findById("wnd[0]/tbar[1]/btn[32]").press()
    session.findById("wnd[1]/tbar[0]/btn[9]").press()
    session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "I-Saldo"
    session.findById("wnd[2]/tbar[0]/btn[0]").press()
    session.findById("wnd[3]/usr/lbl[2,2]").setFocus()
    session.findById("wnd[3]/usr/lbl[2,2]").caretPosition = 3
    session.findById("wnd[3]").sendVKey(2)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    #5 - Descendo o resultado do relatorio clicando na setinha para baixo
    session.findById("wnd[0]/usr").verticalScrollbar.position = session.findById("wnd[0]/usr").verticalScrollbar.Maximum
    i = 13
    while session.findById(f"wnd[0]/usr/lbl[5,{i}]").text != '* Total':
        i += 1
    #6 - Clicar no somatorio da coluna orcado
    session.findById(f"wnd[0]/usr/lbl[56,{i}]").setFocus()
    session.findById("wnd[0]").sendVKey(2)
    #7 - Procurar os documentos e as partidas individuais
    session.findById("wnd[1]/usr/lbl[1,1]").caretPosition = 25
    session.findById("wnd[1]").sendVKey(2)

    #8 - clicar para selecionar o layout
    session.findById("wnd[0]/tbar[1]/btn[33]").press()
    #9 - selecionar o layout EXTR_ACOES
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cmbG51_USPEC_LBOX").key = ""
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 119
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 117
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "119"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell()

    #10 - exportar relatorio
    session.findById("wnd[0]/mbar/menu[0]/menu[4]/menu[1]").select()
    #11 - indicar os dados para exportar o relatorio
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\\walberm\\Desktop\\Python" #indicar o diretorio
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = saldo_nome[item] #indicar o nome da planilha
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
    session.findById("wnd[1]/tbar[0]/btn[11]").press() #confirmar clicando no botao substituir