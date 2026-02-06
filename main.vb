Sub lancamento()

Dim i As Integer
Dim chave As String
Dim cBusca As String
Dim nota As Long
Dim dInicio As String
Dim dFim As String
Dim nf As Worksheet
Dim Lmc As Worksheet
Dim plan As Workbook
Dim miros As Collection
Dim nMiro As Variant

Set plan = ActiveWorkbook
Set nf = Worksheets("NF")
Set Lcm = Worksheets("Lançar")

If Not IsObject(app) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set app = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = app.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject app, "on"
End If

session.findById("wnd[0]").maximize

nRows = Lcm.Cells(Lcm.Rows.Count, 2).End(xlUp).Row
nRowsOC = nf.Cells(nf.Rows.Count, 9).End(xlUp).Row + 1

If nRows >= 2 Then
        nf.Range("I2:I" & nRowsOC).ClearContents
        nf.Range("J2:J" & nRowsOC).ClearContents
        nf.Range("K2:K" & nRowsOC).ClearContents
End If

For linha = 2 To nRows
    
    doc = Lcm.Cells(linha, 4).Value
    chave = Lcm.Cells(linha, 2).Value
    nf.Range("E2").Value = chave
    plan.Connections("Consulta - fNotasFiscais").Refresh
    
    Do Until nf.Range("C2").Value = doc
        Application.Wait Now + TimeValue("00:00:01")
    Loop
    
    i = 2
    Set miros = New Collection
    Do Until nf.Cells(i, 1).Value = ""
        nota = nf.Cells(i, 1).Value
        dInicio = nf.Cells(i, 2).Value
        dFim = DateAdd("d", 60, dInicio)
        dInicio = Replace(dInicio, "/", ".")
        dFim = Replace(dFim, "/", ".")
        
        If i = 2 Then
            session.findById("wnd[0]/tbar[0]/okcd").Text = "J1B3N"
            session.findById("wnd[0]").sendVKey 0
        End If

        session.findById("wnd[0]").sendVKey 4
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[1,24]").Text = nota '"189496"
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]").Text = nf.Range("F2").Value '"807"
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]").SetFocus
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]").caretPosition = 3
        session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB006/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/btnG_SELFLD_TAB-MORE[8,56]").press
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL").Select
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").Text = dInicio '"01.05.2025"
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").Text = dFim '"31.05.2025"
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").SetFocus
        session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 10
        session.findById("wnd[2]/tbar[0]/btn[8]").press
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[1]").sendVKey 0
        Application.Wait Now + TimeValue("00:00:01")
        session.findById("wnd[0]").sendVKey 0
        On Error GoTo proximo:
        session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB6").Select
        miro = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB6/ssubHEADER_TAB:SAPLJ1BB2:2600/txtJ_1BDYDOC-BELNR").Text
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        If miro <> "" Then
            miros.Add (miro)
        End If
proximo_i:
        i = i + 1
    Loop
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
If miros.Count = 0 Then GoTo semMiro

    For Each nMiro In miros
        session.findById("wnd[0]/tbar[0]/okcd").Text = "MIR4"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/txtRBKP-BELNR").Text = nMiro
        session.findById("wnd[0]").sendVKey 0
        
        j = 0
        nRowsOC = nf.Cells(nf.Rows.Count, 9).End(xlUp).Row + 1
        session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
        oc = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELN[7," & j & "]").Text
        Do Until oc = "__________"
            If j < 11 Then
                qtd = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-MENGE[4," & j & "]").Text
                Item = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELP[8," & j & "]").Text
                oc = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELN[7," & j & "]").Text
            Else
                qtd = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-MENGE[4,11]").Text
                Item = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELP[8,11]").Text
                oc = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELN[7,11]").Text
                
                k = 1
                On Error Resume Next
                session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M").verticalScrollbar.Position = k
                k = k + 1
            End If
            
            If oc <> "__________" Then
                nf.Cells(nRowsOC, 9).Value = oc
                nf.Cells(nRowsOC, 10).Value = Item
                nf.Cells(nRowsOC, 11).Value = qtd
                nRowsOC = nRowsOC + 1
            End If
            
            j = j + 1
        Loop
        
        
        session.findById("wnd[0]/tbar[0]/btn[3]").press
    Next nMiro

'colocar lançamento aqui

    session.findById("wnd[0]/tbar[0]/okcd").Text = "MIRO"
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").Text = Lcm.Cells(linha, 3).Value '"16.05.2025"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR").Text = Lcm.Cells(linha, 4) '"14263-1"
    session.findById("wnd[0]/usr/cmbRM08M-VORGANG").Key = "3"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").Text = Lcm.Cells(linha, 5) '"872,91"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-WAERS").Text = Lcm.Cells(linha, 6) '"BRL"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BUPLA").Text = Lcm.Cells(linha, 7) '"807"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/btnRM08M-XMSEL").press
    
    idx = 2
    Do While nf.Cells(idx, 9) <> ""
        If idx - 2 < 7 Then
            session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0," & idx - 2 & "]").Text = nf.Cells(idx, 9).Value '"4509958947"
            session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1," & idx - 2 & "]").Text = nf.Cells(idx, 10).Value '"20"
        Else
            session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/ctxtRM08M-EBELN[0,7]").Text = nf.Cells(idx, 9).Value '"4509958947"
            session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST/txtRM08M-EBELP[1,7]").Text = nf.Cells(idx, 10).Value '"20"
            session.findById("wnd[1]/usr/subMSEL:SAPLMR1M:6221/tblSAPLMR1MTC_MSEL_BEST").verticalScrollbar.Position = idx - 8
        End If
        idx = idx + 1
    Loop
    
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
    session.findById("wnd[0]").sendVKey 0
    
    idx_rp = 2
    scroll = 1
    Do While nf.Cells(idx_rp, 14) <> 0
        nItem = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELP[8,0]").Text
        quantidade = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-MENGE[4,0]").Text
        
        If Val(nItem) = nf.Cells(idx_rp, 10) And Val(quantidade) = nf.Cells(idx_rp, 11) Then
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WRBTR[3,0]").Text = Round(nf.Cells(idx_rp, 14).Value, 2) '"16,08 "
            cod = 10
tentei:
            On Error GoTo tente
                session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/cmbDRSEG-MWSKZ[" & cod & ",0]").Key = "C5"
    
            idx_rp = idx_rp + 1
        End If
        
        oc = session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-EBELN[7,1]").Text
        If oc <> "__________" Then
            session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6006/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M").verticalScrollbar.Position = scroll
            scroll = scroll + 1
        End If
    Loop
    
errofim:
    session.findById("wnd[0]/usr/btnRM08M-HEADER_COLLAPSE").press
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI").Select
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-LIFRE").Text = Lcm.Cells(linha, 10).Value '"534543"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-GSBER").Text = "840"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE").Text = "ZH"
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY").Select
    session.findById("wnd[1]").sendVKey 0
    
    session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/txtINVFO-ZBD1T").Text = "45"
    session.findById("wnd[0]/tbar[1]/btn[21]").press
    
    Dim iCFOP As Integer
    Dim cfop As String
    iCFOP = 0
    Do While Not cfop = "__________"
        If iCFOP < 12 Then
            cfop = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[27," & iCFOP & "]").Text
            If cfop = "" Then
                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[27," & iCFOP & "]").Text = Lcm.Cells(linha, 14).Value
            End If
        Else
            cfop = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[27,12]").Text
            If cfop = "" Then
                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL/ctxtJ_1BDYLIN-CFOP[27,12]").Text = Lcm.Cells(linha, 14).Value
            End If
            session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB1/ssubHEADER_TAB:SAPLJ1BB2:2100/tblSAPLJ1BB2ITEM_CONTROL").verticalScrollbar.Position = iCFOP - 11
        End If
        iCFOP = iCFOP + 1
    Loop
    
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_DOCNUM9_DIVIDED-DOCNUM8").Text = Lcm.Cells(linha, 15).Value '"00002572"
    session.findById("wnd[0]").sendVKey 0
    digito = session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpTAB8/ssubHEADER_TAB:SAPLJ1BB2:2800/subRANDOM_NUMBER:SAPLJ1BB2:2801/txtJ_1BNFE_ACTIVE-CDV").Text
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    
    Dim saldo As Double
    Dim texto As String
    
    saldo = CDbl(session.findById("wnd[0]/usr/txtRM08M-DIFFERENZ").Text)
    
    Do While saldo > CDbl("0,01") Or saldo < CDbl("-0,01")
        vAtual = CDbl(session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WRBTR[3,0]").Text)
        session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/tblSAPLMR1MTC_MR1M/txtDRSEG-WRBTR[3,0]").Text = vAtual + saldo
        session.findById("wnd[0]").sendVKey 0
        saldo = CDbl(session.findById("wnd[0]/usr/txtRM08M-DIFFERENZ").Text)
    Loop
    
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'confirmar
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    If nRows >= 2 Then
        nf.Range("I2:I" & nRowsOC).ClearContents
        nf.Range("J2:J" & nRowsOC).ClearContents
        nf.Range("K2:K" & nRowsOC).ClearContents
    End If
    
    Lcm.Cells(linha, 17).Value = "Lançado"
semMiro:
Next

MsgBox "Lançamentos Realizados com Sucesso!"

Exit Sub

proximo:
    session.findById("wnd[1]/tbar[0]/btn[12]").press
Resume proximo_i

tente:
    If cod <= 13 Then
        cod = cod + 1
        Resume tentei
    Else
        MsgBox "Erro"
        GoTo errofim
    End If


End Sub