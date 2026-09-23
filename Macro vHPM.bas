Option Explicit

Private Const ABA_PAINEL As String = "Painel_SP99"
Private Const ABA_BASE As String = "Estoque_SP99"
Private Const TABELA_CENTROS As String = "tbCentrosSP99"
Private Const TABELA_LOG As String = "tbLogSP99"
Private Const TRANSACAO_PADRAO As String = "S_P99_41000062"
Private Const VARIANTE_PADRAO As String = "/ESTOQUE_SP9"
Private Const TIPO_MOEDA As String = "10"
Private Const COLUNAS_EXPORTACAO As Long = 14
Private Const COLUNA_CENTRO As Long = 15
Private Const PRIMEIRA_COLUNA_FORMULA As Long = 16
Private Const ULTIMA_COLUNA_FORMULA As Long = 23
Private Const TIMEOUT_PADRAO As Long = 180

Public Sub Atualizar_SP99_Todos()
    Dim wb As Workbook, wsPainel As Worksheet, wsBase As Worksheet
    Dim tblCentros As ListObject, lr As ListRow, sessao As Object
    Dim centro As String, ativo As String, anoFiscal As Long, periodo As Long
    Dim primeiraColeta As Boolean, totalLinhas As Long, inicio As Date, execucao As String
    Dim calcAnt As XlCalculation, eventosAnt As Boolean, telaAnt As Boolean, alertasAnt As Boolean
    Dim nErro As Long, dErro As String, excelPreparado As Boolean

    On Error GoTo TrataErro

    PrepararExecucao wb, wsPainel, wsBase, anoFiscal, periodo, inicio, execucao, calcAnt, eventosAnt, telaAnt, alertasAnt
    excelPreparado = True
    Set tblCentros = ObterTabela(wsPainel, TABELA_CENTROS)
    primeiraColeta = True

    AtualizarStatus wsPainel, "EM EXECUCAO", "Conectando ao SAP", "-", inicio, Empty, 0, ""
    Set sessao = ObterSessaoSAP()

    For Each lr In tblCentros.ListRows
        ativo = UCase$(Trim$(CStr(ValorTabela(lr, tblCentros, "Ativo"))))
        centro = UCase$(Trim$(CStr(ValorTabela(lr, tblCentros, "Centro"))))

        If (ativo = "SIM" Or ativo = "S" Or ativo = "YES") And Len(centro) > 0 Then
            AtualizarStatus wsPainel, "EM EXECUCAO", "Coletando centro", centro, inicio, Empty, totalLinhas, ""
            Application.StatusBar = "SP99: processando " & centro & "..."
            totalLinhas = totalLinhas + ColetarCentro(sessao, wsPainel, wsBase, centro, anoFiscal, periodo, primeiraColeta, execucao)
            primeiraColeta = False
        End If
    Next lr

    If primeiraColeta Then Err.Raise vbObjectError + 4100, "Atualizar_SP99_Todos", "Nenhum centro ativo foi encontrado."

    FinalizarBase wsBase
    AtualizarStatus wsPainel, "CONCLUIDO", "Finalizado", "-", inicio, Now, totalLinhas, "Coleta concluida."
    RegistrarLog wsPainel, execucao, "TODOS", "Finalizacao", "SUCESSO", totalLinhas, "", DateDiff("s", inicio, Now), "Coleta concluida", Environ$("Username")
    If excelPreparado Then RestaurarExcel calcAnt, eventosAnt, telaAnt, alertasAnt

    wsBase.Activate
    Application.GoTo wsBase.Range("A1"), True
    MsgBox "SP99 atualizada com sucesso." & vbCrLf & vbCrLf & _
           "Exercicio: " & anoFiscal & vbCrLf & _
           "Periodo: " & Format$(periodo, "00") & vbCrLf & _
           "Linhas com estoque maior que zero: " & Format$(totalLinhas, "#,##0"), vbInformation, "SP99"
    Exit Sub

TrataErro:
    nErro = Err.Number
    dErro = Err.Description
    On Error Resume Next
    AtualizarStatus wsPainel, "ERRO", "Execucao interrompida", centro, inicio, Now, totalLinhas, dErro
    RegistrarLog wsPainel, execucao, centro, "Erro", "ERRO", 0, "", DateDiff("s", inicio, Now), dErro, Environ$("Username")
    If excelPreparado Then RestaurarExcel calcAnt, eventosAnt, telaAnt, alertasAnt
    On Error GoTo 0
    MsgBox "Nao foi possivel concluir a SP99." & vbCrLf & vbCrLf & "Erro: " & nErro & vbCrLf & "Descricao: " & dErro, vbCritical, "SP99"
End Sub

Public Sub Atualizar_SP99_Centro()
    Dim centro As String
    centro = UCase$(Trim$(InputBox("Informe o centro." & vbCrLf & "Exemplo: BUSB", "SP99 - Centro")))
    If Len(centro) = 0 Then Exit Sub
    AtualizarCentroUnico centro
End Sub

Private Sub AtualizarCentroUnico(ByVal centro As String)
    Dim wb As Workbook, wsPainel As Worksheet, wsBase As Worksheet, sessao As Object
    Dim anoFiscal As Long, periodo As Long, linhas As Long, inicio As Date, execucao As String
    Dim calcAnt As XlCalculation, eventosAnt As Boolean, telaAnt As Boolean, alertasAnt As Boolean
    Dim nErro As Long, dErro As String, excelPreparado As Boolean

    On Error GoTo TrataErro
    PrepararExecucao wb, wsPainel, wsBase, anoFiscal, periodo, inicio, execucao, calcAnt, eventosAnt, telaAnt, alertasAnt
    excelPreparado = True

    Set sessao = ObterSessaoSAP()
    RemoverCentroDaBase wsBase, centro
    linhas = ColetarCentro(sessao, wsPainel, wsBase, centro, anoFiscal, periodo, False, execucao)
    FinalizarBase wsBase
    AtualizarStatus wsPainel, "CONCLUIDO", "Finalizado", centro, inicio, Now, linhas, "Centro atualizado."
    If excelPreparado Then RestaurarExcel calcAnt, eventosAnt, telaAnt, alertasAnt
    MsgBox "Centro " & centro & " atualizado." & vbCrLf & "Linhas com estoque maior que zero: " & linhas, vbInformation, "SP99"
    Exit Sub

TrataErro:
    nErro = Err.Number
    dErro = Err.Description
    On Error Resume Next
    AtualizarStatus wsPainel, "ERRO", "Execucao interrompida", centro, inicio, Now, linhas, dErro
    RegistrarLog wsPainel, execucao, centro, "Erro", "ERRO", 0, "", DateDiff("s", inicio, Now), dErro, Environ$("Username")
    If excelPreparado Then RestaurarExcel calcAnt, eventosAnt, telaAnt, alertasAnt
    On Error GoTo 0
    MsgBox "Nao foi possivel atualizar " & centro & "." & vbCrLf & vbCrLf & dErro, vbCritical, "SP99"
End Sub

Private Function ColetarCentro(ByVal sessao As Object, ByVal wsPainel As Worksheet, ByVal wsBase As Worksheet, ByVal centro As String, ByVal anoFiscal As Long, ByVal periodo As Long, ByVal limparDados As Boolean, ByVal execucao As String) As Long
    Dim inicio As Date, pasta As String, arquivo As String, caminho As String
    Dim wbExp As Workbook, wsExp As Worksheet, antes As Object
    Dim linhas As Long, timeout As Long, nErro As Long, dErro As String

    On Error GoTo TrataErro
    inicio = Now
    pasta = ResolverPastaExportacao()
    arquivo = MontarNomeArquivo(centro)
    caminho = pasta & arquivo
    timeout = CLng(Val(CStr(ValorNome("cfgSP99_Timeout", TIMEOUT_PADRAO))))
    If timeout <= 0 Then timeout = TIMEOUT_PADRAO

    On Error Resume Next
    If Len(Dir$(caminho)) > 0 Then Kill caminho
    On Error GoTo TrataErro

    RegistrarLog wsPainel, execucao, centro, "Inicio", "EM EXECUCAO", 0, arquivo, 0, "Iniciando coleta", Environ$("Username")
    Set antes = CapturarWorkbooksAbertos()
    ExecutarTransacao sessao, centro, periodo, anoFiscal, pasta, arquivo
    Set wbExp = AguardarExportacao(caminho, arquivo, timeout, antes)

    If wbExp Is Nothing Then Err.Raise vbObjectError + 4200, "ColetarCentro", "A exportacao do centro " & centro & " nao foi localizada."
    Set wsExp = LocalizarAbaExportada(wbExp)
    If wsExp Is Nothing Then Err.Raise vbObjectError + 4201, "ColetarCentro", "Nenhuma aba valida foi encontrada na exportacao de " & centro & "."

    linhas = ImportarCentro(wsExp, wsBase, centro, limparDados)
    wbExp.Close SaveChanges:=False
    Set wbExp = Nothing

    On Error Resume Next
    If Len(Dir$(caminho)) > 0 Then Kill caminho
    On Error GoTo TrataErro

    RegistrarLog wsPainel, execucao, centro, "Importacao", "SUCESSO", linhas, arquivo, DateDiff("s", inicio, Now), "Somente estoque > 0", Environ$("Username")
    ColetarCentro = linhas
    Exit Function

TrataErro:
    nErro = Err.Number
    dErro = Err.Description
    On Error Resume Next
    If Not wbExp Is Nothing Then wbExp.Close SaveChanges:=False
    RegistrarLog wsPainel, execucao, centro, "Importacao", "ERRO", 0, arquivo, DateDiff("s", inicio, Now), dErro, Environ$("Username")
    On Error GoTo 0
    Err.Raise nErro, "ColetarCentro", dErro
End Function

Private Sub ExecutarTransacao(ByVal sessao As Object, ByVal centro As String, ByVal periodo As Long, ByVal anoFiscal As Long, ByVal pasta As String, ByVal arquivo As String)
    Dim transacao As String, grid As Object
    transacao = CStr(ValorNome("cfgSP99_Transacao", TRANSACAO_PADRAO))

    sessao.findById("wnd[0]").maximize
    sessao.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & transacao
    sessao.findById("wnd[0]").sendVKey 0
    AguardarSAP sessao

    DefinirTextoSAP sessao, "wnd[0]/usr/ctxtP_WERKS", centro
    sessao.findById("wnd[0]").sendVKey 0
    DefinirTextoSAP sessao, "wnd[0]/usr/txtP_POPER", CStr(periodo)
    sessao.findById("wnd[0]").sendVKey 0
    DefinirTextoSAP sessao, "wnd[0]/usr/txtP_BDATJ", CStr(anoFiscal)
    sessao.findById("wnd[0]").sendVKey 0

    If ExisteObjetoSAP(sessao, "wnd[0]/usr/cmbP_CURTP") Then sessao.findById("wnd[0]/usr/cmbP_CURTP").Key = TIPO_MOEDA
    DefinirTextoSAP sessao, "wnd[0]/usr/ctxtP_VARIAN", VARIANTE_PADRAO
    sessao.findById("wnd[0]").sendVKey 0
    sessao.findById("wnd[0]/tbar[1]/btn[8]").Press
    AguardarSAP sessao

    Set grid = sessao.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")
    grid.SetCurrentCell 0, "VBELN"
    grid.SelectedRows = "0"
    grid.ContextMenu
    grid.SelectContextMenuItem "&XXL"
    AguardarSAP sessao

    If ExisteObjetoSAP(sessao, "wnd[1]/tbar[0]/btn[0]") Then
        sessao.findById("wnd[1]/tbar[0]/btn[0]").Press
        AguardarSAP sessao
    End If

    PreencherDialogoExportacao sessao, pasta, arquivo
End Sub

Private Sub PreencherDialogoExportacao(ByVal sessao As Object, ByVal pasta As String, ByVal arquivo As String)
    Dim inicio As Date
    inicio = Now

    Do While DateDiff("s", inicio, Now) < 30
        DoEvents

        If ExisteObjetoSAP(sessao, "wnd[1]/usr/ctxtDY_PATH") Then
            sessao.findById("wnd[1]/usr/ctxtDY_PATH").Text = RemoverBarraFinal(pasta)
        End If

        If ExisteObjetoSAP(sessao, "wnd[1]/usr/ctxtDY_FILENAME") Then
            sessao.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = arquivo
            sessao.findById("wnd[1]/usr/ctxtDY_FILENAME").SetFocus
            sessao.findById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = Len(arquivo)

            If ExisteObjetoSAP(sessao, "wnd[1]/tbar[0]/btn[0]") Then
                sessao.findById("wnd[1]/tbar[0]/btn[0]").Press
            ElseIf Not PressionarBotaoPorTexto(sessao.findById("wnd[1]"), "GERAR") Then
                sessao.findById("wnd[1]").sendVKey 0
            End If

            AguardarSAP sessao
            Exit Sub
        End If

        Pausar 1
    Loop

    Err.Raise vbObjectError + 4300, "PreencherDialogoExportacao", "A janela de geracao do arquivo nao foi localizada."
End Sub

Private Function PressionarBotaoPorTexto(ByVal componente As Object, ByVal textoProcurado As String) As Boolean
    Dim i As Long, filho As Object, texto As String, tipo As String

    On Error Resume Next
    tipo = CStr(componente.Type)
    texto = UCase$(Trim$(CStr(componente.Text)))
    On Error GoTo 0

    If (tipo = "GuiButton" Or InStr(1, tipo, "Button", vbTextCompare) > 0) And InStr(1, texto, UCase$(textoProcurado), vbTextCompare) > 0 Then
        componente.Press
        PressionarBotaoPorTexto = True
        Exit Function
    End If

    On Error Resume Next
    For i = 0 To componente.Children.Count - 1
        Set filho = Nothing
        Set filho = componente.Children(i)
        If Not filho Is Nothing Then
            If PressionarBotaoPorTexto(filho, textoProcurado) Then
                PressionarBotaoPorTexto = True
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
End Function

Private Function ImportarCentro(ByVal wsOrigem As Worksheet, ByVal wsDestino As Worksheet, ByVal centro As String, ByVal limparDados As Boolean) As Long
    Dim cabecalho As Long, ultimaLinha As Long, ultimaColuna As Long
    Dim colunaEstoque As Long, linhaOrigem As Long, linhaDestino As Long
    Dim colunaMaterial As Long, estoque As Double, quantidade As Long, dados As Variant

    cabecalho = LocalizarCabecalho(wsOrigem)
    ultimaLinha = UltimaLinhaWorksheet(wsOrigem)
    ultimaColuna = UltimaColunaWorksheet(wsOrigem)

    If cabecalho = 0 Then Err.Raise vbObjectError + 4400, "ImportarCentro", "O cabecalho da exportacao nao foi localizado."
    If ultimaColuna < COLUNAS_EXPORTACAO Then Err.Raise vbObjectError + 4401, "ImportarCentro", "A exportacao deve possuir 14 colunas, de Tipo de material ate Moeda."

    colunaEstoque = LocalizarColuna(wsOrigem, cabecalho, "Estoque total")
    If colunaEstoque = 0 Then Err.Raise vbObjectError + 4402, "ImportarCentro", "A coluna Estoque total nao foi localizada."

    colunaMaterial = LocalizarColunaMaterial(wsOrigem, cabecalho)
    If colunaMaterial = 0 Then Err.Raise vbObjectError + 4403, "ImportarCentro", "A coluna Material nao foi localizada."

    If limparDados Then
        LimparBase wsDestino
    End If
    linhaDestino = UltimaLinhaDadosBase(wsDestino) + 1
    If linhaDestino < 3 Then linhaDestino = 3

    For linhaOrigem = cabecalho + 1 To ultimaLinha
        estoque = NumeroSeguro(wsOrigem.Cells(linhaOrigem, colunaEstoque).Value)

        If estoque > 0 And Len(Trim$(CStr(wsOrigem.Cells(linhaOrigem, colunaMaterial).Value))) > 0 Then
            dados = wsOrigem.Cells(linhaOrigem, 1).Resize(1, COLUNAS_EXPORTACAO).Value2
            wsDestino.Cells(linhaDestino, 1).Resize(1, COLUNAS_EXPORTACAO).Value2 = dados
            wsDestino.Cells(linhaDestino, 3).Value = NumeroMaterial(wsDestino.Cells(linhaDestino, 3).Value)
            wsDestino.Cells(linhaDestino, COLUNA_CENTRO).Value = centro
            linhaDestino = linhaDestino + 1
            quantidade = quantidade + 1
        End If
    Next linhaOrigem

    ImportarCentro = quantidade
End Function

Private Function NumeroSeguro(ByVal valor As Variant) As Double
    Dim texto As String

    If IsNumeric(valor) Then
        NumeroSeguro = CDbl(valor)
        Exit Function
    End If

    texto = Trim$(CStr(valor))
    If Len(texto) = 0 Then Exit Function
    texto = Replace(texto, ".", "")
    texto = Replace(texto, ",", Application.DecimalSeparator)

    If IsNumeric(texto) Then NumeroSeguro = CDbl(texto)
End Function

Private Sub LimparBase(ByVal ws As Worksheet)
    Dim ultima As Long
    ultima = UltimaLinhaWorksheet(ws)
    If ultima >= 3 Then ws.Range(ws.Cells(3, 1), ws.Cells(ultima, ULTIMA_COLUNA_FORMULA)).Delete Shift:=xlUp
End Sub

Private Function NumeroMaterial(ByVal valor As Variant) As Variant
    Dim texto As String

    texto = Trim$(CStr(valor))
    If Len(texto) = 0 Then Exit Function
    texto = Replace(texto, ".", "")
    texto = Replace(texto, ",", "")

    If IsNumeric(texto) Then
        NumeroMaterial = CDbl(texto)
    Else
        NumeroMaterial = valor
    End If
End Function

Private Sub RemoverCentroDaBase(ByVal ws As Worksheet, ByVal centro As String)
    Dim r As Long
    For r = UltimaLinhaDadosBase(ws) To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, COLUNA_CENTRO).Value)), centro, vbTextCompare) = 0 Then ws.Rows(r).Delete
    Next r
End Sub

Private Sub FinalizarBase(ByVal ws As Worksheet)
    Dim ultima As Long, ultimaUsada As Long, primeiraLinhaLimpeza As Long
    Dim modelo As Range, destino As Range

    ultima = UltimaLinhaDadosBase(ws)
    If ultima < 2 Then ultima = 2
    ultimaUsada = UltimaLinhaWorksheet(ws)

    DefinirModeloFormulas ws
    Set modelo = ws.Range(ws.Cells(2, PRIMEIRA_COLUNA_FORMULA), ws.Cells(2, ULTIMA_COLUNA_FORMULA))
    If ultima >= 3 Then
        Set destino = ws.Range(ws.Cells(3, PRIMEIRA_COLUNA_FORMULA), ws.Cells(ultima, ULTIMA_COLUNA_FORMULA))
        modelo.Copy
        destino.PasteSpecial xlPasteFormulas
        destino.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    If ultimaUsada > ultima Then
        primeiraLinhaLimpeza = ultima + 1
        If primeiraLinhaLimpeza < 3 Then primeiraLinhaLimpeza = 3
        If primeiraLinhaLimpeza <= ultimaUsada Then
            ws.Range(ws.Cells(primeiraLinhaLimpeza, PRIMEIRA_COLUNA_FORMULA), ws.Cells(ultimaUsada, ULTIMA_COLUNA_FORMULA)).ClearContents
        End If
    End If

    ws.Calculate
End Sub

Private Sub DefinirModeloFormulas(ByVal ws As Worksheet)
    With ws
        .Cells(2, 16).FormulaLocal = "=SE(E($H2=0;$J2=0);""-"";PROCV($C2;'https://voestalpine.sharepoint.com/Processos/Bohler/Bohler/10. DIVERSOS/Cadastro/[CadastroSAPxOmega_MM60.xlsx]MM60Rev'!$A:$Z;7;0))"
        .Cells(2, 17).FormulaLocal = "=SE(E($H2=0;$J2=0);""-"";(PROCV($C2;'V:\Processos\Bohler\Bohler\10. DIVERSOS\Cadastro\[CadastroSAPxOmega_MM60.xlsx]MM60Rev'!$A:$B;2;0)))"
        .Cells(2, 18).FormulaLocal = "=SE(E($H2=0;$J2=0);""-"";PROCV($C2;'V:\Processos\Bohler\Bohler\10. DIVERSOS\Cadastro\[CadastroSAPxOmega_MM60.xlsx]MM60Rev'!$A:$W;23;0))"
        .Cells(2, 19).FormulaLocal = "=SEERRO(SE(OU(P2=""Fita"";P2=""Foil"");PROCV(C2;'https://voestalpine.sharepoint.com/Processos/Bohler/Bohler/10. DIVERSOS/Cadastro/[CadastroSAPxOmega_MM60.xlsx]MM60Rev'!$A:$J;10;0)*H2;SE(I2=""kg"";H2;""0""));0)"
        .Cells(2, 20).FormulaLocal = "=SE(E($H2=0;$J2=0);""-"";PROCV($C2;'V:\Processos\Bohler\Bohler\10. DIVERSOS\Cadastro\[CadastroSAPxOmega_MM60.xlsx]MM60Rev'!$A:$W;20;0))"
        .Cells(2, 21).FormulaLocal = "=C2&O2"
        .Cells(2, 22).FormulaLocal = "=C2&R2&P2"
        .Cells(2, 23).FormulaLocal = "=SEERRO(J2/S2;"""")"
    End With
End Sub

Private Function AguardarExportacao(ByVal caminho As String, ByVal arquivo As String, ByVal limite As Long, ByVal antes As Object) As Workbook
    Dim inicio As Date, wb As Workbook, encontrado As String
    Dim tamAnt As Double, tamAtual As Double, estavel As Long

    inicio = Now
    Do While DateDiff("s", inicio, Now) < limite
        DoEvents

        Set wb = LocalizarNovoWorkbook(antes)
        If Not wb Is Nothing Then Set AguardarExportacao = wb: Exit Function

        Set wb = WorkbookAberto(arquivo)
        If Not wb Is Nothing Then Set AguardarExportacao = wb: Exit Function

        encontrado = LocalizarArquivo(caminho, arquivo, inicio)
        If Len(encontrado) > 0 Then
            tamAtual = TamanhoArquivo(encontrado)
            If tamAtual > 0 And tamAtual = tamAnt Then estavel = estavel + 1 Else tamAnt = tamAtual: estavel = 0

            If estavel >= 1 Then
                On Error Resume Next
                Set wb = Workbooks.Open(Filename:=encontrado, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False, Notify:=False)
                On Error GoTo 0
                If Not wb Is Nothing Then Set AguardarExportacao = wb: Exit Function
            End If
        End If

        Pausar 1
    Loop
End Function

Private Function CapturarWorkbooksAbertos() As Object
    Dim d As Object, wb As Workbook
    Set d = CreateObject("Scripting.Dictionary")
    For Each wb In Application.Workbooks
        If Not d.Exists(LCase$(wb.Name)) Then d.Add LCase$(wb.Name), True
    Next wb
    Set CapturarWorkbooksAbertos = d
End Function

Private Function LocalizarNovoWorkbook(ByVal antes As Object) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If Not antes.Exists(LCase$(wb.Name)) Then
                If Application.CountA(wb.Worksheets(1).Cells) > 0 Then Set LocalizarNovoWorkbook = wb: Exit Function
            End If
        End If
    Next wb
End Function

Private Function LocalizarArquivo(ByVal caminhoEsperado As String, ByVal arquivoEsperado As String, ByVal inicio As Date) As String
    Dim fso As Object, pastas As Object, item As Variant, pasta As Object, arq As Object
    Dim melhor As String, melhorData As Date, nome As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(caminhoEsperado) Then LocalizarArquivo = caminhoEsperado: Exit Function

    Set pastas = CreateObject("Scripting.Dictionary")
    AdicionarPasta pastas, Environ$("TEMP")
    AdicionarPasta pastas, Environ$("TMP")
    AdicionarPasta pastas, Environ$("LOCALAPPDATA") & "\Temp"
    AdicionarPasta pastas, Environ$("USERPROFILE") & "\Downloads"

    For Each item In pastas.Items
        Set pasta = Nothing
        On Error Resume Next
        Set pasta = fso.GetFolder(CStr(item))
        On Error GoTo 0

        If Not pasta Is Nothing Then
            For Each arq In pasta.Files
                nome = LCase$(CStr(arq.Name))
                If nome = LCase$(arquivoEsperado) Or Left$(nome, 7) = "export_" Or Left$(nome, 5) = "sp99_" Then
                    If arq.DateLastModified >= DateAdd("s", -5, inicio) And arq.DateLastModified >= melhorData Then
                        melhorData = arq.DateLastModified
                        melhor = arq.Path
                    End If
                End If
            Next arq
        End If
    Next item

    LocalizarArquivo = melhor
End Function

Public Sub Abrir_Estoque_SP99()
    Dim ws As Worksheet
    Set ws = ObterAba(ObterWorkbookSP99(), ABA_BASE)
    ws.Activate
    Application.GoTo ws.Range("A1"), True
End Sub

Public Sub Limpar_Log_SP99()
    Dim wb As Workbook, wsBase As Worksheet, tbl As ListObject
    Set wb = ObterWorkbookSP99()
    Set wsBase = ObterAba(wb, ABA_BASE)
    Set tbl = ObterTabelaLog(wb)
    ExcluirLinhasLog tbl
    LimparBase wsBase
    MsgBox "Log e base de estoque limpos.", vbInformation, "SP99"
End Sub

Private Sub ExcluirLinhasLog(ByVal tbl As ListObject)
    Dim indice As Long

    For indice = tbl.ListRows.Count To 1 Step -1
        tbl.ListRows(indice).Delete
    Next indice
End Sub

Private Sub AtualizarStatus(ByVal ws As Worksheet, ByVal status As String, ByVal etapa As String, ByVal centro As String, ByVal inicio As Variant, ByVal fim As Variant, ByVal linhas As Long, ByVal mensagem As String)
    ws.Range("H6").Value = status
    ws.Range("H7").Value = etapa
    ws.Range("H8").Value = centro
    ws.Range("H9").Value = Now
    ws.Range("H10").Value = inicio
    ws.Range("H11").Value = fim
    ws.Range("H12").Value = linhas
    ws.Range("H13").Value = mensagem
    ws.Range("H9:H11").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    Select Case UCase$(status)
        Case "CONCLUIDO": ws.Range("H6:J6").Interior.Color = RGB(198, 239, 206): ws.Range("H6:J6").Font.Color = RGB(0, 97, 0)
        Case "ERRO": ws.Range("H6:J6").Interior.Color = RGB(255, 199, 206): ws.Range("H6:J6").Font.Color = RGB(156, 0, 6)
        Case Else: ws.Range("H6:J6").Interior.Color = RGB(255, 235, 156): ws.Range("H6:J6").Font.Color = RGB(156, 101, 0)
    End Select
End Sub

Private Sub RegistrarLog(ByVal ws As Worksheet, ByVal execucao As String, ByVal centro As String, ByVal etapa As String, ByVal status As String, ByVal linhas As Long, ByVal arquivo As String, ByVal duracao As Long, ByVal mensagem As String, ByVal usuario As String)
    Dim tbl As ListObject, nova As ListRow
    Set tbl = ObterTabela(ws, TABELA_LOG)
    Set nova = tbl.ListRows.Add
    With nova.Range
        .Cells(1, 1).Value = Now
        .Cells(1, 2).Value = execucao
        .Cells(1, 3).Value = centro
        .Cells(1, 4).Value = etapa
        .Cells(1, 5).Value = status
        .Cells(1, 6).Value = linhas
        .Cells(1, 7).Value = arquivo
        .Cells(1, 8).Value = duracao
        .Cells(1, 9).Value = mensagem
        .Cells(1, 10).Value = usuario
        .Cells(1, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End With
End Sub

Private Function SolicitarAnoFiscal() As Long
    Dim valor As String, padrao As Long, ano As Long
    If Month(Date) >= 4 Then padrao = Year(Date) + 1 Else padrao = Year(Date)
    valor = InputBox("Informe o exercicio fiscal." & vbCrLf & "Exemplo: 2027", "SP99 - Exercicio", CStr(padrao))
    If Len(valor) = 0 Then Err.Raise vbObjectError + 4500, , "Execucao cancelada."
    If Not IsNumeric(valor) Then Err.Raise vbObjectError + 4501, , "Exercicio fiscal invalido."
    ano = CLng(Val(valor))
    If ano < 1900 Or ano > 9999 Then Err.Raise vbObjectError + 4501, , "Exercicio fiscal invalido."
    SolicitarAnoFiscal = ano
End Function

Private Function SolicitarPeriodoFiscal() As Long
    Dim valor As String, padrao As Long
    If Month(Date) >= 4 Then padrao = Month(Date) - 3 Else padrao = Month(Date) + 9
    valor = InputBox("Informe o periodo fiscal." & vbCrLf & "01 = abril ... 12 = marco", "SP99 - Periodo", Format$(padrao, "00"))
    If Len(valor) = 0 Then Err.Raise vbObjectError + 4510, , "Execucao cancelada."
    If CLng(Val(valor)) < 1 Or CLng(Val(valor)) > 12 Then Err.Raise vbObjectError + 4511, , "Periodo invalido."
    SolicitarPeriodoFiscal = CLng(Val(valor))
End Function

Private Function MontarNomeArquivo(ByVal centro As String) As String
    Dim padrao As String
    padrao = CStr(ValorNome("cfgSP99_PadraoArquivo", "SP99_{CENTRO}_{AAAAMMDD_HHMMSS}.xlsx"))
    padrao = Replace(padrao, "{CENTRO}", centro, 1, -1, vbTextCompare)
    padrao = Replace(padrao, "{AAAAMMDD_HHMMSS}", Format$(Now, "yyyymmdd_hhnnss"), 1, -1, vbTextCompare)
    If InStrRev(padrao, ".") = 0 Then padrao = padrao & ".xlsx"
    MontarNomeArquivo = padrao
End Function

Private Function ResolverPastaExportacao() As String
    Dim pasta As String
    pasta = CStr(ValorNome("cfgSP99_PastaExportacao", "%TEMP%"))
    pasta = Replace(pasta, "%TEMP%", Environ$("TEMP"), 1, -1, vbTextCompare)
    pasta = Replace(pasta, "%TMP%", Environ$("TMP"), 1, -1, vbTextCompare)
    pasta = Replace(pasta, "%LOCALAPPDATA%", Environ$("LOCALAPPDATA"), 1, -1, vbTextCompare)
    pasta = Replace(pasta, "%USERPROFILE%", Environ$("USERPROFILE"), 1, -1, vbTextCompare)
    If Len(Trim$(pasta)) = 0 Then pasta = Environ$("TEMP")
    If Right$(pasta, 1) <> "\" Then pasta = pasta & "\"
    ResolverPastaExportacao = pasta
End Function

Private Function ObterSessaoSAP() As Object
    Dim sapAuto As Object, sapApp As Object, conexao As Object, sessao As Object
    Dim i As Long, j As Long

    On Error Resume Next
    Set sapAuto = GetObject("SAPGUI")
    Set sapApp = sapAuto.GetScriptingEngine
    Set conexao = sapApp.Children(0)
    Set sessao = conexao.Children(0)
    On Error GoTo 0

    If Not sessao Is Nothing Then Set ObterSessaoSAP = sessao: Exit Function

    On Error Resume Next
    For i = 0 To sapApp.Children.Count - 1
        Set conexao = sapApp.Children(i)
        For j = 0 To conexao.Children.Count - 1
            Set sessao = conexao.Children(j)
            If Not sessao Is Nothing Then Set ObterSessaoSAP = sessao: On Error GoTo 0: Exit Function
        Next j
    Next i
    On Error GoTo 0

    Err.Raise vbObjectError + 4600, , "Nenhuma sessao SAP foi disponibilizada."
End Function

Private Sub AguardarSAP(ByVal sessao As Object)
    Dim inicio As Date, timeout As Long
    inicio = Now
    timeout = CLng(Val(CStr(ValorNome("cfgSP99_Timeout", TIMEOUT_PADRAO))))
    If timeout <= 0 Then timeout = TIMEOUT_PADRAO
    Do While sessao.Busy
        DoEvents
        If DateDiff("s", inicio, Now) > timeout Then Err.Raise vbObjectError + 4610, , "O SAP ultrapassou o tempo limite."
    Loop
End Sub

Private Sub DefinirTextoSAP(ByVal sessao As Object, ByVal id As String, ByVal valor As String)
    If Not ExisteObjetoSAP(sessao, id) Then Err.Raise vbObjectError + 4620, , "Campo SAP nao localizado:" & vbCrLf & id
    sessao.findById(id).Text = valor
    sessao.findById(id).SetFocus
    On Error Resume Next
    sessao.findById(id).CaretPosition = Len(valor)
    On Error GoTo 0
End Sub

Private Function ExisteObjetoSAP(ByVal sessao As Object, ByVal id As String) As Boolean
    Dim obj As Object
    On Error Resume Next
    Set obj = sessao.findById(id, False)
    ExisteObjetoSAP = Not obj Is Nothing
    On Error GoTo 0
End Function

Private Function LocalizarCabecalho(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long, encontrados As Long, valor As String
    For r = 1 To WorksheetFunction.Min(30, ws.UsedRange.Rows.Count)
        encontrados = 0
        For c = 1 To WorksheetFunction.Min(25, ws.UsedRange.Columns.Count)
            valor = UCase$(Trim$(CStr(ws.Cells(r, c).Value)))
            If valor = "TIPO DE MATERIAL" Or valor = "MATERIAL" Or valor = "ESTOQUE TOTAL" Or valor = "MOEDA" Then encontrados = encontrados + 1
        Next c
        If encontrados >= 3 Then LocalizarCabecalho = r: Exit Function
    Next r
End Function

Private Function LocalizarColuna(ByVal ws As Worksheet, ByVal linhaCabecalho As Long, ByVal titulo As String) As Long
    Dim c As Long
    For c = 1 To UltimaColunaWorksheet(ws)
        If StrComp(Trim$(CStr(ws.Cells(linhaCabecalho, c).Value)), titulo, vbTextCompare) = 0 Then LocalizarColuna = c: Exit Function
    Next c
End Function

Private Function LocalizarColunaMaterial(ByVal ws As Worksheet, ByVal linhaCabecalho As Long) As Long
    Dim c As Long, titulo As String

    For c = 1 To UltimaColunaWorksheet(ws)
        titulo = UCase$(Trim$(CStr(ws.Cells(linhaCabecalho, c).Value)))
        If titulo = "MATERIAL" Or InStr(1, titulo, "MATERIAL", vbTextCompare) > 0 Then
            If titulo <> "TIPO DE MATERIAL" Then
                LocalizarColunaMaterial = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function LocalizarAbaExportada(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet, maior As Double, qtd As Double
    For Each ws In wb.Worksheets
        qtd = Application.CountA(ws.Cells)
        If qtd > maior Then maior = qtd: Set LocalizarAbaExportada = ws
    Next ws
End Function

Private Function WorkbookAberto(ByVal nome As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nome, vbTextCompare) = 0 Then Set WorkbookAberto = wb: Exit Function
    Next wb
End Function

Private Function UltimaLinhaDadosBase(ByVal ws As Worksheet) As Long
    Dim c As Long, maior As Long, atual As Long
    For c = 1 To COLUNA_CENTRO
        atual = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If atual > maior Then maior = atual
    Next c
    If maior < 1 Then maior = 1
    UltimaLinhaDadosBase = maior
End Function

Private Function UltimaLinhaWorksheet(ByVal ws As Worksheet) As Long
    Dim cel As Range
    Set cel = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByRows, xlPrevious)
    If Not cel Is Nothing Then UltimaLinhaWorksheet = cel.Row
End Function

Private Function UltimaColunaWorksheet(ByVal ws As Worksheet) As Long
    Dim cel As Range
    Set cel = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, xlPart, xlByColumns, xlPrevious)
    If Not cel Is Nothing Then UltimaColunaWorksheet = cel.Column
End Function

Private Function ObterAba(ByVal wb As Workbook, ByVal nome As String) As Worksheet
    On Error Resume Next
    Set ObterAba = wb.Worksheets(nome)
    On Error GoTo 0
    If ObterAba Is Nothing Then Err.Raise vbObjectError + 4700, , "A aba '" & nome & "' nao foi encontrada."
End Function

Private Sub PrepararExecucao(ByRef wb As Workbook, ByRef wsPainel As Worksheet, ByRef wsBase As Worksheet, ByRef anoFiscal As Long, ByRef periodo As Long, ByRef inicio As Date, ByRef execucao As String, ByRef calcAnt As XlCalculation, ByRef eventosAnt As Boolean, ByRef telaAnt As Boolean, ByRef alertasAnt As Boolean)
    Set wb = ObterWorkbookSP99()
    Set wsPainel = ObterAba(wb, ABA_PAINEL)
    Set wsBase = ObterAba(wb, ABA_BASE)

    calcAnt = Application.Calculation
    eventosAnt = Application.EnableEvents
    telaAnt = Application.ScreenUpdating
    alertasAnt = Application.DisplayAlerts
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    anoFiscal = SolicitarAnoFiscal()
    periodo = SolicitarPeriodoFiscal()
    inicio = Now
    execucao = Format$(inicio, "yyyymmdd_hhnnss")
End Sub

Private Function ObterWorkbookSP99() As Workbook
    Dim wb As Workbook

    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If WorkbookTemAba(wb, ABA_BASE) Then Set ObterWorkbookSP99 = wb: Exit Function

    If WorkbookTemAba(ThisWorkbook, ABA_BASE) Then Set ObterWorkbookSP99 = ThisWorkbook: Exit Function

    For Each wb In Application.Workbooks
        If WorkbookTemAba(wb, ABA_BASE) Then Set ObterWorkbookSP99 = wb: Exit Function
    Next wb

    Err.Raise vbObjectError + 4720, , "Nenhum arquivo aberto possui a aba '" & ABA_BASE & "'."
End Function

Private Function WorkbookTemAba(ByVal wb As Workbook, ByVal nome As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = wb.Worksheets(nome)
    On Error GoTo 0
    WorkbookTemAba = Not ws Is Nothing
End Function

Private Function ObterTabelaLog(ByVal wb As Workbook) As ListObject
    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = wb.Worksheets(ABA_PAINEL).ListObjects(TABELA_LOG)
    On Error GoTo 0

    If tbl Is Nothing Then Err.Raise vbObjectError + 4730, , "A tabela '" & TABELA_LOG & "' nao foi encontrada no arquivo aberto."
    Set ObterTabelaLog = tbl
End Function

Private Function ObterTabela(ByVal ws As Worksheet, ByVal nome As String) As ListObject
    On Error Resume Next
    Set ObterTabela = ws.ListObjects(nome)
    On Error GoTo 0
    If ObterTabela Is Nothing Then Err.Raise vbObjectError + 4710, , "A tabela '" & nome & "' nao foi encontrada em '" & ws.Name & "'."
End Function

Private Function ValorTabela(ByVal linha As ListRow, ByVal tabela As ListObject, ByVal nomeColuna As String) As Variant
    ValorTabela = linha.Range.Cells(1, tabela.ListColumns(nomeColuna).Index).Value
End Function

Private Function ValorNome(ByVal nome As String, ByVal padrao As Variant) As Variant
    Dim rg As Range, wb As Workbook
    On Error Resume Next
    Set wb = ObterWorkbookSP99()
    Set rg = wb.Names(nome).RefersToRange
    On Error GoTo 0
    If rg Is Nothing Then
        ValorNome = padrao
    ElseIf Len(Trim$(CStr(rg.Value))) = 0 Then
        ValorNome = padrao
    Else
        ValorNome = rg.Value
    End If
End Function

Private Function RemoverBarraFinal(ByVal pasta As String) As String
    Do While Len(pasta) > 0 And (Right$(pasta, 1) = "\" Or Right$(pasta, 1) = "/")
        pasta = Left$(pasta, Len(pasta) - 1)
    Loop
    RemoverBarraFinal = pasta
End Function

Private Sub AdicionarPasta(ByVal d As Object, ByVal pasta As String)
    If Len(Trim$(pasta)) > 0 Then If Not d.Exists(LCase$(pasta)) Then d.Add LCase$(pasta), pasta
End Sub

Private Function TamanhoArquivo(ByVal caminho As String) As Double
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(caminho) Then TamanhoArquivo = CDbl(fso.GetFile(caminho).Size)
    On Error GoTo 0
End Function

Private Sub Pausar(ByVal segundos As Long)
    Dim fim As Date
    fim = DateAdd("s", segundos, Now)
    Do While Now < fim
        DoEvents
    Loop
End Sub

Private Sub RestaurarExcel(ByVal calc As XlCalculation, ByVal eventos As Boolean, ByVal tela As Boolean, ByVal alertas As Boolean)
    Application.StatusBar = False
    Application.Calculation = calc
    Application.EnableEvents = eventos
    Application.ScreenUpdating = tela
    Application.DisplayAlerts = alertas
    Application.CutCopyMode = False
End Sub


