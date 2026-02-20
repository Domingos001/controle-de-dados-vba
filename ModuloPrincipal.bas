'======================================================
' MACRO DE ATUALIZAÇÃO DA LISTA MESTRA (AMOSTRAS)
'======================================================
Sub AtualizarListaMestra()
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet, wsDestino As Worksheet, wsTemp As Worksheet
    Dim caminhoRede As String, caminhoLocalTemp As String
    Dim nomePlanilha As String
    Dim senhaLocal As String, senhaMestre As String
    Dim lastRowOrigem As Long, lastRowDestino As Long, lastRowNovo As Long
    Dim i As Long
    Dim btn As Button
    Dim celula As Range
    Dim Top As Double, Left As Double, Height As Double, Width As Double
    
    caminhoRede = "\\s01\Controle_Padroes\FORM497 - Controle de Padrões Referência e Amostras Padrão de Clientes.xlsx"
    nomePlanilha = "Amostra Referência e Padrão"
    senhaLocal = "1234"
    senhaMestre = "KISGQ"
    caminhoLocalTemp = Environ("TEMP") & "\temp_form497.xlsx"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Dir(caminhoRede) = "" Then GoTo ErroNaoEncontrado
    If MsgBox("Isso irá substituir os dados das colunas A-F pelos dados do arquivo mestre (Original)." & _
              vbCrLf & vbCrLf & "Deseja continuar?", vbYesNo + vbQuestion, "Confirmar Atualização") = vbNo Then GoTo LimpezaFinal
              
    On Error GoTo ErroDeAcesso
    If Dir(caminhoLocalTemp) <> "" Then Kill caminhoLocalTemp
    FileCopy caminhoRede, caminhoLocalTemp
    On Error GoTo ErroGeral
    
    Set wbOrigem = Workbooks.Open(Filename:=caminhoLocalTemp, ReadOnly:=True, UpdateLinks:=False, Password:=senhaMestre)
    Set wsOrigem = wbOrigem.Sheets(nomePlanilha)
    lastRowOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "TEMP_UPDATE"
    
    If lastRowOrigem >= 5 Then
        wsOrigem.Range("A5:F" & lastRowOrigem).Copy Destination:=wsTemp.Range("A1")
    End If
    
    wbOrigem.Close SaveChanges:=False
    Kill caminhoLocalTemp
    
    Set wsDestino = ThisWorkbook.Sheets(nomePlanilha)
    wsDestino.Unprotect Password:=senhaLocal
    
    Dim lastRowAbsoluto As Long
    Dim foundGB As Boolean
    lastRowAbsoluto = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row
    foundGB = False
    lastRowDestino = 1
    If lastRowAbsoluto >= 2 Then
        For i = 2 To lastRowAbsoluto
            If VBA.Strings.Left(LCase(CStr(wsDestino.Cells(i, 1).Value)), 2) = "gb" Then
                lastRowDestino = i - 1
                foundGB = True
                Exit For
            End If
        Next i
        If Not foundGB Then lastRowDestino = lastRowAbsoluto
    End If
    If lastRowDestino < 2 Then lastRowDestino = 1 
    
    If lastRowDestino >= 2 Then
        wsDestino.Range("A2:G" & lastRowDestino).ClearContents
        wsDestino.Range("A2:G" & lastRowDestino).ClearFormats
    End If
    
    Dim totalLinhasCopiadas As Long
    totalLinhasCopiadas = 0
    If wsTemp.Range("A1").Value <> "" Then
        totalLinhasCopiadas = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        wsTemp.Range("A1:F" & totalLinhasCopiadas).Copy Destination:=wsDestino.Range("A2")
    End If
    Application.CutCopyMode = False
    
    lastRowNovo = (2 - 1) + totalLinhasCopiadas
    
    If lastRowNovo >= 2 Then
        For i = 2 To lastRowNovo
            Set celula = wsDestino.Range("G" & i)
            Top = celula.Top + 1
            Left = celula.Left + 1
            Height = celula.Height - 2
            Width = celula.Width - 2
            
            Set btn = wsDestino.Buttons.Add(Left, Top, Width, Height)
            With btn
                .Caption = "Registrar Saída"
                .OnAction = "RegistrarSaida"
                .Name = "btnSaida_" & i
            End With
        Next i
        ' Aplica a fonte padrão
        wsDestino.Range("A2:G" & lastRowNovo).Font.Name = "Aptos Narrow"
    End If
    
    wsTemp.Delete
    wsDestino.Protect Password:=senhaLocal, UserInterfaceOnly:=True, AllowFiltering:=True
    
    MsgBox "A lista de amostras e padrões foi atualizada com sucesso!", vbInformation, "Atualização Concluída"
    GoTo LimpezaFinal

ErroNaoEncontrado:
    MsgBox "ERRO: Não foi possível encontrar o arquivo mestre (Original) na rede.", vbCritical, "Falha na Atualização"
    GoTo LimpezaFinal
ErroDeAcesso:
    MsgBox "ERRO DE ACESSO: Acesso negado pelo TI/Antivírus.", vbCritical, "Bloqueio de Segurança"
    GoTo LimpezaFinal
ErroGeral:
    MsgBox "Ocorreu um erro inesperado: " & Err.Description, vbCritical, "Erro de Macro"
LimpezaFinal:
    If Dir(caminhoLocalTemp) <> "" Then Kill caminhoLocalTemp
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


'======================================================
' MACRO DE ATUALIZAÇÃO DA LISTA DE GABARITOS (GBs)
'======================================================
Sub AtualizarListaGBs()
    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet, wsDestino As Worksheet, wsTemp As Worksheet
    Dim caminhoRede_GB As String, caminhoLocalTemp_GB As String
    Dim nomePlanilha_GB As String
    Dim senhaLocal As String
    Dim lastRowOrigem As Long, lastRowFimGB As Long, lastRowNovo As Long
    Dim i As Long
    Dim linhaInicioGB As Long
    Dim btn As Button
    Dim celula As Range
    Dim Top As Double, Left As Double, Height As Double, Width As Double
    
    caminhoRede_GB = "\\s01\Calibração de Instrumentos\FORM503 - Controle de Verificação dos Gabaritos - 2025.xls"
    nomePlanilha_GB = "GABARITOS"
    senhaLocal = "1234"
    caminhoLocalTemp_GB = Environ("TEMP") & "\temp_form503_gb.xls"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Dir(caminhoRede_GB) = "" Then GoTo ErroNaoEncontrado_GB
    If MsgBox("Isso irá substituir TODOS os dados da lista de GABARITOS (GBs) pelos dados do arquivo da rede." & _
              vbCrLf & vbCrLf & "Deseja continuar?", vbYesNo + vbQuestion, "Confirmar Atualização de GBs") = vbNo Then GoTo LimpezaFinal_GB
    
    On Error GoTo ErroDeAcesso_GB
    If Dir(caminhoLocalTemp_GB) <> "" Then Kill caminhoLocalTemp_GB
    FileCopy caminhoRede_GB, caminhoLocalTemp_GB
    On Error GoTo ErroGeral_GB
    
    Set wbOrigem = Workbooks.Open(Filename:=caminhoLocalTemp_GB, ReadOnly:=True, UpdateLinks:=False)
    Set wsOrigem = wbOrigem.Sheets(nomePlanilha_GB)
    lastRowOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
    Set wsTemp = ThisWorkbook.Sheets.Add
    wsTemp.Name = "TEMP_UPDATE_GB"
    
    ' Inversão das colunas feita diretamente na cópia
    If lastRowOrigem >= 5 Then
        wsOrigem.Range("A5:A" & lastRowOrigem).Copy Destination:=wsTemp.Range("A1")
        wsOrigem.Range("B5:B" & lastRowOrigem).Copy Destination:=wsTemp.Range("C1")
        wsOrigem.Range("C5:C" & lastRowOrigem).Copy Destination:=wsTemp.Range("B1")
    End If
    
    wbOrigem.Close SaveChanges:=False
    Kill caminhoLocalTemp_GB
    
    Set wsDestino = ThisWorkbook.Sheets("Amostra Referência e Padrão")
    wsDestino.Unprotect Password:=senhaLocal
    
    Dim lastRowAbsoluto As Long
    On Error Resume Next
    lastRowAbsoluto = wsDestino.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If lastRowAbsoluto = 0 Then lastRowAbsoluto = 1
    
    linhaInicioGB = 0
    If lastRowAbsoluto >= 2 Then
        For i = 2 To lastRowAbsoluto
            If VBA.Strings.Left(LCase(CStr(wsDestino.Cells(i, 1).Value)), 2) = "gb" Then
                linhaInicioGB = i
                Exit For
            End If
        Next i
    End If
    
    If linhaInicioGB = 0 Then GoTo ErroBlocoGBNaoEncontrado
    
    lastRowFimGB = lastRowAbsoluto
    
    ' Apaga as LINHAS INTEIRAS do bloco GB antigo (Limpeza garantida)
    If lastRowFimGB >= linhaInicioGB Then
        wsDestino.Rows(linhaInicioGB & ":" & lastRowFimGB).Delete
    End If
    
    Dim totalLinhasCopiadas As Long
    totalLinhasCopiadas = 0
    If wsTemp.Range("A1").Value <> "" Then
        totalLinhasCopiadas = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
        wsTemp.Range("A1:C" & totalLinhasCopiadas).Copy Destination:=wsDestino.Range("A" & linhaInicioGB)
    End If
    Application.CutCopyMode = False
    
    Dim lastRowNovoGB As Long
    lastRowNovoGB = (linhaInicioGB - 1) + totalLinhasCopiadas
    
    If totalLinhasCopiadas > 0 Then
        With wsDestino.Range("D" & linhaInicioGB & ":F" & lastRowNovoGB)
            .Value = " - "
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If

    If totalLinhasCopiadas > 0 Then
        For i = linhaInicioGB To lastRowNovoGB
            Set celula = wsDestino.Range("G" & i)
            Top = celula.Top + 1
            Left = celula.Left + 1
            Height = celula.Height - 2
            Width = celula.Width - 2
            
            Set btn = wsDestino.Buttons.Add(Left, Top, Width, Height)
            With btn
                .Caption = "Registrar Saída"
                .OnAction = "RegistrarSaida"
                .Name = "btnSaida_" & i
            End With
        Next i
    End If
    
    ' Aplica a fonte e bordas em tudo
    If totalLinhasCopiadas > 0 Then
        With wsDestino.Range("A" & linhaInicioGB & ":F" & lastRowNovoGB)
            .Font.Name = "Aptos Narrow"
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If
    
    wsTemp.Delete
    wsDestino.Protect Password:=senhaLocal, UserInterfaceOnly:=True, AllowFiltering:=True
    
    MsgBox "A lista de GABARITOS (GBs) foi atualizada com sucesso!" & vbCrLf & _
           "Colunas B e C foram invertidas.", vbInformation, "Atualização Concluída"
    GoTo LimpezaFinal_GB

ErroNaoEncontrado_GB:
    MsgBox "ERRO: Não foi possível encontrar o arquivo mestre de GBs na rede.", vbCritical, "Falha na Atualização"
    GoTo LimpezaFinal_GB
ErroBlocoGBNaoEncontrado:
    wsTemp.Delete
    wsDestino.Protect Password:=senhaLocal, UserInterfaceOnly:=True, AllowFiltering:=True
    MsgBox "ERRO: Não foi possível encontrar a tabela de 'GBs' na sua planilha para substituir.", vbCritical
    GoTo LimpezaFinal_GB
ErroDeAcesso_GB:
    MsgBox "ERRO DE ACESSO (GBs): Acesso negado pelo TI.", vbCritical, "Bloqueio de Segurança"
    GoTo LimpezaFinal_GB
ErroGeral_GB:
    MsgBox "Ocorreu um erro inesperado (GBs): " & Err.Description, vbCritical, "Erro de Macro"
LimpezaFinal_GB:
    If Dir(caminhoLocalTemp_GB) <> "" Then Kill caminhoLocalTemp_GB
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


'======================================================
' MACRO PARA REGISTRAR SAÍDA
'======================================================
Sub RegistrarSaida()
    Dim wsOrigem As Worksheet, wsDestino As Worksheet
    Dim linha As Long
    Dim ultimaLinha As Long
    Dim btn As Button
    Dim intervaloNovaLinha As Range
    Dim resposta As VbMsgBoxResult
    Dim ci As Variant, produto As Variant
    Dim senha As String
    Dim btnClicado As Object
    Dim nomeBotao As String
    
    senha = "1234"
    Set wsOrigem = Sheets("Amostra Referência e Padrão")
    Set wsDestino = Sheets("Movimentações")
    
    nomeBotao = Application.Caller
    Set btnClicado = wsOrigem.Shapes(nomeBotao)
    linha = btnClicado.TopLeftCell.Row
    
    ci = wsOrigem.Cells(linha, 1).Value
    produto = wsOrigem.Cells(linha, 3).Value
    
    If IsEmpty(ci) Or IsEmpty(produto) Then
        MsgBox "Não foi possível encontrar dados válidos nesta linha."
        Exit Sub
    End If
    
    resposta = MsgBox("Tem certeza que deseja registrar a saída deste item?" & vbCrLf & _
                      "CI: " & ci & vbCrLf & "Produto: " & produto, vbYesNo + vbQuestion, "Confirmar Saída")
    
    If resposta = vbYes Then
        wsDestino.Unprotect Password:=senha
        
        ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
        If ultimaLinha < 3 Then ultimaLinha = 3
        ultimaLinha = ultimaLinha + 1
        
        wsDestino.Cells(ultimaLinha, 1).Value = ci
        wsDestino.Cells(ultimaLinha, 2).Value = produto
        wsDestino.Cells(ultimaLinha, 3).Value = Date
        wsDestino.Cells(ultimaLinha, 4).Value = Time
        wsDestino.Cells(ultimaLinha, 3).NumberFormat = "dd/mm/yyyy"
        wsDestino.Cells(ultimaLinha, 4).NumberFormat = "hh:mm:ss"
        
        wsDestino.Cells(ultimaLinha, 5).Value = ""
        wsDestino.Cells(ultimaLinha, 6).Value = ""
        
        Set btn = wsDestino.Buttons.Add(wsDestino.Cells(ultimaLinha, 7).Left, _
                                        wsDestino.Cells(ultimaLinha, 7).Top, _
                                        wsDestino.Cells(ultimaLinha, 7).Width, _
                                        wsDestino.Cells(ultimaLinha, 7).Height)
        With btn
            .Caption = "Registrar Retorno"
            .OnAction = "RegistrarRetornoBotao"
            .Name = "btnRetorno_" & ultimaLinha
        End With
        
        Set intervaloNovaLinha = wsDestino.Range("A" & ultimaLinha & ":G" & ultimaLinha)
        With intervaloNovaLinha.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        intervaloNovaLinha.Font.Name = "Aptos Narrow"
        
        wsDestino.Activate
        wsDestino.Rows(ultimaLinha).Select
        wsDestino.Cells(ultimaLinha, 1).Activate
        
        MsgBox "Saída registrada com sucesso!"
    Else
        MsgBox "Registro de saída cancelado."
    End If
    
    wsDestino.Protect Password:=senha, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub


'======================================================
' MACRO PARA REGISTRAR RETORNO
'======================================================
Sub RegistrarRetornoBotao()
    Dim ws As Worksheet
    Dim linha As Long
    Dim nomeBotao As String
    Dim intervaloTabela As Range
    Dim senha As String
    Dim btn As Object
    Dim ci As Variant
    Dim produto As Variant
    Dim resposta As VbMsgBoxResult
    
    senha = "1234"
    Set ws = ThisWorkbook.Sheets("Movimentações")
    
    nomeBotao = Application.Caller
    Set btn = ws.Shapes(nomeBotao)
    linha = btn.TopLeftCell.Row
    
    ci = ws.Cells(linha, 1).Value
    produto = ws.Cells(linha, 2).Value
    
    resposta = MsgBox("Tem certeza que deseja registrar o RETORNO deste item?" & vbCrLf & _
                      "CI: " & ci & vbCrLf & "Produto: " & produto, vbYesNo + vbQuestion, "Confirmar Retorno")
                      
    If resposta = vbYes Then
        ws.Unprotect Password:=senha
        
        If ws.Cells(linha, 5).Value = "" And ws.Cells(linha, 6).Value = "" Then
            ws.Cells(linha, 5).Value = Date
            ws.Cells(linha, 6).Value = Time
            ws.Cells(linha, 5).NumberFormat = "dd/mm/yyyy"
            ws.Cells(linha, 6).NumberFormat = "hh:mm:ss"
            
            Set intervaloTabela = ws.Range("A" & linha & ":G" & linha)
            intervaloTabela.Interior.Color = RGB(198, 239, 206)
            intervaloTabela.Font.Name = "Aptos Narrow"
            
            MsgBox "Retorno registrado com sucesso!"
            btn.Delete
        Else
            MsgBox "O retorno para este item já foi registrado anteriormente!"
        End If
        
        ws.Protect Password:=senha, UserInterfaceOnly:=True, AllowFiltering:=True
    Else
        MsgBox "Registro de retorno cancelado."
    End If
End Sub


'======================================================
' MACRO PARA EXPORTAR O RELATÓRIO PDF MENSAL
'======================================================
Sub ExportarMovimentacoesPDF()
    Dim ws As Worksheet, tempWS As Worksheet
    Dim lastRow As Long, i As Long, destRow As Long
    Dim caminhoPasta As String, nomeArquivo As String, caminhoCompleto As String
    Dim mesRelatorio As Integer, anoRelatorio As Integer
    Dim dataSaida As Variant, dataRetorno As Variant
    Dim incluirLinha As Boolean
    Dim tituloRelatorio As String
    Dim resposta As VbMsgBoxResult
    Dim senhaLocal As String
    
    senhaLocal = "1234"
    caminhoPasta = ThisWorkbook.Path & "\Relatórios\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ws = ThisWorkbook.Sheets("Movimentações")
    mesRelatorio = Month(Date)
    anoRelatorio = Year(Date)
    nomeArquivo = "Relatorio_Movimentacoes_" & Format(Date, "mm-yyyy") & ".pdf"
    caminhoCompleto = caminhoPasta & nomeArquivo
    
    If Dir(caminhoPasta, vbDirectory) = "" Then
        On Error Resume Next
        MkDir caminhoPasta
        If Err.Number <> 0 Then
            MsgBox "Erro! Não foi possível criar a pasta de Relatórios no caminho:" & vbCrLf & caminhoPasta, vbCritical
            GoTo LimpezaFinal
        End If
        On Error GoTo 0
    End If
    
    If Dir(caminhoCompleto) <> "" Then GoTo LimpezaFinal
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set tempWS = ThisWorkbook.Sheets.Add
    
    tituloRelatorio = "Relatório Controle de amostras e GB do mês de " & Format(Date, "mmmm")
    tempWS.Range("A1").Value = tituloRelatorio
    With tempWS.Range("A1:F1")
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    ws.Range("A3:F3").Copy Destination:=tempWS.Range("A2")
    destRow = 3
    For i = 4 To lastRow
        dataSaida = ws.Cells(i, 3).Value
        dataRetorno = ws.Cells(i, 5).Value
        incluirLinha = False
        
        If IsDate(dataSaida) Then
            If Month(dataSaida) = mesRelatorio And Year(dataSaida) = anoRelatorio Then incluirLinha = True
        End If
        If Not incluirLinha And IsDate(dataRetorno) Then
            If Month(dataRetorno) = mesRelatorio And Year(dataRetorno) = anoRelatorio Then incluirLinha = True
        End If
        If Not incluirLinha And IsDate(dataSaida) Then
            If dataSaida < DateSerial(anoRelatorio, mesRelatorio, 1) Then
                If Not IsDate(dataRetorno) Or dataRetorno >= DateSerial(anoRelatorio, mesRelatorio, 1) Then incluirLinha = True
            End If
        End If
        
        If incluirLinha Then
            ws.Range("A" & i & ":F" & i).Copy Destination:=tempWS.Range("A" & destRow)
            destRow = destRow + 1
        End If
    Next i
    
    If destRow > 2 Then
        tempWS.Columns("A:F").EntireColumn.AutoFit
        tempWS.UsedRange.Font.Name = "Aptos Narrow"
        
        With tempWS.PageSetup
            .Orientation = xlPortrait
            .Zoom = 75
        End With
        
        tempWS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminhoCompleto
        tempWS.Delete
        MsgBox "Relatório do mês " & Format(Date, "mmmm") & " foi salvo com sucesso em:" & vbCrLf & caminhoCompleto
        
        resposta = MsgBox("Deseja arquivar (remover da lista) as movimentações que já foram concluídas?", vbYesNo + vbQuestion, "Arquivar Movimentações Concluídas")
        
        If resposta = vbYes Then
            ws.Unprotect Password:=senhaLocal
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            For i = lastRow To 4 Step -1
                If ws.Cells(i, 5).Value <> "" Then
                    ws.Rows(i).Delete
                End If
            Next i
            ws.Protect Password:=senhaLocal, UserInterfaceOnly:=True, AllowFiltering:=True
            MsgBox "Movimentações concluídas arquivadas."
        End If
    Else
        tempWS.Delete
        MsgBox "Nenhuma movimentação encontrada para o relatório do mês " & Format(Date, "mmmm") & "."
    End If

LimpezaFinal:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
