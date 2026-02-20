Private Sub Workbook_Open()
    Dim senha As String
    senha = "111111" ' <-- SENHA LOCAL
    
    ' Protege a planilha de Movimentações, permitindo macros E FILTROS
    Sheets("Movimentações").Protect Password:=senha, UserInterfaceOnly:=True, AllowFiltering:=True
    
    ' Protege a planilha principal, permitindo macros E FILTROS
    With Sheets("Amostra Referência e Padrão")
        .Protect Password:=senha, UserInterfaceOnly:=True, AllowFiltering:=True
    End With

    ' --- Lógica do Relatório em PDF ---
    If Date = DateSerial(Year(Date), Month(Date) + 1, 0) Then
        Call ExportarMovimentacoesPDF
    End If
End Sub
