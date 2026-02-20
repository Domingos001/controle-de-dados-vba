Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    
    If Target.Column = 1 And Target.Row >= 2 Then
        Application.EnableEvents = False
        
        Dim celula As Range
        
        If Target.Value <> "" Then
            Set celula = Me.Range("G" & Target.Row)
            
            On Error Resume Next
            Me.Shapes("btnSaida_" & Target.Row).Delete
            On Error GoTo 0
            
            Dim Top As Double, Left As Double, Height As Double, Width As Double
            Top = celula.Top + 1
            Left = celula.Left + 1
            Height = celula.Height - 2
            Width = celula.Width - 2

            Dim btn As Button
            Set btn = Me.Buttons.Add(Left, Top, Width, Height)
            With btn
                .Caption = "Registrar Saída"
                .OnAction = "RegistrarSaida"
                .Name = "btnSaida_" & Target.Row
            End With
            
            ' Aplica a fonte na linha inteira
            Me.Rows(Target.Row).Font.Name = "Aptos Narrow"
            
        ElseIf Target.Value = "" Then
            On Error Resume Next
            Me.Shapes("btnSaida_" & Target.Row).Delete
            On Error GoTo 0
        End If
        
        Application.EnableEvents = True
    End If
End Sub
