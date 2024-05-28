Private Sub Worksheet_Change(ByVal Target As Range)
    Dim Col As Range

    ' Evita loop infinito de eventos
    Application.EnableEvents = False

    ' Ajusta a largura da coluna do intervalo modificado
    On Error GoTo ExitSub ' Garante que os eventos sejam reativados em caso de erro
    For Each Col In Target.Columns
        Col.EntireColumn.AutoFit
    Next Col

ExitSub:
    ' Habilita eventos novamente
    Application.EnableEvents = True
End Sub

