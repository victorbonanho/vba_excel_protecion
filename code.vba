Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Dim username As String
    Dim usuariosPermitidos As New Collection
    Dim usuarioPermitido As Boolean
    Dim usuario As Variant
    Dim nomeDaAba As String

    ' Captura o nome do usuário logado no Excel
    username = Application.username

    ' Captura o nome da aba ativa
    nomeDaAba = ThisWorkbook.Sheets(ActiveSheet.Name).Name
    
    Debug.Print "Usuário atual: " & username
    Debug.Print "Nome da aba: " & nomeDaAba
    
    ' Adiciona usuários padrão que têm permissão para acessar todas as planilhas
    usuariosPermitidos.Add "Vanessa | Grupo Araujo Engenharia"
    usuariosPermitidos.Add "Ivete Vespasiano | Grupo Araujo Engenharia"
    ' usuariosPermitidos.Add "Victor Timotti | Grupo Araujo Engenharia"
    
    ' Define a lista de usuários permitidos para cada aba
    Select Case nomeDaAba
        Case "Gustavo"
            usuariosPermitidos.Add "Gustavo Migray | Grupo Araújo Engenharia"
        Case "Andre"
            usuariosPermitidos.Add "Andre Padua | Grupo Araujo Engenharia"
        Case "Marco"
            usuariosPermitidos.Add "Marco Oliveira | Grupo Araujo Engenharia"
        Case "João"
            usuariosPermitidos.Add "Joao Paulo | Grupo Araujo Engenharia"
        Case "Fernanda"
            usuariosPermitidos.Add "Fernanda Bueno"
        Case "Renato"
            usuariosPermitidos.Add "Renato Carvalho | Grupo Araujo Engenharia"
        Case "Marcos"
            usuariosPermitidos.Add "Renato Carvalho | Grupo Araujo Engenharia"
        Case "Cleo"
            usuariosPermitidos.Add "Qualidade | Grupo Araujo Engenharia"
        Case "Vanessa"
            usuariosPermitidos.Add "Vanessa | Grupo Araujo Engenharia"
    End Select

    ' Inicializa a variável de controle
    usuarioPermitido = False

    ' Verifica se o usuário está na lista de permitidos para a aba
    For Each usuario In usuariosPermitidos
        If usuario = username Then
            usuarioPermitido = True
            Exit For
        End If
    Next usuario

    ' Desbloqueia a aba para garantir acesso a todos os usuários
    With ThisWorkbook.Sheets(nomeDaAba)
        .Unprotect ' Desprotege a planilha

        ' Se o usuário for permitido, mantém a aba desprotegida
        If usuarioPermitido Then
            .Cells.Locked = False
        Else
            ' Se o usuário não estiver na lista, protege a planilha
            Debug.Print "Usuário não permitido: " & username
            .Cells.Locked = True
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
            ' MsgBox "Você não tem permissão para editar esta planilha!"
        End If
    End With
End Sub

