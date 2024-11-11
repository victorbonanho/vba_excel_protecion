Private Sub Workbook_Open()
    ' Executa a verificação de permissões ao abrir o arquivo
    VerificarPermissoes
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Executa a verificação de permissões ao mudar de aba
    VerificarPermissoes
End Sub

Private Sub VerificarPermissoes()
    Dim username As String
    Dim usuariosPermitidosCompletos As New Collection
    Dim usuariosPermitidosNomes As New Collection
    Dim usuarioPermitido As Boolean
    Dim usuario As Variant
    Dim nomeDaAba As String

    ' Captura o nome do usuário logado no Excel
    username = Application.username

    ' Captura o nome da aba ativa
    nomeDaAba = ActiveSheet.Name
    
    Debug.Print "Usuário atual: " & username
    Debug.Print "Nome da aba: " & nomeDaAba
    
    ' Adiciona usuários padrão com nome completo
    usuariosPermitidosCompletos.Add "Vanessa | Grupo Araujo Engenharia"
    usuariosPermitidosCompletos.Add "Ivete Vespasiano | Grupo Araujo Engenharia"
    usuariosPermitidosCompletos.Add "Victor Timotti | Grupo Araujo Engenharia"
    
    ' Adiciona apenas os nomes
    usuariosPermitidosNomes.Add "Vanessa"
    usuariosPermitidosNomes.Add "Ivete Vespasiano"
    usuariosPermitidosNomes.Add "Victor Timotti"
    
    ' Define a lista de usuários permitidos para cada aba individualmente
    Select Case nomeDaAba
        Case "Gustavo"
            usuariosPermitidosCompletos.Add "Gustavo Migray | Grupo Araújo Engenharia"
            usuariosPermitidosNomes.Add "Gustavo Migray"
        Case "Andre"
            usuariosPermitidosCompletos.Add "Andre Padua | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Andre Padua"
        Case "Marco"
            usuariosPermitidosCompletos.Add "Marco Oliveira | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Marco Oliveira"
        Case "João"
            usuariosPermitidosCompletos.Add "Joao Paulo | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Joao Paulo"
        Case "Fernanda"
            usuariosPermitidosCompletos.Add "Fernanda Bueno | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Fernanda Bueno"
        Case "Renato"
            usuariosPermitidosCompletos.Add "Renato Carvalho | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Renato Carvalho"
        Case "Marcos"
            usuariosPermitidosCompletos.Add "Renato Carvalho | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Renato Carvalho"
        Case "Cleo"
            usuariosPermitidosCompletos.Add "Qualidade | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Qualidade"
        Case "Vanessa"
            usuariosPermitidosCompletos.Add "Vanessa | Grupo Araujo Engenharia"
            usuariosPermitidosNomes.Add "Vanessa"
    End Select

    ' Inicializa a variável de controle
    usuarioPermitido = False

    ' Verifica se o usuário está na lista de permitidos com o nome completo
    For Each usuario In usuariosPermitidosCompletos
        If usuario = username Then
            Debug.Print "Usuário permitido: " & usuario
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
            MsgBox "Você não tem permissão para editar esta planilha!"
        End If
    End With
End Sub
