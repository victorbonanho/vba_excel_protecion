Private Sub Workbook_Open()
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
    usuariosPermitidos.Add "UsuarioAdmin1"
    usuariosPermitidos.Add "UsuarioAdmin2"
    
    ' Define a lista de usuários permitidos para cada aba
    ' Você pode configurar as permissões para cada aba individualmente
    Select Case nomeDaAba
        Case "RH"
            usuariosPermitidos.Add "UsuarioRH1"
            usuariosPermitidos.Add "UsuarioRH2"
        Case "Financeiro"
            usuariosPermitidos.Add "UsuarioFinanceiro1"
            usuariosPermitidos.Add "UsuarioFinanceiro2"
        Case "Vendas"
            usuariosPermitidos.Add "UsuarioVendas1"
            usuariosPermitidos.Add "UsuarioVendas2"
        ' Adicione mais casos conforme necessário
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

    ' Se o usuário não estiver na lista, protege a planilha
    If Not usuarioPermitido Then
    
        ' MsgBox "Você não tem permissão para editar esta planilha!"

        ' Desprotege a planilha antes de alterar o bloqueio das células
        With ThisWorkbook.Sheets(nomeDaAba)
            .Unprotect ' Desprotege a planilha

            ' Desbloqueia todas as células primeiro (caso contrário, elas ficam bloqueadas automaticamente)
            .Cells.Locked = False
            
            ' Bloqueia todas as células novamente, para garantir que não sejam editadas
            .Cells.Locked = True

            ' Protege a planilha, impedindo edição de células bloqueadas
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
        End With
    End If
End Sub

