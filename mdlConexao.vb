Module mdlConexao
    Private aConexao As New ADODB.Connection
    Const BancoDeDados = "dbVendas"

    Const acess_db = "bancodedados"
    Const acess_usr = "usuario"
    Const acess_pwd = "senha"
    Const acess_server = "(local)"

    Public Enum tpServidor
        Access = 0
        SqlServer = 1
    End Enum

    Public Function RecebeTabela(ByVal Sql As String, Optional ByVal servidor As tpServidor = tpServidor.Access)

        Dim aResultado As New ADODB.Recordset
        Dim acess_db2 As String
        acess_db2 = My.Computer.FileSystem.CurrentDirectory
        acess_db2 &= IIf(Right(acess_db2, 1) = "\", "", "\")
        acess_db2 &= BancoDeDados & ".mdb"
        'MsgBox(acess_db2)
        If aConexao.State = 1 Then aConexao.Close()

        If servidor = tpServidor.SqlServer Then

            aConexao.ConnectionString = "driver={sql server};" & _
                                            "server=" + acess_server + ";" & _
                                            "Database=" + acess_db + ";" & _
                                            "PWD=" + acess_pwd + ";" & _
                                            "UID=" + acess_usr + ";"
        Else
            aConexao.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & acess_db2 & ";Persist Security Info=False"
        End If
        Try
            aConexao.Open()

        Catch ex As Exception
            aResultado = Nothing
            MsgBox("Banco de Dados não encontrado!")
            GoTo fim
        End Try

        aResultado.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        aResultado.Open(Sql, aConexao, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
fim:
        RecebeTabela = aResultado
        aResultado = Nothing
    End Function

    Public Sub CarregarList(ByVal tabela As String, ByRef list As ListView, ByVal nome As String, ByVal codigo As String, Optional ByVal filtro As String = "", Optional ByVal Campos As String = "*")
        list.Visible = False

        Dim variavel As ADODB.Recordset, X As Integer, Y As Integer

        variavel = RecebeTabela("select " & Campos & " from " & tabela & " where " & IIf(IsNumeric(filtro), codigo & "=" & filtro, nome & " like'" & filtro & "%'"))

        list.Clear()
        list.FullRowSelect = True
        list.MultiSelect = True
        list.View = View.Details
        'If Not variavel.EOF Then
        For Y = 0 To variavel.Fields.Count - 1
            list.Columns.Add(variavel(Y).Name)
        Next
        list.Items.Clear()
        ' End If

        Do Until variavel.EOF
            list.Items.Add(variavel(0).Value.ToString)
            For Y = 1 To list.Columns.Count - 1
                list.Items(X).SubItems.Add(variavel(Y).Value.ToString)
            Next
            variavel.MoveNext()
            X += 1
        Loop
        variavel.Close()
        list.Visible = True


    End Sub


End Module
