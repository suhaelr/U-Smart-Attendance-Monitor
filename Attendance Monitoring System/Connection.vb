Module Connection

    Public message As String = ""


    Public Function myconnection() As OleDb.OleDbConnection
        Return New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb")
    End Function

    Dim dt As New DataTable
    Dim sql As String
    Dim da As New OleDb.OleDbDataAdapter

#Region "LOGOUT"
    Public Sub Exit_Program()
        Dim notif As String
        Dim result As DialogResult

        notif = "Are you sure you want to exit program?"
        result = MessageBox.Show(notif, " Exit Program ", MessageBoxButtons.YesNo, _
                        MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

        If result = Windows.Forms.DialogResult.Yes Then
            LoginForm1.Close()
            Main_Form.Close()
        ElseIf result = Windows.Forms.DialogResult.No Then
            LoginForm1.Button1.Show()
            LoginForm1.OK.SendToBack()
            Main_Form.Show()
            enablefalse()
            storing()
        End If
    End Sub
#End Region

#Region "LoadAdmin"
    Public Sub Loadadmin()
        dt = New DataTable
        myconnection.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb")
        myconnection.Open()
        sql = "SELECT * FROM TableAdmin;"
        da = New OleDb.OleDbDataAdapter(sql, myconnection)
        da.Fill(dt)
        Main_Form.DataGridView4.DataSource = dt
        myconnection.Close()
    End Sub
#End Region

#Region "enable false"
    Public Sub enablefalse()
        With Main_Form
            .GroupBox2.Enabled = False
            .Panel3.Enabled = False
            .GroupBox3.Enabled = False
            .GroupBox4.Enabled = False
        End With
    End Sub

#End Region

#Region "enable true"
    Public Sub enabletrue()
        With Main_Form
            .GroupBox2.Enabled = True
            .Panel3.Enabled = True
            .GroupBox3.Enabled = True
            .GroupBox4.Enabled = True
        End With
    End Sub

#End Region

End Module
