Imports System.Data.OleDb
Public Class LoginForm1
    Dim provider As String
    Dim datafile As String
    Dim constring As String
    Dim myconnection As OleDbConnection = New OleDbConnection
    Dim ds As DataSet
    Dim da As New OleDb.OleDbDataAdapter

#Region "ClearLogin"
    Public Sub cleartextfields()
        UsernameTextBox.Clear()
        PasswordTextBox.Clear()
    End Sub
#End Region

#Region "Login"
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If UsernameTextBox.Text = "" Then
            ErrorProvider1.SetError(UsernameTextBox, "Please input Username")
        Else
            ErrorProvider1.SetError(UsernameTextBox, "Incorrect Username")
        End If
        If PasswordTextBox.Text = "" Then
            ErrorProvider1.SetError(PasswordTextBox, "Please input Password")
        Else
            ErrorProvider1.SetError(PasswordTextBox, "Incorrect Password")
        End If

        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
        constring = provider & datafile
        myconnection.ConnectionString = constring
        myconnection.Open()

        Dim cmd As OleDbCommand = New OleDbCommand(" SELECT * FROM [TableUser] WHERE [username] = '" & UsernameTextBox.Text & "' AND [password] = '" & PasswordTextBox.Text & "'", myconnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader

        Dim userfound As Boolean = False
        Dim FirstName As String = ""
        Dim LastName As String = ""

        While dr.Read
            userfound = True
            FirstName = dr("FirstName").ToString
            LastName = dr("LastName").ToString
        End While

        If userfound = True Then
            storing()
            Main_Form.Show()
            cleartextfields()
            ErrorProvider1.Clear()
            Main_Form.Label8.Text = FirstName & " " & LastName
            enablefalse()

        End If
        myconnection.Close()
    End Sub
#End Region

   
   
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Exit_Program()
    End Sub
End Class
