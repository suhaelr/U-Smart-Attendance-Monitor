Module StoreEvent
    Dim con As OleDb.OleDbConnection = New  _
OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb")

    Dim da As OleDb.OleDbDataAdapter
    Dim ds As DataSet
    Dim dt As DataTable
    Dim sql As String

#Region "StoreEvent"
    Public Sub storing()

        con.Open()
        da = New OleDb.OleDbDataAdapter("SELECT EventDate,(EventName) from TableEvent ", con)
        ds = New DataSet
        da.Fill(ds, "SSAOAMS")
        With Main_Form.ComboBox1
            .DataSource = ds.Tables(0)
            .DisplayMember = "EventName"
            .ValueMember = "EventDate"
        End With
        con.Close()
    End Sub
#End Region

#Region "Load Student Record"
    Public Sub LoadRecords()
        dt = New DataTable
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb")
        con.Open()
        sql = "SELECT * FROM TableStudent;"
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(dt)
        Main_Form.DataGridView1.DataSource = dt
        con.Close()
    End Sub
#End Region

#Region "Load Administrator Records"
    Public Sub LoadAdminRecords()
        dt = New DataTable
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb")
        con.Open()
        sql = "SELECT * FROM TableAdmin;"
        da = New OleDb.OleDbDataAdapter(sql, con)
        da.Fill(dt)
        Main_Form.DataGridView4.DataSource = dt
        Main_Form.DataGridView4.Columns(0).Visible = False
        Main_Form.DataGridView4.Columns(1).Visible = False
        con.Close()
    End Sub
#End Region

End Module
