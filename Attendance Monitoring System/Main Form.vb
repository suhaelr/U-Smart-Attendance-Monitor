Imports System.Data.OleDb
Public Class Main_Form
    Dim con As OleDb.OleDbConnection = myconnection()
    Dim dt As New DataTable
    Dim sql As String
    Dim da As New OleDb.OleDbDataAdapter
    Dim cmd As New OleDb.OleDbCommand
    Dim result As Integer
    Dim str As String



#Region "Clear Registration"
    Private Sub ClearRegistration()
        For Each crt As Control In GroupBox2.Controls
            If crt.GetType Is GetType(TextBox) Then
                crt.Text = Nothing
            End If
            ComboBox5.Text = ""
            ComboBox3.Text = ""
            ComboBox4.Text = ""
            DateTimePicker1.Text = ""
        Next
    End Sub
#End Region

#Region "Clear Event"
    Private Sub ClearEvent()
        For Each crt As Control In GroupBox3.Controls
            If crt.GetType Is GetType(TextBox) Then
                crt.Text = Nothing
            End If
            DateTimePicker1.Text = ""
        Next
    End Sub
#End Region

#Region "Clear User"
    Private Sub ClearUser()
        For Each crt As Control In GroupBox4.Controls
            If crt.GetType Is GetType(TextBox) Then
                crt.Text = Nothing
            End If
        Next
    End Sub
#End Region

#Region "Timer"
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label1.Text = My.Computer.Clock.LocalTime.Date
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Label2.Text = TimeOfDay
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
        Timer2.Start()
        LoginForm1.OK.SendToBack()
       
    End Sub
#End Region

#Region "Gender"
    Dim gender As String

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        gender = "Male"
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        gender = "Female"
    End Sub
#End Region

#Region "Save Student Record"
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Please fill up the given data", MsgBoxStyle.Question, MessageBoxDefaultButton.Button1)
        Else
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            con.Open()
            sql = "insert into TableStudent (YearLevel,StudentID,FirstName,MiddleName,LastName,Gender,Birthdate,Age,Course,Status)values('" & ComboBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & gender & "','" & DateTimePicker1.Text & "','" & TextBox4.Text & "','" & ComboBox5.Text & "','" & ComboBox3.Text & "')"
            With cmd
                .CommandText = sql
                .Connection = con
                result = cmd.ExecuteNonQuery
            End With

            If result > 0 Then
                MsgBox("New Student has added")
                ClearRegistration()
            Else
                MsgBox("No Student record has been added!", MsgBoxStyle.Critical)
            End If
        End If
        con.Close()
    End Sub
#End Region

#Region "SaveEvent"
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim ds As DataSet
        con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
        con.Open()
        da = New OleDb.OleDbDataAdapter("INSERT INTO TableEvent (EventName,EventDate,EventVenue) VALUES ('" & TextBox8.Text & "','" & DateTimePicker2.Text & "','" & TextBox7.Text & "')", con)
        ds = New DataSet
        da.Fill(ds)
        MsgBox("New Event has been Saved", MsgBoxStyle.OkOnly)
        con.Close()
        storing()
    End Sub
#End Region

#Region "Age Conversion"
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim day As Integer = DateDiff(DateInterval.Day, DateTimePicker1.Value, Now) Mod 365
        Dim yr As Integer = DateDiff(DateInterval.Year, DateTimePicker1.Value, Now)
        Dim month As Integer = DateDiff(DateInterval.Month, DateTimePicker1.Value, Now) Mod 12
        TextBox4.Text = yr & " Years, " & month & " Months "
    End Sub
#End Region

#Region " Adding new Encoder"
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If TextBox11.Text = "" Or TextBox12.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Then
            MsgBox("No encoder has been saved!", MsgBoxStyle.Critical)
        Else
            Dim connect As New OleDb.OleDbConnection
            connect.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            connect.Open()
            str = "INSERT INTO TableUser ([username],[password], [FirstName], [LastName])values (?, ?, ?, ?)"
            Dim cmd As OleDb.OleDbCommand = New OleDbCommand(str, connect)

            cmd.Parameters.Add(New OleDbParameter("username", CType(TextBox11.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("password", CType(TextBox12.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("FirstName", CType(TextBox9.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LastName", CType(TextBox10.Text, String)))

            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            connect.Close()
            ClearUser()
            MsgBox("New User has been added", MsgBoxStyle.OkOnly)
        End If
    End Sub
#End Region

#Region "TIME SETTINGS"
    Dim exactamIn As String = "7:30:00 AM"
    Dim exactamOut As String = "11:30:00 AM"
    Dim exactpmIn As String = "1:30:00 PM"
    Dim exactpmOut As String = "5:30:00 PM"
    Dim remark As String
#End Region

#Region "Time Notification"
    Public Sub ONTIMEAM()
        If Label2.Text <= exactamIn Then
            remark = "ON TIME"
            MsgBox("VERY GOOD YOUR EARLY")
        ElseIf Label2.Text >= exactamIn Then
            remark = "LATE"
            MsgBox("YOUR LATE")
        End If
    End Sub

    Public Sub ONTIMEAMOUT()
        If Label2.Text <= exactamOut Then
            remark = "EARLY"
            MsgBox("TOO EARLY")
        ElseIf Label2.Text >= exactamOut Then
            remark = "LATE"
            MsgBox("TOO LATE")
        End If
    End Sub


    Public Sub ONTIMEPM()
        If Label2.Text <= exactpmIn Then
            remark = "ON TIME"
            MsgBox("VERY GOOD YOUR EARLY")
        ElseIf Label2.Text >= exactpmIn Then
            remark = "LATE"
            MsgBox("YOUR LATE")
        End If
    End Sub

    Public Sub ONTIMEPMOUT()
        If Label2.Text <= exactpmOut Then
            remark = "NOT ON TIME "
        ElseIf Label2.Text >= exactpmOut Then
            remark = "ON TIME"
        End If
    End Sub
#End Region

#Region "Sign InAM"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim provider As String
        Dim datafile As String = ""
        Dim connstring As String
        Dim myconnection As OleDbConnection = New OleDbConnection

        If TextBox1.Text = "" Then
            MsgBox("Please input your I.D", MsgBoxStyle.Critical)
        Else

            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            connstring = provider & datafile
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim cmd As OleDbCommand = New OleDbCommand(" SELECT * FROM [TableStudent] WHERE [StudentID] = '" & TextBox1.Text & "'", myconnection)
            Dim dr As OleDbDataReader = cmd.ExecuteReader
            Dim userfound As Boolean = False
            Dim stud As String
            Dim fname As String
            Dim lname As String
            While dr.Read
                userfound = True
                stud = dr("StudentID").ToString
                fname = dr("FirstName").ToString
                lname = dr("LastName").ToString

                Name = fname & " " & lname
            End While
            If userfound = True Then

                ONTIMEAM()

                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
                con.Open()
                sql = "Insert into TableInsert (StudID,Am_In , Am_InRemark , EventName,Encoder) values ('" & TextBox1.Text & " ','" & Label2.Text & "','" & remark & "','" & ComboBox1.Text & "','" & Label8.Text & "')"

                With cmd
                    .CommandText = sql
                    .Connection = con
                    result = cmd.ExecuteNonQuery
                End With
                MsgBox("Successfully Sign In" & " " & Name)
                con.Close()

            Else
                MsgBox("Student ID Not found or Incorrect", MsgBoxStyle.Critical)
                TextBox1.Clear()
            End If
        End If
    End Sub
#End Region

#Region "Sign OutAM"

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim provider As String
        Dim datafile As String = ""
        Dim connstring As String
        Dim myconnection As OleDbConnection = New OleDbConnection

        If TextBox1.Text = "" Then
            MsgBox("Please input your I.D", MsgBoxStyle.Critical)
        Else


            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            connstring = provider & datafile
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim cmd As OleDbCommand = New OleDbCommand(" SELECT * FROM [TableStudent] WHERE [StudentID] = '" & TextBox1.Text & "'", myconnection)
            Dim dr As OleDbDataReader = cmd.ExecuteReader
            Dim userfound As Boolean = False
            Dim stud As String
            Dim fname As String
            Dim lname As String
            While dr.Read
                userfound = True
                stud = dr("StudentID").ToString
                fname = dr("FirstName").ToString
                lname = dr("LastName").ToString

                Name = fname & " " & lname
            End While
            If userfound = True Then

                ONTIMEAMOUT()

                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
                con.Open()
                sql = "Update TableInsert set  StudID ='" & TextBox1.Text & "', Am_Out ='" & Label2.Text & "', Am_OutRemark ='" & remark & "'Where EventName = '" & ComboBox1.Text & " '"
                With cmd
                    .CommandText = sql
                    .Connection = con
                End With
                result = cmd.ExecuteNonQuery
                MsgBox("Successfully Sign Out" & " " & Name)
                con.Close()

            ElseIf userfound = False Then
                MsgBox("Student ID Not found or Incorrect", MsgBoxStyle.Critical)
                TextBox1.Clear()
            End If
        End If
    End Sub
#End Region

#Region "Sign InPm"
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim provider As String
        Dim datafile As String = ""
        Dim connstring As String
        Dim myconnection As OleDbConnection = New OleDbConnection
        Dim ssql As String
        Dim olcmd As New OleDb.OleDbCommand

        If TextBox1.Text = "" Then
            MsgBox("Please input your I.D", MsgBoxStyle.Critical)
        End If
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
        connstring = provider & datafile
        myconnection.ConnectionString = connstring
        myconnection.Open()

        Dim cmd As OleDbCommand = New OleDbCommand(" SELECT * FROM [TableStudent] WHERE [StudentID] = '" & TextBox1.Text & "'", myconnection)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        Dim userfound As Boolean = False
        Dim stud As String
        Dim fname As String
        Dim lname As String
        Dim name As String = ""

        While dr.Read
            userfound = True
            stud = dr("StudentID").ToString
            fname = dr("FirstName").ToString
            lname = dr("LastName").ToString
            name = (fname & " " & lname)
        End While

        If userfound = False Then

            MsgBox("Student ID not found or Incorrect", MsgBoxStyle.Critical)
        End If

        If userfound = True Then
            ONTIMEPM()
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            con.Open()
            sql = "Update TableInsert set  StudID ='" & TextBox1.Text & "', Pm_In ='" & Label2.Text & "', Pm_InRemark ='" & remark & "'Where EventName = '" & ComboBox1.Text & " '"

            With cmd
                .CommandText = sql
                .Connection = con
                result = cmd.ExecuteNonQuery
            End With

            MsgBox("Successfully Sign In" & " " & name)
            con.Close()
        Else
            ONTIMEPM()
            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            con.Open()
            ssql = "Insert into TableInsert (StudID ,Pm_In , Pm_InRemark , EventName,Encoder) values ('" & TextBox1.Text & " ','" & Label2.Text & "','" & remark & "','" & ComboBox1.Text & "','" & Label8.Text & "')"

            With olcmd
                .CommandText = ssql
                .Connection = con
                result = olcmd.ExecuteNonQuery
            End With

            MsgBox("Successfully Sign In" & " " & name)
            con.Close()

        End If

    End Sub
#End Region

#Region "Sign OutPM"
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim provider As String
        Dim datafile As String = ""
        Dim connstring As String
        Dim myconnection As OleDbConnection = New OleDbConnection

        If TextBox1.Text = "" Then
            MsgBox("Please input your I.D", MsgBoxStyle.Critical)
        Else


            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            connstring = provider & datafile
            myconnection.ConnectionString = connstring
            myconnection.Open()

            Dim cmd As OleDbCommand = New OleDbCommand(" SELECT * FROM [TableStudent] WHERE [StudentID] = '" & TextBox1.Text & "'", myconnection)
            Dim dr As OleDbDataReader = cmd.ExecuteReader
            Dim userfound As Boolean = False
            Dim stud As String
            Dim fname As String
            Dim lname As String
            While dr.Read
                userfound = True
                stud = dr("StudentID").ToString
                fname = dr("FirstName").ToString
                lname = dr("LastName").ToString

                Name = fname & " " & lname
            End While
            If userfound = True Then

                ONTIMEAMOUT()

                con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
                con.Open()
                sql = "Update TableInsert set  StudID ='" & TextBox1.Text & "', Pm_Out ='" & Label2.Text & "', Pm_OutRemark ='" & remark & "'Where EventName = '" & ComboBox1.Text & " '"
                With cmd
                    .CommandText = sql
                    .Connection = con
                End With
                result = cmd.ExecuteNonQuery

                MsgBox("Successfully Sign Out" & " " & Name)
                con.Close()

            Else
                MsgBox("Student ID Not found or Incorrect", MsgBoxStyle.Critical)
                TextBox1.Clear()

            End If
        End If
    End Sub
#End Region

#Region "Load Student"
    Public Sub load_Student()
        myconnection.Open()
        dt = New DataTable

        With cmd
            .Connection = myconnection()
            .CommandText = "Select * from TableStudent '"
        End With




        da.SelectCommand = cmd
        da.Fill(dt)



        DataGridView2.DataSource = dt


        da.Dispose()

        myconnection.Close()

    End Sub
#End Region

#Region "Add Administrator"
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If TextBox22.Text = "" Or TextBox23.Text = "" Or TextBox21.Text = "" Then
            MsgBox("No Administrator has been saved!", MsgBoxStyle.Critical)

        Else
            Dim connect As New OleDb.OleDbConnection
            connect.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            connect.Open()
            str = "INSERT INTO TableAdmin ([username], [FirstName], [LastName])values (?, ?, ?)"
            Dim cmd As OleDb.OleDbCommand = New OleDbCommand(str, connect)

            cmd.Parameters.Add(New OleDbParameter("username", CType(TextBox22.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("FirstName", CType(TextBox23.Text, String)))
            cmd.Parameters.Add(New OleDbParameter("LastName", CType(TextBox21.Text, String)))

            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            connect.Close()
            MsgBox("New Administrator has been added", MsgBoxStyle.OkOnly)
            TextBox22.Clear()
            TextBox23.Clear()
            TextBox21.Clear()
            LoadAdminRecords()
        End If

    End Sub
#End Region

#Region "Admin Sign In"
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click

        Dim provider As String
        Dim datafile As String = ""
        Dim constring As String
        Dim myconnection As OleDbConnection = New OleDbConnection
        Dim da As New OleDb.OleDbDataAdapter

        If TextBox24.Text = "" Then
            ErrorProvider1.SetError(TextBox24, "Please input username")
        Else
            provider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\SSAOAMS.accdb"
            constring = provider & datafile
            myconnection.ConnectionString = constring
            myconnection.Open()

            Dim cmd As OleDbCommand = New OleDbCommand(" SELECT * FROM [TableAdmin] WHERE [Username] = '" & TextBox24.Text & " " & "'", myconnection)
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
                MsgBox("Welcome as Administrator" & " " & FirstName & " " & LastName)
                GroupBox7.Enabled = True
                ErrorProvider1.Clear()
                TextBox24.Clear()
                LoadAdminRecords()
                enabletrue()
                Button13.Hide()
                Button9.Show()
                DataGridView4.Visible = True

            Else
                ErrorProvider1.SetError(TextBox24, "Wrong Username")
                TextBox24.Clear()
            End If
            myconnection.Close()
        End If
    End Sub
#End Region

#Region "Delete Admin"
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Try
            sql = "DELETE * FROM TableAdmin  WHERE AdminID=" & Me.Text
            con.Open()
            With cmd
                .CommandText = sql
                .Connection = con
            End With

            result = cmd.ExecuteNonQuery
            If result > 0 Then
                MsgBox("NEW RECORD HAS BEEN DELETED!")
                con.Close()
                Loadadmin()
                TextBox23.Clear()
                TextBox22.Clear()
                TextBox21.Clear()
            Else
                MsgBox("NO RECORD HAS BEEN DELETED!", MsgBoxStyle.Critical)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()


        End Try
    End Sub
#End Region

#Region "Update Admin"
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            sql = "UPDATE TableAdmin set  Username= '" & TextBox22.Text & "', Firstname = '" & TextBox23.Text & "', Lastname = '" & TextBox21.Text & "'WHERE AdminID = " & Me.Text

            con.Open()
            With cmd
                .CommandText = sql
                .Connection = con
            End With
            result = cmd.ExecuteNonQuery
            If result > 0 Then
                MsgBox("Admin record has been Updated")

            Else
                MsgBox("Admin record was not been Updated", MsgBoxStyle.Critical)
            End If
            Loadadmin()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Private Sub DataGridView4_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        Try
            Me.Text = DataGridView4.CurrentRow.Cells(0).Value.ToString
            TextBox22.Text = DataGridView4.CurrentRow.Cells(1).Value.ToString
            TextBox23.Text = DataGridView4.Rows(e.RowIndex).Cells(2).Value.ToString
            TextBox21.Text = DataGridView4.Rows(e.RowIndex).Cells(3).Value.ToString
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#End Region

    Private Sub TabControl3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl3.Click
        LoadRecords()
        load_Student()

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ClearEvent()
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try
            TextBox25.Text = DataGridView1.Rows(e.RowIndex).Cells(1).Value.ToString



            Dim cmd1 As New OleDbCommand
            Dim dt1 As New DataTable
            Dim da1 As New OleDb.OleDbDataAdapter

            myconnection.Open()
            dt = New DataTable

            With cmd1
                .Connection = myconnection()
                .CommandText = "Select * from TableInsert where StudID like '" & TextBox25.Text & "%'"
            End With

            da1.SelectCommand = cmd1
            da1.Fill(dt1)
            DataGridView3.DataSource = dt1

            da.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            myconnection.Close()
        End Try
    End Sub





    Dim acscmd As New OleDb.OleDbCommand
    Dim acsda As New OleDb.OleDbDataAdapter
    Dim acscon As OleDb.OleDbConnection = myconnection()
    Dim acsds As New DataSet
    Dim strsql As String
    Dim strreportname As String

    Public Sub getmessage()
        If message = "all reports" Then
            report("SELECT * FROM TableInsert", "Report1")

        ElseIf message = "Stud report" Then
            report("SELECT * FROM TableInsert where StudID = '" & Me.TextBox25.Text & "'", "Report1")

        ElseIf message = "Course Reports" Then
            report("SELECT * FROM TableStudent where Course = '" & Me.ComboBox2.Text & "'", "rptcourse")

        ElseIf message = "Course Year Reports" Then
            report("SELECT * FROM TableStudent where YearLevel = '" & Me.ComboBox6.Text & "'and Course = '" & Me.ComboBox2.Text & "'", "rptyearlevel")

        ElseIf message = "Year Reports" Then
            report("SELECT * FROM TableStudent where YearLevel = '" & Me.ComboBox6.Text & "'", "rptyearlevel")
        End If
    End Sub
#Region "reports"
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        message = "Stud report"
        getmessage()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        message = "all reports"
        getmessage()
    End Sub
#End Region

    '-----------------------------------------------------------
#Region "reports"

    Public Sub report(ByVal sql As String, ByVal rptname As String)
        acsds = New DataSet
        strsql = sql
        acscmd.CommandText = strsql
        acscmd.Connection = acscon
        acsda.SelectCommand = acscmd
        acsda.Fill(acsds)

        strreportname = rptname
        Dim strreportpath As String = Application.StartupPath & "\report\" & strreportname & ".rpt"

        If Not IO.File.Exists(strreportpath) Then
            MsgBox("Unable to locate file :" & vbCrLf & strreportpath)
        End If

        Dim reportdoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        reportdoc.Load(strreportpath)
        reportdoc.SetDataSource(acsds.Tables(0))
        CrystalReportViewer1.ShowRefreshButton = False
        CrystalReportViewer1.ShowCloseButton = False
        CrystalReportViewer1.ShowGroupTreeButton = False
        CrystalReportViewer1.ReportSource = reportdoc

    End Sub
#End Region

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ComboBox2.Text = "" Then
            MsgBox("Please select Course first")
        Else
            message = "Course Reports"
            getmessage()
        End If
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ComboBox2.Text = "" Or ComboBox6.Text = "" Then
            MsgBox("Please select Course or Year Level first")
        Else
            message = "Course Year Reports"
            getmessage()
        End If
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If ComboBox6.Text = "" Then
            MsgBox("Please select Year Level first")
        Else
            message = "Year Reports"
            getmessage()
        End If
    End Sub


    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Loadadmin()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Button13.Show()
        Button9.Hide()
        enablefalse()
        GroupBox7.Enabled = False
        DataGridView4.Visible = False
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox21.Clear()

    End Sub

    Private Sub TabPage8_Click(sender As Object, e As EventArgs) Handles TabPage8.Click

    End Sub

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

    End Sub
End Class