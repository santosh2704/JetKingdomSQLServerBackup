Imports System.IO
Imports System.Security.Permissions.FileIOPermissionAccess
Imports System.Security.SecurityException
Public Class Backup
    Private Part1, Part2, Part3, CMonth As String
    Private BackupFileName As String

    Private Sub Backup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        txtFN.Focus()
        Button1.Enabled = False

        Call ReturnMonthValue(Today.Month)
        Part1 = CMonth


        Part2 = Today.Day.ToString
        Part3 = Today.Year.ToString
        BackupFileName = "D:\SQLBackup\JetKingdom_" & Part2 & "_" & Part1 & "_" & Part3 & ".Bak"
        txtFN.Text = BackupFileName

        If Len(txtFN.Text) > 1 Then
            Button1.Enabled = True

            Try
                ApplnPath = Application.StartupPath()
                ApplnDrive = Application.StartupPath.Chars(0)

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Try
                Dim fs As New FileStream("C:\JetServerSetup.txt", FileMode.Open)
                Dim sr As New StreamReader(fs)

                ' First Line to read Data SOurce Name

                ServerName = sr.ReadLine()


                ' Second Line to read Password
                SQLServerPassword = sr.ReadLine()

                Call ObjClass1.ReturnFromServerSetupFile1(SQLServerPassword)
                SQLServerPassword = FinalString

                ''Third line  to read number of active nodes/connections
                ud_NumberNodes = sr.ReadLine()

                'Call ObjClass1.ReturnNodesFromServerSetupFile1(ud_NumberNodes)
                ud_NumberNodes = FinalNodes

                ''fourth line to read name of the database
                ud_DBName = sr.ReadLine()

                ''fifthline to read sms authorization key
                'AutorizationKey = sr.ReadLine

                ''Sixthline to read sms SenderID
                'SenderID = sr.ReadLine

                ''seventh line for SMS Pack URL
                'SMSURL = sr.ReadLine

                fs.Close()
                sr.Close()

            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try

            Try
                Dim fs1 As New FileStream("D:\JetServerSetup.txt", FileMode.Open)
                Dim sr1 As New StreamReader(fs1)

                ' First Line to read Data SOurce Name

                ServerName = sr1.ReadLine()


                ' Second Line to read Password
                SQLServerPassword = sr1.ReadLine()

                Call ObjClass1.ReturnFromServerSetupFile1(SQLServerPassword)
                SQLServerPassword = FinalString

                ''Third line  to read number of active nodes/connections
                ud_NumberNodes = sr1.ReadLine()

                'Call ObjClass1.ReturnNodesFromServerSetupFile1(ud_NumberNodes)
                ud_NumberNodes = FinalNodes

                ''fourth line to read name of the database
                ud_DBName = sr1.ReadLine()

                ''fifthline to read sms authorization key
                'AutorizationKey = sr1.ReadLine

                ''Sixthline to read sms SenderID
                'SenderID = sr1.ReadLine

                ''seventh line for SMS Pack URL
                'SMSURL = sr1.ReadLine

                fs1.Close()
                sr1.Close()
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        Else
            MsgBox("Please Enter FileName, FileName Cannot be Empty")
        End If

        'connection to the server
        ObjClass1.TestConnection(ServerName)
        ObjClass2.TestConnection(ServerName)
        ObjClass3.TestConnection(ServerName)

        Label1.Text = "Connected to SQL Server " & UCase(ServerName)
        Label2.Text = "Software Powered by Santosh Shrigadiwar "
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        BackupFileName = txtFN.Text

        Dim Found As Boolean
        'check if backup file exist
        Try
            Dim fs2 As New FileStream(BackupFileName, FileMode.Open)
            Dim sr2 As New StreamReader(fs2)

            Found = True
            fs2.Close()
            sr2.Close()
        Catch ex As Exception
            Found = False
        End Try
        Try
            If Found = True Then
                My.Computer.FileSystem.DeleteFile(BackupFileName, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
                MsgBox("Databse Backup is In Progress, Please do not Close Program or Restart System")
                Call Backup()
            Else
                MsgBox("Database Backup is In Progress, Please do not Close Program or Restart System")
                Call Backup()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub
    Private Sub Backup()
        QUERY = "Backup Database JetKingdom to DISK ='" & BackupFileName & "'"

        cmd = New SqlClient.SqlCommand(QUERY, ObjClass1.Myconn)
        If ObjClass1.Myconn.State = ConnectionState.Closed Then
            ObjClass1.Myconn.Open()
        End If
        Try
            cmd.ExecuteNonQuery()
            ObjClass1.Myconn.Close()
            MsgBox("Database Backup done Successfully!")
            txtFN.Text = ""
            Button1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ReturnMonthValue(ByVal IntMon As Integer)
        If IntMon = 1 Then
            CMonth = "JAN"
        End If
        If IntMon = 2 Then
            CMonth = "FEB"
        End If
        If IntMon = 3 Then
            CMonth = "MAR"
        End If
        If IntMon = 4 Then
            CMonth = "APR"
        End If
        If IntMon = 5 Then
            CMonth = "MAY"
        End If
        If IntMon = 6 Then
            CMonth = "JUN"
        End If
        If IntMon = 7 Then
            CMonth = "JUL"
        End If
        If IntMon = 8 Then
            CMonth = "AUG"
        End If
        If IntMon = 9 Then
            CMonth = "SEP"
        End If
        If IntMon = 10 Then
            CMonth = "OCT"
        End If
        If IntMon = 11 Then
            CMonth = "NOV"
        End If
        If IntMon = 12 Then
            CMonth = "DEC"
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim obj1 As New XMLBackupRestore
        obj1.Show()
    End Sub
End Class
