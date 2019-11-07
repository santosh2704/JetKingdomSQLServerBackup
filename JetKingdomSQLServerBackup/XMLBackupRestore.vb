Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
Imports System.Security.Permissions.FileIOPermissionAccess

Public Class XMLBackupRestore
    Private i, TotalTables, RowNo As Short
    Private XMLFileName As String
    Private Increment As Integer = 0

   
    Private Sub XMLBackupRestore_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Call Proc1()
        Call DataGridFormatting1()

    End Sub

  
    Private Sub DataGridFormatting1()
        Try
            DataGridView2.Columns(0).Width = 180
            DataGridView2.Columns(1).Width = 100
            DataGridView2.Columns(2).Width = 110
            DataGridView2.Columns(3).Width = 110
            DataGridView2.Columns(4).Width = 90

            DataGridView2.Columns(0).DefaultCellStyle.Font = New Font("Calibri", 11, FontStyle.Bold)

        Catch ex As Exception

        End Try


    End Sub
    Private Sub Proc1()

        Try

            Query1 = "select Name,object_id,create_date,modify_date,max_column_id_used  from " & ud_DBName & ".sys.tables order by name"

            da = New SqlClient.SqlDataAdapter(Query1, ObjClass1.Myconn)
            ds = New DataSet
            da.Fill(ds, "Tan3")
            DataGridView2.DataSource = ds.Tables("Tan3")
            '----------------end
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdBackup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBackup.Click
        Call Proc1()
        Call DataGridFormatting1()

        TotalTables = DataGridView2.Rows.Count
        Label2.Text = "Total Tables " & TotalTables
        'ProgressBar1.PerformStep()
        'Timer1.Start()
        Try
            For i = 0 To TotalTables - 1
                'Increment = Increment + 10
                TableName = DataGridView2.Rows(i).Cells.Item("Name").Value
                If TableName = "" Then
                    Exit For
                End If
                'Creates XML Document
                QUERY = "Select * from " & TableName
                da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
                ds = New DataSet
                da.Fill(ds, "LLL")
                XMLFileName = "D:\SQLBackupXML\" & TableName & ".XML"
                ds.WriteXml(XMLFileName)
                'If Increment > ProgressBar1.Maximum Then
                '    Increment = ProgressBar1.Maximum
                'End If
                'ProgressBar1.Value = Increment
                'Timer1.Interval = 1
            Next
            MsgBox("Backup Process Completed.....")
        Catch ex As Exception
        Finally

            Me.Close()
        End Try
       

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

   


   
End Class