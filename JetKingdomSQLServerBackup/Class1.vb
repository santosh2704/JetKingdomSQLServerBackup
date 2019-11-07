Imports System.Net
Imports System.IO
Imports System.Drawing
Imports System.ComponentModel
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.CrystalReports.Engine.Tables
Imports System.Net.NetworkInformation
Public Class ConnectionClass
    Public Myconn As New SqlClient.SqlConnection
    Public Myconn1 As New SqlClient.SqlConnection
    Public str1 As String
    Private rno, Count, Count1 As Short
    Private aa As Integer
    Private StrArr(500) As String
    Private StrArr1(500) As String
    Private ArrIndex As Short
    Private ArrIndex1 As Short
    Private number As Short
    Private x, y As Short
    Private TableName1 As String
    Private ds_Count1, ds_Count2, ds_Rno1, ds_Rno2 As Short
    Private ds_ChitGroup, ds_ClientCode, ds_ChitTableName, ds_Query2, ds_Query3, ds_ClientType, ds_Status, GroupNumber As String
    Private ds_ChitInstall, ds_ChitAmt, ds_Amt As Integer
    Private ds_InsNo, ds_InsNo1 As Short
    Private ds_DateR, ds_DateADD As Date
    Private ds_SrNo, ds_SNO1 As Integer
    Private Chit_Auction As Short
    Private ForemenCommission, Dividend As Integer
    Private GQuery, CC, Query3, CCode As String
    Private IR, CountInt, RowInt, RCount, CountWan1, RWan1, Array2, Array1 As Short
    Private SerialNumber As Integer
    Private ReturnDate As Date
    Private ud_Amount As Double
    Private ud_Status, CustID As String
    Private IntRate, InterestAmt As Double
    Private IntIncr As Double
    Private CB As String

    Public Sub TestConnection(ByVal SName As String)
        Call CloseConnections()
        str1 = "User id=sa;Password=" & SQLServerPassword & ";Database= " & ud_DBName & ";"
        str1 = str1 & "Data Source=" & SName
        Myconn.ConnectionString = str1
    End Sub
    Public Sub CloseConnections()
        If ObjClass1.Myconn.State = ConnectionState.Open Then
            ObjClass1.Myconn.Close()
        End If
        If ObjClass2.Myconn.State = ConnectionState.Open Then
            ObjClass2.Myconn.Close()
        End If
        If ObjClass3.Myconn.State = ConnectionState.Open Then
            ObjClass3.Myconn.Close()
        End If

    End Sub
    
    Public Sub ReturnReceiptNumber1(ByVal Tablename As String, ByVal Constr As SqlClient.SqlConnection)
        QUERY = "Select MAX(RNAlloted) FROM " & Tablename
        cmd = New SqlClient.SqlCommand(QUERY, Constr)
        If Constr.State = ConnectionState.Closed Then
            Constr.Open()
        End If

        Try
            ReceiptNumber = cmd.ExecuteScalar
            If ReceiptNumber = 0 Then
                ReceiptNumber = SRN - 1
            End If
        Catch ex As Exception
            ReceiptNumber = SRN
            MsgBox(ex.Message)
        Finally
            Constr.Close()
        End Try

    End Sub
    Public Sub ReturnCalculatedData(ByVal QueryString As String, ByVal ConnString As SqlClient.SqlConnection)
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass2.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San2")

            If ds.Tables("San2").Rows(0).Item("Expr1") Is DBNull.Value Then
                ud_ReturnData = 0
            Else
                ud_ReturnData = ds.Tables("San2").Rows(0).Item("Expr1")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub ReturnNodesFromServerSetupFile1(ByVal Str As String)
        Dim L1 As Integer
        Dim StrArray(255) As String
        'Dim LoopControl As Boolean

        Str = Trim(Str)
        L1 = Mid(Str, Len(Trim(Str)) - 1, 1)

        FinalNodes = Mid(Str, 13, L1)
    End Sub
    Public Sub ReturnMACAddress0()
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
        ud_MACAddress0 = Trim(nics(0).GetPhysicalAddress.ToString())
    End Sub
    Public Sub ReturnMACAddress1()
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
        ud_MACAddress1 = Trim(nics(1).GetPhysicalAddress.ToString())
    End Sub
    Public Sub TableCreationJKMergeChits(ByVal ConnString As SqlClient.SqlConnection)
        Try
            'to check if table exist in Database   i.e Table:  JKMergeChits
            TableName = "JKMergeChits"
            ObjClass1.TableExistence(TableName, ObjClass1.Myconn)
            If ud_TableFound = True Then
                QUERY = "Delete From  " & TableName
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)

                'QUERY = "DROP TABLE " & TableName
                'ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)

                'QUERY = "SELECT * INTO " & TableName & " FROM ChitSetupTable"
                'ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
            Else
                QUERY = "SELECT * INTO " & TableName & " FROM ChitSetupTable"
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
            End If

            If ConnString.State = ConnectionState.Open Then
                ConnString.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub TableCreationJKMergeData(ByVal ConnString As SqlClient.SqlConnection)
        Try
            'to check if table exist in Database   i.e Table:  JKMergeData
            TableName = "JKMergeData"
            ObjClass1.TableExistence(TableName, ObjClass1.Myconn)
            If ud_TableFound = True Then
                QUERY = "Delete From  " & TableName
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)

                'QUERY = "SELECT * INTO " & TableName & " FROM JkMasterTable"
                'ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
            Else
                QUERY = "SELECT * INTO " & TableName & " FROM JKMasterTable"
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
            End If

            If ConnString.State = ConnectionState.Open Then
                ConnString.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub TableCreationJkRecPay(ByVal ConnString As SqlClient.SqlConnection)
        Try
            'to check if table exist in Database   i.e Table:  JKMergeData
            TableName = "JKRecPay"
            ObjClass1.TableExistence(TableName, ObjClass1.Myconn)
            If ud_TableFound = True Then
                QUERY = "DROP TABLE " & TableName
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)

                QUERY = "SELECT * INTO " & TableName & " FROM JKRecPayMaster"
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
            Else
                QUERY = "SELECT * INTO " & TableName & " FROM JKRecPayMaster"
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
            End If

            If ConnString.State = ConnectionState.Open Then
                ConnString.Close()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub ReturnData(ByVal QueryString As String, ByVal ConnString As SqlClient.SqlConnection)

        cmd = New SqlClient.SqlCommand(QueryString, ObjClass1.Myconn)
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        Try
            If cmd.ExecuteScalar Is DBNull.Value Then
                ReturnValue = 0
            Else
                ReturnValue = cmd.ExecuteScalar
            End If

            ObjClass1.Myconn.Close()
        Catch ex As Exception
            MsgBox("Class: RetrnData" + ex.Message)
        End Try
    End Sub
    Public Sub ReturnStringData(ByVal QueryString As String, ByVal ConnString As SqlClient.SqlConnection)

        cmd = New SqlClient.SqlCommand(QueryString, ObjClass1.Myconn)
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        Try
            If cmd.ExecuteScalar Is DBNull.Value Then
                ReturnString = ""
            Else
                ReturnString = cmd.ExecuteScalar
            End If

            ObjClass1.Myconn.Close()
        Catch ex As Exception
            MsgBox("Class: RetrnStringData" + ex.Message)
        End Try
    End Sub
    Public Sub ReturnDateData(ByVal QueryString As String, ByVal ConnString As SqlClient.SqlConnection)

        cmd = New SqlClient.SqlCommand(QueryString, ObjClass1.Myconn)
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        Try
            If cmd.ExecuteScalar Is DBNull.Value Then
                ReturnDateValue = Now.Date
            Else
                ReturnDateValue = cmd.ExecuteScalar
            End If

            ObjClass1.Myconn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub FireSMS(ByVal rmn1 As String, ByVal msg As String)
        Dim strUrl As String = SMSURL & _
       "authkey=" & AutorizationKey & _
       "&mobiles=" & rmn1 & _
       "&message=" & msg & _
       "&sender=" & SenderID & _
       "&route=6" & _
       "&country=91"
        Try
            Dim request As WebRequest = HttpWebRequest.Create(strUrl)
            Dim response As HttpWebResponse = DirectCast(request.GetResponse, HttpWebResponse)
            Dim s As Stream = DirectCast(response.GetResponseStream(), Stream)
            Dim readStream As New StreamReader(s)
            Dim dataString As String = readStream.ReadToEnd()
            'RichTextBox1.Text = dataString.ToString
            response.Close()
            s.Close()
            readStream.Close()
            'MsgBox("SMS Delivered to RMN : " & rmn1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub
    Public Sub ReturnRegisteredMobileNumber(ByVal ClientCode As String, ByVal ConnString As SqlClient.SqlConnection)

        Query2 = "Select Cellno1 from ClientMaster WHERE ClientID='" & ClientCode & "'"

        da = New SqlClient.SqlDataAdapter(Query2, ConnString)
        ds = New DataSet
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        Try
            da.Fill(ds, "Pan1")
            If ds.Tables("Pan1").Rows(0).Item("CellNo1") Is DBNull.Value Then
                RMN = 0
            Else
                RMN = ds.Tables("Pan1").Rows(0).Item("CellNo1")
            End If
            ConnString.Close()
        Catch ex As Exception
            RMN = 0
            'MsgBox("Class: ReturnFieldValue " + ex.Message)
        End Try
    End Sub
    Public Sub ReturnFieldValue(ByVal QueryString As String, ByVal DataField As String, ByVal ConnString As SqlClient.SqlConnection)
        da = New SqlClient.SqlDataAdapter(QueryString, ConnString)
        ds = New DataSet
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        Try
            da.Fill(ds, "Pan1")
            If ds.Tables("Pan1").Rows(0).Item(DataField) Is DBNull.Value Then
                ud_FieldValue = 0
            Else
                ud_FieldValue = ds.Tables("Pan1").Rows(0).Item(DataField)
            End If
            ConnString.Close()
        Catch ex As Exception
            ud_FieldValue = 0
            'MsgBox("Class: ReturnFieldValue " + ex.Message)
        End Try
    End Sub
    Public Sub RecordExist(ByVal QueryString As String, ByVal ConnString As SqlClient.SqlConnection)
        da = New SqlClient.SqlDataAdapter(QueryString, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "Wan1")
            Count = ds.Tables("Wan1").Rows.Count
            If Count > 0 Then
                ud_RecordExist = True
            Else
                ud_RecordExist = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub
    Public Sub CountData(ByVal QueryString As String, ByVal ConnString As SqlClient.SqlConnection)
        da = New SqlClient.SqlDataAdapter(QueryString, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            ud_RecordCount = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub
    Public Sub CheckData(ByVal QueryString As String, ByVal CheckCode As String, ByVal TableField As String)
        Count = 0
        Count1 = 0
        Dim st1 As String
        da = New SqlClient.SqlDataAdapter(QueryString, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
            rno = 0
            Do While rno <= Count - 1
                st1 = ds.Tables("San1").Rows(rno).Item(TableField)
                If CheckCode <> st1 Then
                    rno = rno + 1
                Else
                    Count1 = Count1 + 1
                    Exit Do
                End If
            Loop

            If Count1 > 0 Then
                ud_RecordFound = True
            Else
                ud_RecordFound = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
   
    Public Sub ReturnFromServerSetupFile1(ByVal Str As String)
        Dim L1 As Integer
        Dim StrArray(255) As String
        'Dim LoopControl As Boolean

        Str = Trim(Str)
        L1 = Mid(Str, Len(Trim(Str)) - 1, 2)

        FinalString = Mid(Str, 3, L1)
    End Sub
    Public Sub GetSetActiveNode(ByVal LL As Integer)
        QUERY = "Select * from NODES"
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        da.Fill(ds, "San1")
        ud_non = ds.Tables("San1").Rows(0).Item("ActiveNode")
        If LL = 1 Then
            QUERY = "UPDATE NODES set ActiveNode=" & ud_non + 1
            ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
        End If

        If LL = 2 Then
            QUERY = "UPDATE NODES set ActiveNode=" & ud_non - 1
            ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
        End If
    End Sub
    Public Sub DataGridRefresh(ByVal QueryString As String, ByVal NameofDatagrid As Object, ByVal connstring As SqlClient.SqlConnection)
        da = New SqlClient.SqlDataAdapter(QueryString, connstring)
        ds = New DataSet
        Try
            da.Fill(ds, "San5")
            NameofDatagrid.DataSource = ds.Tables("San5")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub InsertUpdateDelete(ByVal AUDQuery As String, ByVal connstring As SqlClient.SqlConnection)
        IUDFlag = False
        Dim cmd As New SqlClient.SqlCommand(AUDQuery, connstring)
        If connstring.State = ConnectionState.Closed Then
            connstring.Open()
        End If
        Try
            cmd.ExecuteNonQuery()
            connstring.Close()
            IUDFlag = True
            ' Throw New MyException("User Exception: Will Create duplicate values")
        Catch ex As Exception
            'Add your parameter
            'cmd.Parameters.Add("Expr1", SqlDbType.BigInt)
            'specify it is an output parameter
            'cmd.Parameters("Expr1").Direction = ParameterDirection.ReturnValue
            'aa = cmd.Parameters("Expr1").Value

            IUDFlag = False
            MsgBox("Class1 (InsertUpdateDelete)" & Chr(13) & ex.Message)
        End Try

    End Sub

  
    Public Sub MyDataReader(ByVal Query As String, ByVal ConnString As SqlClient.SqlConnection, ByVal ControlName As Object, ByVal AttributeName As String)
        ' for single attribute name
        Dim TempDr As SqlClient.SqlDataReader
        Dim TempCmd As New SqlClient.SqlCommand(Query, ConnString)
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        TempDr = TempCmd.ExecuteReader
        With TempDr
            Do While .Read
                ControlName.items.add(TempDr.Item(AttributeName))
            Loop
        End With
        ControlName.SelectedItem = ControlName.items(0)
        TempDr.Close()
        ConnString.Close()
    End Sub
    Public Sub MyDataReader1(ByVal Query As String, ByVal ConnString As SqlClient.SqlConnection, ByVal ControlName As Object, ByVal AttributeName1 As String, ByVal AttributeName2 As String)
        ' for Two attribute name
        Dim TempDr As SqlClient.SqlDataReader
        Dim TempCmd As New SqlClient.SqlCommand(Query, ConnString)
        If ConnString.State = ConnectionState.Closed Then
            ConnString.Open()
        End If
        TempDr = TempCmd.ExecuteReader
        With TempDr
            Do While .Read
                ControlName.items.add(TempDr.Item(AttributeName1) & " # " & TempDr.Item(AttributeName2))
            Loop
        End With
        ControlName.SelectedItem = ControlName.items(0)
        TempDr.Close()
        ConnString.Close()
    End Sub
    Public Sub TableExistence(ByVal Tablename1 As String, ByVal ConStr As SqlClient.SqlConnection)
        ud_TableFound = False

        Dim Query As String = "Select * from " & Tablename1
        Dim cmd2 As New SqlClient.SqlCommand(Query, ConStr)

        If ConStr.State = ConnectionState.Closed Then
            ConStr.Open()
        End If
        Try
            ud_TableFound = cmd2.ExecuteNonQuery()
        Catch ex As Exception
            ud_TableFound = False
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub FileExistence(ByVal TableName1 As String, ByVal TableTemplateName As String, ByVal PrimaryKeyValue As String, ByVal ConStr As SqlClient.SqlConnection)

        Dim drt2 As Boolean
        Dim Query As String = "Select * from " & TableName1
        Dim cmd2 As New SqlClient.SqlCommand(Query, ConStr)

        If ConStr.State = ConnectionState.Closed Then
            ConStr.Open()
        End If
        Try
            drt2 = cmd2.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Process is Creating New Table")
        End Try


        If drt2 = False Then
            Query = "SELECT * INTO " & TableName1 & " FROM " & TableTemplateName
            'Query = "Create Table " & SemBranchSession & " as Select * from " & TableTemplateName
            Dim cmd3 As New SqlClient.SqlCommand(Query, ConStr)

            Try
                cmd3.ExecuteNonQuery()
                cmd3.Dispose()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Query = "Alter table " & TableName1 & " ADD PRIMARY KEY(" & PrimaryKeyValue & ")"
            Dim cmd4 As New SqlClient.SqlCommand(Query, ConStr)
            Try
                cmd4.ExecuteNonQuery()
                cmd4.Dispose()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        cmd2.Dispose()

        ConStr.Close()

    End Sub
  
    'Public Sub CalculateGDN(ByVal t1, ByVal t2, ByVal t3, ByVal t4, ByVal t5, ByVal t6, ByVal t7, ByVal t8, ByVal t9, ByVal t10, ByVal t11, ByVal t12, ByVal t13, ByVal t14, ByVal t15, ByVal t16, ByVal t17, ByVal t18, ByVal t19, ByVal t20, ByVal t21, ByVal t22, ByVal t23, ByVal t24, ByVal t25, ByVal t26, ByVal t27, ByVal t28, ByVal t29, ByVal t30, ByVal t31, ByVal t32)
    Public Sub calculateGDN(ByVal Array1 As Array)
        GrossPay = 0
        TotalDed = 0
        For i As Short = 1 To 10
            GrossPay = GrossPay + Array1(i)
        Next
        For i As Short = 11 To 32
            TotalDed = TotalDed + Array1(i)
        Next
        GrossPay = GrossPay + Array1(33)
        GrossPay = GrossPay + Array1(34)
        NetPay = GrossPay - TotalDed

    End Sub
    Public Sub FillDescription(ByVal ComboObject As Object, ByVal TableName As String)
        Dim Count, rno As Short
        QUERY = "SELECT distinct(description) from " & TableName
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If Count > 0 Then
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    ComboObject.Items.Add(.Item("Description"))
                End With
                rno = rno + 1
            Loop
            ComboObject.SelectedItem = ComboObject.Items(0)
        End If
    End Sub
    Public Sub FillEmployeeCombo(ByVal ComboObject As Object)
        Dim Count, rno As Short
        QUERY = "SELECT EmployeeCode,EmployeeName from SALEmployeeMaster"
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If Count > 0 Then
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    ComboObject.Items.Add(.Item("EmployeeCode") & " # " & .Item("EmployeeName"))
                End With
                rno = rno + 1
            Loop
            ComboObject.SelectedItem = ComboObject.Items(0)
        End If
    End Sub
    Public Sub FillComboDeduction(ByVal ComboObject1 As Object, ByVal ComboObject2 As Object, ByVal ComboObject3 As Object, ByVal ComboObject4 As Object, ByVal ComboObject5 As Object, ByVal ComboObject6 As Object)
        Dim Count, rno As Short
        QUERY = "Select * from SALDeductionMaster"
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        da.Fill(ds, "San1")
        Count = ds.Tables("San1").Rows.Count
        rno = 0
        ComboObject1.items.clear()
        Do While rno <= Count - 1
            With ds.Tables("San1").Rows(rno)
                ComboObject1.Items.Add(.Item("dedcode"))
                ComboObject2.Items.Add(.Item("dedcode"))
                ComboObject3.Items.Add(.Item("dedcode"))
                ComboObject4.Items.Add(.Item("dedcode"))
                ComboObject5.Items.Add(.Item("dedcode"))
                ComboObject6.Items.Add(.Item("dedcode"))

            End With
            rno = rno + 1
        Loop
    End Sub
    Public Sub ReturnChitMonthOrder(ByVal MonthName As String)
        If UCase(MonthName) = "MAR-" Then
            MOrder = 3
        End If
        If UCase(MonthName) = "APR-" Then
            MOrder = 4
        End If
        If UCase(MonthName) = "MAY-" Then
            MOrder = 5
        End If
        If UCase(MonthName) = "JUN-" Then
            MOrder = 6
        End If
        If UCase(MonthName) = "JUL-" Then
            MOrder = 7
        End If
        If UCase(MonthName) = "AUG-" Then
            MOrder = 8
        End If
        If UCase(MonthName) = "SEP-" Then
            MOrder = 9
        End If
        If UCase(MonthName) = "OCT-" Then
            MOrder = 10
        End If
        If UCase(MonthName) = "NOV-" Then
            MOrder = 11
        End If
        If UCase(MonthName) = "DEC-" Then
            MOrder = 12
        End If
        If UCase(MonthName) = "JAN-" Then
            MOrder = 1
        End If
        If UCase(MonthName) = "FEB-" Then
            MOrder = 2
        End If

    End Sub
   
    Public Sub MonthOrder(ByVal MonthName As String)
        If UCase(MonthName) = "MAR" Then
            MOrder = 1
        End If
        If UCase(MonthName) = "APR" Then
            MOrder = 2
        End If
        If UCase(MonthName) = "MAY" Then
            MOrder = 3
        End If
        If UCase(MonthName) = "JUN" Then
            MOrder = 4
        End If
        If UCase(MonthName) = "JUL" Then
            MOrder = 5
        End If
        If UCase(MonthName) = "AUG" Then
            MOrder = 6
        End If
        If UCase(MonthName) = "SEP" Then
            MOrder = 7
        End If
        If UCase(MonthName) = "OCT" Then
            MOrder = 8
        End If
        If UCase(MonthName) = "NOV" Then
            MOrder = 9
        End If
        If UCase(MonthName) = "DEC" Then
            MOrder = 10
        End If
        If UCase(MonthName) = "JAN" Then
            MOrder = 11
        End If
        If UCase(MonthName) = "FEB" Then
            MOrder = 12
        End If

    End Sub
    Public Sub FillCombo(ByVal NameofCombo As Object, ByVal Query As String)
        ud_RecordFound = False
        ObjClass1.TestConnection(ServerName)
        da = New SqlClient.SqlDataAdapter(Query, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox("Class1 : FillCombo " & Chr(13) & ex.Message)
        End Try
        If Count > 0 Then
            NameofCombo.items.clear()
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    NameofCombo.items.add(.Item("ChitCode"))
                    ud_RecordFound = True
                End With
                rno = rno + 1
            Loop
        End If

    End Sub
    Public Sub FillComboMonth(ByVal NameofCombo As Object, ByVal Query As String)
        ud_RecordFound = False
        ObjClass1.TestConnection(ServerName)
        da = New SqlClient.SqlDataAdapter(Query, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox("Class1 : FillCombo " & Chr(13) & ex.Message)
        End Try
        If Count > 0 Then
            NameofCombo.items.clear()
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    NameofCombo.items.add(.Item("PaymentMonth"))
                    ud_RecordFound = True
                End With
                rno = rno + 1
            Loop
        End If

    End Sub
    Public Sub FillComboCity(ByVal NameofCombo As Object, ByVal Query As String)
        ud_RecordFound = False
        ObjClass1.TestConnection(ServerName)
        da = New SqlClient.SqlDataAdapter(Query, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox("Class1 : FillComboCity " & Chr(13) & ex.Message)
        End Try
        If Count > 0 Then
            NameofCombo.items.clear()
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    NameofCombo.items.add(.Item("CityName"))
                    ud_RecordFound = True
                End With
                rno = rno + 1
            Loop
        End If

    End Sub
    Public Sub CalculateTotalAmt(ByVal Tablename As String, ByVal Fieldname As String, ByVal Code As String, ByVal Constr As SqlClient.SqlConnection)
        QUERY = "Select Sum(" & Fieldname & ") from " & Tablename & " Where EmployeeCode='" & Code & "' AND WITHELD='N'"
        cmd = New SqlClient.SqlCommand(QUERY, Constr)
        If Constr.State = ConnectionState.Closed Then
            Constr.Open()
        End If
        Try
            ReturnTotalAmt = cmd.ExecuteScalar
        Catch ex As Exception
            ReturnTotalAmt = 0

        End Try
        Constr.Close()

    End Sub
    Public Sub ITFillDeduction(ByVal NameofCombo As Object, ByVal Query As String)
        ObjClass1.TestConnection(ServerName)
        da = New SqlClient.SqlDataAdapter(Query, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox("Class1 : FillCombo " & Chr(13) & ex.Message)
        End Try
        If Count > 0 Then
            NameofCombo.items.clear()
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    NameofCombo.items.add(.Item("DeductionName"))
                End With
                rno = rno + 1
            Loop
        End If

    End Sub
    Public Sub ITFillS80C(ByVal NameofCombo As Object, ByVal Query As String)
        ObjClass1.TestConnection(ServerName)
        da = New SqlClient.SqlDataAdapter(Query, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
        Catch ex As Exception
            MsgBox("Class1 : FillCombo " & Chr(13) & ex.Message)
        End Try
        If Count > 0 Then
            NameofCombo.items.clear()
            rno = 0
            Do While rno <= Count - 1
                With ds.Tables("San1").Rows(rno)
                    NameofCombo.items.add(.Item("Section80CName"))
                End With
                rno = rno + 1
            Loop
        End If

    End Sub
    Public Sub ReturnMaxNumber(ByVal Tablename As String, ByVal Fieldname As String, ByVal Funct As String)
        QUERY = "Select " & Funct & "(" & Fieldname & ") FROM " & Tablename
        cmd = New SqlClient.SqlCommand(QUERY, ObjClass1.Myconn)
        If ObjClass1.Myconn.State = ConnectionState.Closed Then
            ObjClass1.Myconn.Open()
        End If

        Try
            ReturnValue = cmd.ExecuteScalar
            ObjClass1.Myconn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Function ReturnCurrentMonth() As String
        Dim sCurrMon As String
        sCurrMon = ""
        QUERY = "select top 1 month from salteachingmaster"
        cmd = New SqlClient.SqlCommand(QUERY, ObjClass1.Myconn)
        If ObjClass1.Myconn.State = ConnectionState.Closed Then
            ObjClass1.Myconn.Open()
        End If
        Try
            sCurrMon = cmd.ExecuteScalar
            ObjClass1.Myconn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sCurrMon
    End Function
    Public Sub SetFormToolBoxControlBackColor(ByVal ColorString As String, ByVal FormORToolboxControl As Object)
        FormORToolboxControl.backcolor = CType(TypeDescriptor.GetConverter(GetType(Color)).ConvertFromString(ColorString), Color)
    End Sub
    Public Function CheckMonthValidity() As Boolean
        Dim sCurrMon As String
        Dim sSessionYear As String
        sCurrMon = ObjClass1.ReturnCurrentMonth()
        If (UCase(Mid$(sCurrMon, 1, 3)) = "JAN") Or (UCase(Mid$(sCurrMon, 1, 3)) = "FEB") Then
            sSessionYear = "20" & Mid$(sCurrMon, 5, 2)
            sSessionYear = Trim(Str(Val(sSessionYear) - 1)) & "-" & sSessionYear
        Else
            sSessionYear = "20" & Mid$(sCurrMon, 5, 2)
            sSessionYear = sSessionYear & "-" & Trim(Str(Val(sSessionYear) + 1))
        End If
        If sSessionYear = Session Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Sub ProcessAutoMergeData()

        'merging data
        ObjClass1.TableCreationJKMergeData(ObjClass1.Myconn)

        ArrIndex = 0
        number = 1
        QUERY = "Select Chitcode from chitmaster where chitactive=1"
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
            rno = 0
            Do While rno <= Count - 1
                StrArr1(rno) = ds.Tables("San1").Rows(rno).Item("ChitCode")
                rno = rno + 1
            Loop
            ArrIndex1 = rno
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        TableName = "JKMergeData"

        Try
            rno = 0
            Do While rno <= ArrIndex1 - 1
                ChitTableName = "CHIT_" & Mid(StrArr1(rno), 1, 4) + Mid(StrArr1(rno), 6, 2) + Mid(StrArr1(rno), 9, 2) + Mid(StrArr1(rno), 12, 2)
                QUERY = "INSERT INTO " & TableName & _
                " (Date,ClientID,Groupno,PaymentMonth,InstallmentNo,Amount,Dividend,Interest,ModeofPayment,ChqNo,ChqDate,BankName,BookSrNo,SubscriberNo,RegCharges,AgentID,Remarks,MonthOrder,Type,Intrate,IntAmt,Legal,Visiting,Other,TotalAmt,PayRecd,ExecutiveID, Forfeited,CollectionMode,TransferTo,Rebate ) " & _
                " SELECT  Date,ClientID,Groupno,PaymentMonth,InstallmentNo,Amount,Dividend,Interest,ModeofPayment,ChqNo,ChqDate,BankName,BookSrNo,SubscriberNo,RegCharges,AgentID,Remarks,MonthOrder,Type,Intrate,IntAmt,Legal,Visiting,Other,TotalAmt,PayRecd,ExecutiveID, Forfeited,CollectionMode,TransferTo,Rebate from " & ChitTableName
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
                rno = rno + 1
            Loop
            'MsgBox("Table Created Successflly!")

            'To Delete all Records with ClientID=TMP... (Temporary Clients)
            QUERY = "Delete from JKMergeData where SubString(Clientid,1,3)='TMP'"
            ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Sub ProcessAutoMergeDataBranch(ByVal CB As String)

        'merging data
        ObjClass1.TableCreationJKMergeData(ObjClass1.Myconn)

        ArrIndex = 0
        number = 1
        QUERY = "Select Chitcode from chitmaster where ( chitactive=1 AND ChitBranch='" & CB & "')"

        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
            rno = 0
            Do While rno <= Count - 1
                StrArr1(rno) = ds.Tables("San1").Rows(rno).Item("ChitCode")
                rno = rno + 1
            Loop
            ArrIndex1 = rno
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        TableName = "JKMergeData"

        Try
            rno = 0
            Do While rno <= ArrIndex1 - 1
                ChitTableName = "CHIT_" & Mid(StrArr1(rno), 1, 4) + Mid(StrArr1(rno), 6, 2) + Mid(StrArr1(rno), 9, 2) + Mid(StrArr1(rno), 12, 2)
                QUERY = "INSERT INTO " & TableName & _
                " (Date,ClientID,Groupno,PaymentMonth,InstallmentNo,Amount,Dividend,Interest,ModeofPayment,ChqNo,ChqDate,BankName,BookSrNo,SubscriberNo,RegCharges,AgentID,Remarks,MonthOrder,Type,Intrate,IntAmt,Legal,Visiting,Other,TotalAmt,PayRecd,ExecutiveID, Forfeited,CollectionMode,TransferTo,Rebate ) " & _
                " SELECT  Date,ClientID,Groupno,PaymentMonth,InstallmentNo,Amount,Dividend,Interest,ModeofPayment,ChqNo,ChqDate,BankName,BookSrNo,SubscriberNo,RegCharges,AgentID,Remarks,MonthOrder,Type,Intrate,IntAmt,Legal,Visiting,Other,TotalAmt,PayRecd,ExecutiveID, Forfeited,CollectionMode,TransferTo,Rebate from " & ChitTableName
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
                rno = rno + 1
            Loop
            'MsgBox("Table Created Successflly!")
            'To Delete all Records with ClientID=TMP... (Temporary Clients)
            QUERY = "Delete from JKMergeData where SubString(Clientid,1,3)='TMP'"
            ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Sub ProcessAutoMergeChits()

        'merging data
        ObjClass1.TableCreationJKMergeChits(ObjClass1.Myconn)

        ArrIndex = 0
        number = 1
        QUERY = "Select Chitcode from chitmaster where chitactive=1"
        da = New SqlClient.SqlDataAdapter(QUERY, ObjClass1.Myconn)
        ds = New DataSet
        Try
            da.Fill(ds, "San1")
            Count = ds.Tables("San1").Rows.Count
            rno = 0
            Do While rno <= Count - 1
                StrArr1(rno) = ds.Tables("San1").Rows(rno).Item("ChitCode")
                rno = rno + 1
            Loop
            ArrIndex1 = rno
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        TableName = "JKMergeChits"

        Try
            rno = 0
            Do While rno <= ArrIndex1 - 1
                ChitTableName = Mid(StrArr1(rno), 1, 4) + Mid(StrArr1(rno), 6, 2) + Mid(StrArr1(rno), 9, 2) + Mid(StrArr1(rno), 12, 2)
                QUERY = "INSERT INTO " & TableName & _
                " (Groupno,ClientID,AgentID,Date,Status,ActiveStatus,Remarks,SubscriberNo,PayAmount,PayMode,PayChqNo,PayBankName,PayRecBookNo,PayChqDate,ChitAuction,NoOfChitAuction,ExecutiveID,DocAmt,VerificationAmt,OtherAmt,AdjustAmt,TotalofPayAmount ) " & _
                " SELECT  Groupno,ClientID,AgentID,Date,Status,ActiveStatus,Remarks,SubscriberNo,PayAmount,PayMode,PayChqNo,PayBankName,PayRecBookNo,PayChqDate,ChitAuction,NoOfChitAuction,ExecutiveID,DocAmt,VerificationAmt,OtherAmt,AdjustAmt,TotalofPayAmount from " & ChitTableName
                ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
                rno = rno + 1
            Loop
            'MsgBox("Table Created Successflly!")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Sub ProcessAutoInterest(ByVal CurrDate As Date, ByVal PrizeInt As Short, ByVal NonPrizeInt As Short)
        GQuery = "Select  Chitcode from ChitMaster where chitactive=1"
        'temporary for testing
        'GQuery = "Select  Chitcode from ChitMaster where (chitactive=1 and chitcode='JKCH-01-05-15')"

        da = New SqlClient.SqlDataAdapter(GQuery, ObjClass2.Myconn)
        ds = New DataSet
        da.Fill(ds, "INT1")
        CountInt = ds.Tables("INT1").Rows.Count
        RowInt = 0
        Do While RowInt <= CountInt - 1
            CC = ds.Tables("INT1").Rows(RowInt).Item("Chitcode")

            ChitTableName = "CHIT_" & Mid(CC, 1, 4) + Mid(CC, 6, 2) + Mid(CC, 9, 2) + Mid(CC, 12, 2)
            TableName = Mid(CC, 1, 4) + Mid(CC, 6, 2) + Mid(CC, 9, 2) + Mid(CC, 12, 2)


            Try
                Query3 = "Select ClientID from " & TableName
                da2 = New SqlClient.SqlDataAdapter(Query3, ObjClass2.Myconn)
                ds2 = New DataSet
                da2.Fill(ds2, "Wan1")
                CountWan1 = ds2.Tables("Wan1").Rows.Count
                RWan1 = 0
                Do While RWan1 < CountWan1 - 1
                    CCode = ds2.Tables("Wan1").Rows(RWan1).Item("ClientID")

                    Array1 = 0
                    Array2 = 0


                    QUERY = "Select * from " & ChitTableName & " Where ( ClientID='" & CCode & "' AND Date<='" & Format(CurrDate, "dd MMM yyyy") & "' AND PAyRECD=0)"

                    da1 = New SqlClient.SqlDataAdapter(QUERY, ObjClass3.Myconn)
                    ds1 = New DataSet
                    da1.Fill(ds1, "Tan4")
                    rno = 0
                    Array1 = 0
                    Count = ds1.Tables("Tan4").Rows.Count
                    RCount = Count
                    'IR = 1
                    IR = Count
                    Do While rno <= Count - 1
                        SerialNumber = ds1.Tables("Tan4").Rows(rno).Item("SNO")
                        ReturnDate = ds1.Tables("Tan4").Rows(rno).Item("Date")
                        ud_Amount = ds1.Tables("Tan4").Rows(rno).Item("amount")
                        ud_Status = ds1.Tables("Tan4").Rows(rno).Item("type")
                        CustID = ds1.Tables("Tan4").Rows(rno).Item("ClientID")



                        If Trim(ud_Status) = "PRIZE (REGULAR)" Or Trim(ud_Status) = "PRIZE (CANCEL)" Or Trim(ud_Status) = "PRIZE (DEFAULT)" Or Trim(ud_Status) = "PRIZE (BIDDER)" Then
                            IntRate = PrizeInt
                            'IntIncr = IntRate
                        End If
                        If Trim(ud_Status) = "NON-PRIZE (REGULAR)" Or Trim(ud_Status) = "NON-PRIZE (CANCEL)" Or Trim(ud_Status) = "NON-PRIZE (DEFAULT)" Then
                            IntRate = NonPrizeInt
                            ' IntIncr = IntRate
                        End If


                        InterestAmt = (ud_Amount * (IntRate * IR)) / 100

                        QUERY = "Update " & ChitTableName & " SET " & _
                                "Intrate=" & (IntRate * IR) & "," & _
                                "IntAmt=" & InterestAmt & "," & _
                                "Interest=" & InterestAmt & "," & _
                                "TotalAmt=1" & _
                                "  WHERE sno=" & SerialNumber
                        ObjClass1.InsertUpdateDelete(QUERY, ObjClass1.Myconn)
                        If ObjClass1.Myconn.State = ConnectionState.Open Then
                            ObjClass1.Myconn.Close()
                        End If


                        rno = rno + 1
                        'IR = IR + 1
                        IR = IR - 1
                    Loop
                    RWan1 = RWan1 + 1
                Loop



            Catch ex As Exception
                MsgBox("ChitInterestAll Exception")
            End Try
            RowInt = RowInt + 1
        Loop

    End Sub
    

    Public Sub ProcessAutoDefaultStatus()

        'QUERY = "Select ChitAuction from " & TableName & " WHERE ClientID='" & Trim(txtClientCode.Text) & "'"
        'ObjClass1.ReturnData(QUERY, ObjClass1.Myconn)
        'Chit_Auction = ReturnValue


        ds_Status = ""

        'QUERY = "Select * from Chitmaster where chitactive=1"

        'temporary for testing
        QUERY = "Select * from chitmaster where (chitActive=1 and chitcode='JKCH-01-05-16')"

        da2 = New SqlClient.SqlDataAdapter(QUERY, ObjClass3.Myconn)
        ds2 = New DataSet

        If ObjClass3.Myconn.State = ConnectionState.Closed Then
            ObjClass3.Myconn.Open()
        End If

        da2.Fill(ds2, "Lan1")
        ds_Count1 = ds2.Tables("Lan1").Rows.Count
        ds_Rno1 = 0
        Do While ds_Rno1 <= ds_Count1 - 1
            ds_ChitGroup = ds2.Tables("Lan1").Rows(ds_Rno1).Item("ChitCode")
            ds_ChitAmt = ds2.Tables("Lan1").Rows(ds_Rno1).Item("ChitAmount")
            ds_ChitInstall = ds2.Tables("Lan1").Rows(ds_Rno1).Item("ChitInstallment")

            QUERY = "Select ClientID from ClientMaster"
            da1 = New SqlClient.SqlDataAdapter(QUERY, ObjClass2.Myconn)
            ds1 = New DataSet
            da1.Fill(ds1, "Lan2")
            ds_Count2 = ds1.Tables("Lan2").Rows.Count
            ds_Rno2 = 0
            Do While ds_Rno2 <= ds_Count2 - 1
                ds_ClientCode = ds1.Tables("Lan2").Rows(ds_Rno2).Item("ClientID")
                'check if client exist
                ChitTableName = "CHIT_" & Mid(ds_ChitGroup, 1, 4) + Mid(ds_ChitGroup, 6, 2) + Mid(ds_ChitGroup, 9, 2) + Mid(ds_ChitGroup, 12, 2)
                TableName = Mid(ds_ChitGroup, 1, 4) + Mid(ds_ChitGroup, 6, 2) + Mid(ds_ChitGroup, 9, 2) + Mid(ds_ChitGroup, 12, 2)


                Query1 = "Select ClientID from " & ChitTableName & " WHERE ClientID='" & ds_ClientCode & "'"
                ObjClass1.RecordExist(Query1, ObjClass1.Myconn)

                If ud_RecordExist = True Then

                    ds_Query2 = " Select Max(installmentno)  from " & ChitTableName & " where (clientid='" & ds_ClientCode & "' AND payrecd=1)"
                    ObjClass1.ReturnData(ds_Query2, ObjClass1.Myconn)
                    ds_InsNo = ReturnValue

                    ds_Query2 = " Select Max(Sno)  from " & ChitTableName & " where (clientid='" & ds_ClientCode & "' AND payrecd=1)"
                    ObjClass1.ReturnData(ds_Query2, ObjClass1.Myconn)
                    ds_SrNo = ReturnValue

                    ' to get Type or Status of client
                    ds_Query2 = "Select Type from " & ChitTableName & " WHERE SNO=" & ds_SrNo
                    ObjClass1.ReturnFieldValue(ds_Query2, "type", ObjClass1.Myconn)
                    ds_ClientType = ud_FieldValue

                    ds_Query2 = "Select GroupNo from " & ChitTableName & " where sno=" & ds_SrNo
                    ObjClass1.ReturnFieldValue(ds_Query2, "Groupno", ObjClass1.Myconn)
                    GroupNumber = ud_FieldValue


                    ds_Query2 = " Select Max(Date) as Expr1  from " & ChitTableName & " where (clientid='" & ds_ClientCode & "' AND payrecd=1)"
                    ObjClass1.ReturnDateData(ds_Query2, ObjClass1.Myconn)
                    ds_DateR = ReturnDateValue


                    If ds_ClientType = "NON-PRIZE (REGULAR)" Then
                        ds_Status = "NON-PRIZE (DEFAULT)"
                    End If

                    If ds_ClientType = "NON-PRIZE (DEFAULT)" Then
                        ds_Status = "NON-PRIZE (DEFAULT)"
                    End If

                    If ds_ClientType = "PRIZE (REGULAR)" Then
                        ds_Status = "PRIZE (DEFAULT)"
                    End If

                    If ds_ClientType = "PRIZE (BIDDER)" Then
                        ds_Status = "PRIZE (BIDDER)"
                    End If

                    QUERY = "Select ChitAuction from " & TableName & " WHERE ClientID='" & ds_ClientCode & "'"
                    ObjClass1.ReturnData(QUERY, ObjClass1.Myconn)
                    Chit_Auction = ReturnValue

                    ds_DateADD = ds_DateR.AddDays(ud_DefaultValue)

                    If (Format(Date.Now, "dd MMM yyyy") >= ds_DateADD) And ds_Status <> "" Then
                        ds_SNO1 = ds_SrNo + 1
                        ds_InsNo1 = ds_InsNo + 1
                        Do While ds_InsNo1 <= ds_ChitInstall
                            ds_Amt = ds_ChitAmt / ds_ChitInstall

                            If ud_PrizeScheme = "ID" Then
                                ForemenCommission = ((ds_ChitAmt * 5) / 100) / ds_ChitInstall
                                Dividend = (((ds_ChitAmt * Chit_Auction) / 100) / ds_ChitInstall) - ForemenCommission
                                ds_Amt = ((ds_ChitAmt / ds_ChitInstall) - Dividend)

                                ds_Query3 = "UPDATE " & ChitTableName & " SET AMount=" & ds_Amt & ",Dividend=" & Dividend & ",Type='" & ds_Status & "'" & _
                                            " WHERE Sno=" & ds_SNO1
                                ObjClass1.InsertUpdateDelete(ds_Query3, ObjClass1.Myconn)
                            Else
                                ds_Amt = ds_ChitAmt / ds_ChitInstall
                                Dividend = 0
                                ds_Query3 = "UPDATE " & ChitTableName & " SET AMount=" & ds_Amt & ",Dividend=" & Dividend & ",Type='" & ds_Status & "'" & _
                                            " WHERE Sno=" & ds_SNO1
                                ObjClass1.InsertUpdateDelete(ds_Query3, ObjClass1.Myconn)

                            End If
                            'ds_Query3 = "UPDATE " & ChitTableName & " SET AMount=" & ds_Amt & ",Dividend=0,Type='" & ds_Status & "'" & _
                            '" WHERE Sno=" & ds_SNO1
                            '                            ObjClass1.InsertUpdateDelete(ds_Query3, ObjClass1.Myconn)

                            ds_InsNo1 = ds_InsNo1 + 1
                            ds_SNO1 = ds_SNO1 + 1

                        Loop
                        Try
                            ds_Query3 = "UPDATE " & TableName & " SET Status='" & ds_Status & "' WHERE (Clientid='" & ds_ClientCode & "' AND groupno='" & GroupNumber & "')"
                            ObjClass1.InsertUpdateDelete(ds_Query3, ObjClass1.Myconn)
                        Catch ex As Exception
                            MsgBox("DefaultStatus.vb:: Update Query")
                        End Try

                    End If
                End If

                ds_Rno2 = ds_Rno2 + 1


            Loop


            ds_Rno1 = ds_Rno1 + 1

        Loop
    End Sub

End Class

'Public Class MyException : Inherits System.Exception
'    Sub New(ByVal message As String)
'        MyBase.new(message)
'    End Sub
'End Class