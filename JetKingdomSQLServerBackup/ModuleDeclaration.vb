Imports System.Net
Imports System.IO
Imports System.Drawing
Imports System.ComponentModel

Imports System.Net.NetworkInformation
Imports System.Data.SqlClient

Module ModuleDeclaration
    'Public myConnectionInfo As ConnectionInfo = New ConnectionInfo
    Public ud_LoginLogoff As Integer
    Public ud_non As Integer
    Public ud_DBName As String
    Public ud_NumberNodes As String
    Public FinalNodes As Integer
    Public LoggedCode As String
    Public Flag As Boolean
    Public FocusColor As String
    Public OddEven As String = "ODD"
    Public FormColor As String
    Public ButtonColor As String
    Public Session As String
    Public SessionYear As String
    Public ServerName As String = " "

    Public SuperAdmin As Boolean
    Public UserAdmin As Boolean
    Public DeleteAdmin As Boolean
    Public ReportAdmin As Boolean

    Public SalAdmin As Boolean
    Public LoggedUser As String
    Public ApplnPath As String
    Public ApplnDrive As String
    Public ObjClass1 As New ConnectionClass
    Public ObjClass2 As New ConnectionClass
    Public ObjClass3 As New ConnectionClass
    Public ObjMM As New MainMenu
    Public BackupLog As String
    Public SQLServerPassword As String
    Public BackupFileName As String
    Public FinalString As String
    ' Public ObjLogin As New Login
    Public da As SqlDataAdapter
    Public ds As DataSet
    Public da1 As SqlDataAdapter
    Public ds1 As DataSet
    Public da2 As SqlDataAdapter
    Public ds2 As DataSet

    Public cmdB As SqlCommandBuilder
    Public cmd As SqlCommand
    Public cmd1 As SqlCommand
    'Public rowno As Short
    Public dr As SqlClient.SqlDataReader
    Public QUERY As String
    Public Query1 As String
    Public Query2 As String
    Public Query3 As String
    Public tcode As String
    Public sess As String

    Public ReturnValue As Integer
    Public ReturnString As String

    Public TableName As String
    Public Filename As String
    Public ITFileName As String
    Public ITCFileName As String
    Public Department As String
  
    Public GrossPay As Integer
    Public TotalDed As Integer
    Public NetPay As Integer
    Public Array1(75) As Integer
    Public DAPER As Decimal
    Public HRAPER As Decimal
    Public CPFPER As Decimal
    Public MAXCPF As Integer
    Public DAPER6P As Decimal
    Public HRAPER6P As Decimal
    Public CPFPER6P As Decimal
    Public MAXCPF6P As Integer
    Public TNT As String  'teaching non-teaching staff variable
    Public CMonth As String 'process month ex. JAN-06
    Public CurrentMonth As String
    Public MOrder As Byte = Microsoft.VisualBasic.Month(Now)
    Public ReturnTotalAmt As Decimal

    Public ud_CompanyName As String
    Public ud_CompanyAddress As String
    Public ud_CompanyContact As String
    Public ud_CompanyDirector As String
    Public ud_CompanyEmail As String
    Public ud_CompanyRegNo As String
    Public ud_CompanyAct As String

    Public CounterValue As Integer
    Public MaxCounterValue As Integer

    Public IUDFlag As Boolean

    Public ud_MACAddress0 As String
    Public ud_MACAddress1 As String

    Public ud_RecordFound As Boolean
    Public ud_RecordCount As Integer
    Public ud_FieldValue As String
    Public ud_TableFound As Boolean
    Public ud_ReturnData As Integer
    Public ud_RecordExist As Boolean
    Public ReturnDateValue As Date
    Public ProcessInterestDate As Date
    Public ChitTableName As String
    Public ud_DefaultValue As Short  ' to set no. of days for making client 'DEFAULT'

    Public ReceiptNumber As Integer
    Public SRN As Integer  ' receipt number

    Public ud_BLN As String ' BYE LAW NUMBER
    Public ud_TRN As String ' TR Number

    Public ud_PrizeScheme As String   ' for Prize Customers, either to collect Divident+Installment or Only Installment
    Public ud_ProcessYesNo As String  ' for processing of "Default and Interest calculation

    Public RMN As String ' Registered Mobile Number for SMS
    Public SMS_Message As String ' for SMS
    Public AutorizationKey As String ' for sms
    Public SenderID As String ' for sms
    Public SMSURL As String ' URL for sending SMS

End Module
