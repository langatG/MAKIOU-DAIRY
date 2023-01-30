Attribute VB_Name = "General"
Option Explicit
Public CosCon As New ADODB.Connection
Public success As Boolean
Public Type Account_Details
    AccountNo As String
    AccountName As String
    OpeningBalance As Double
    currentbalance As Double
    NormalBalance As String
End Type
Public Function GetRecords(str As String)
    On Error GoTo Capture
    Set Rst = New ADODB.Recordset
    sql = mysql
    Set cn = CreateObject("adodb.connection")
    Provider = "MAZIWA"
    cn.Open Provider, "bi"
    Rst.Open str, cn
    success = True
    Exit Function
Capture:
    success = False
    MsgBox err.description
End Function

'Public Sub InitSubClass()
'Set colClass = New Collection
'End Sub
Public Function Get_Account_Details(accno As String, DataSource As String, _
errmsg As String) As Account_Details
    On Error GoTo SysError
    Dim rsAccounts As New Recordset
    ''Open_Database DataSource
    Set rsAccounts = oSaccoMaster.GetRecordset("Select * From GLSETUP Where AccNo='" & accno & "'")
    With rsAccounts
        If .State = adStateOpen Then
            If Not .EOF Then
                Get_Account_Details.AccountName = IIf(IsNull(!GlAccName), "", !GlAccName)
                Get_Account_Details.AccountNo = accno
                Get_Account_Details.currentbalance = IIf(IsNull(!CurrentBal), 0, !CurrentBal)
                Get_Account_Details.NormalBalance = IIf(IsNull(!NormalBal), "DR", IIf(!NormalBal <> "Credit", "DR", "CR"))
                Get_Account_Details.OpeningBalance = IIf(IsNull(!OpeningBal), 0, !OpeningBal)
            End If
        End If
    End With
    Exit Function
SysError:
    errmsg = err.description
    Get_Account_Details.AccountNo = ""
End Function
Public Function getGlBalance(accno As String, Startdate As Date, Enddate As Date) As Double
    On Error GoTo Capture
    Dim OBal As Double
    totalcr = 0
    totaldr = 0
    sql = "set dateformat dmy select gl.Normalbal,op.sumdr DR,op.sumcr CR,op.cbal,gl.GlAccType from dbo.UDF_GL_OpeningBalance ('" & accno & "','" & Enddate & "') op inner join glsetup gl on op.accno=gl.accno where gl.accno='" & accno & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
        OBal = Rst("cbal")
        getGlBalance = OBal
        totaldr = Rst("DR")
        totalcr = Rst("CR")
    End If
    success = True
    Exit Function
Capture:
    success = False
End Function
Public Function getGlPeriodicTrans(accno As String, Startdate As Date, Enddate As Date) As Double
    On Error GoTo Capture
    Dim OBal As Double
    totalcr = 0
    totaldr = 0
    sql = "set dateformat dmy select gl.Normalbal,op.sumdr DR,op.sumcr CR,op.Transbal,gl.GlAccType from UDF_GL_PeriodicTrans ('" & accno & "','" & Startdate & "','" & Enddate & "') op inner join glsetup gl on op.accno=gl.accno where gl.accno='" & accno & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
        OBal = Rst("TransBal")
        getGlPeriodicTrans = OBal
        totaldr = Rst("DR")
        totalcr = Rst("CR")
    End If
    success = True
    Exit Function
Capture:
    success = False
End Function


