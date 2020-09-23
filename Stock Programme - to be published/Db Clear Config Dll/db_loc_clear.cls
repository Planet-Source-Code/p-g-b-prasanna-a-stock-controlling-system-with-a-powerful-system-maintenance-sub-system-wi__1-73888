VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "db_loc_clear"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim vbsobj As Object
Public db_location As String
Const db_reg_path = "HKEY_CURRENT_USER\Software\StockControl\Database\Path"
Function Clear_Db_Location()
On Error GoTo Err
vbsobj.Regdelete (db_reg_path)
MsgBox "Database Location Cleared Successfully...", vbInformation
Exit Function
Err:
MsgBox "Error occurred while clearing the Database Configuration..." & vbCrLf & "dll process was failed...", vbCritical
End Function

Function Get_Db_Location()
On Error GoTo Err
db_location = vbsobj.RegRead(db_reg_path)
Exit Function
Err:
MsgBox "Error occurred while getting the Database location..." & vbCrLf & "dll process was failed...", vbCritical
End Function

Private Sub Class_Initialize()
Set vbsobj = CreateObject("Wscript.Shell")
End Sub
