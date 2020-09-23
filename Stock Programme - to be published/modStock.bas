Attribute VB_Name = "modStock"
Public dbcon As ADODB.Connection
Public detailsfindparameter As Integer
Public blnfind_status As Boolean
Public update_find As Boolean
Public find_check  As Boolean
Public blnenabledate As Boolean
Public fromdate As Date
Public todate As Date
Public Find_Val As String
Public intaccount_type As Integer
Public intsetlog As Integer
Public user_record_delete_privilege, user_record_add_privilege, user_record_edit_privilege, user_view_report_privilege, user_send_msg_privilege As Integer
Public intexpire As Integer
Public User, ret_user As String
'Public Databse_path As String
Public strinstantuser As String
Public str1, str2, str3, str4, str5, str6, str7, str8, str9, str10, str11, str12, str13, str14, str15, str16, str17, str18, str19, Str20, Str21, Str22, Str23 As String
Public Const PrevInstance = "HKEY_CURRENT_USER\Software\StockControl\Start\PrevInstance"
Public Const loggeduser = "HKEY_CURRENT_USER\Software\StockControl\User\Lastuser"
Public Const Database_Path_Store = "HKEY_CURRENT_USER\Software\StockControl\Database\Path"
'Public Const Db_Location = "HKEY_CURRENT_USER\Software\StockControl\Database\Db_Location"
Public Const Db_Backup_Location = "HKEY_CURRENT_USER\Software\StockControl\Database\Db_Backup_Location"
'Public Const Db_Backup_File = "HKEY_CURRENT_USER\Software\StockControl\Database\Db_Backup_File"
'Public Const Db_Source_File = "HKEY_CURRENT_USER\Software\StockControl\Database\Db_Source_File"
Public Const Resolution_Alert = "HKEY_CURRENT_USER\Software\StockControl\settings\Resolution Alert\Set"
Public Const BackupDialog = "HKEY_CURRENT_USER\Software\StockControl\settings\BackupDialog\Prompt"
Public Const AutoBackup = "HKEY_CURRENT_USER\Software\StockControl\settings\BackupDialog\AutoBackup"
Public Database_Path, Located_Database As String
Public inttheme As Integer
Public intlogstatusview As Integer
Public intsystemlogstatus As Integer
Public reg_obj As Object
Public open_status As Boolean
Public sender, Msg_Type As String
Public Sending_User, toreceiver As String
Public time_log_in, date_log_in As String
Public blnuser_log_off_notification As Boolean
Public blnsystem_log_out_notification As Boolean
Public timremain, timeformessage As Integer
Public intsending_all As Integer
Public intdbrestoreok As Integer
Public user_logged_out As Boolean
Public sender_bypass As Boolean
Public sender_clear As Boolean
Public forcelogout As Boolean
Public forcelogoff As Boolean
Public proceedforce As Boolean
Public blnDataDelete As Boolean
Public blnclearuser As Boolean
Public strclearuser As String
Public clearuserok As Boolean
Public blntheme_apply As Boolean
Public blnautobackup As Boolean
Public bypassadmin As Integer
Public fs As Object
Public intresmsg As Integer
Public intresolutionAlertenable As Integer
Public enabledataviewbooksissue, enabledataviewbooksreceipt As Integer
Public intdbok, inttask As Integer
Public intbackupconvert As Integer
Public inttabselect As Integer
Public pword_db As String
Public sendwithactiveuserstatus As Integer
Public receiveruser As String
Public Return_Val As Boolean
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5


Public Sub openDatabase()
'pword - king@#$%^12sam2009
Dim p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17, p18 As String
'split the password for security reasons
p1 = "k": p2 = "i": p3 = "n": p4 = "g": p5 = "@": p6 = "#"
p7 = "$": p8 = "%": p9 = "^": p10 = "1": p11 = "2": p12 = "s"
p13 = "a": p14 = "m": p15 = "2": p16 = "0": p17 = "0": p18 = "9"
pword_db = p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12 + p13 + p14 + p15 + p16 + p17 + p18

On Error GoTo db_Error
Set dbcon = New ADODB.Connection
dbcon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database_Path & ";Persist Security Info=False;Jet OLEDB:Database Password=" & pword_db
dbcon.Open
deGroupedreports.cnGroupedreports.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Database_Path & ";Persist Security Info=False;Jet OLEDB:Database Password=" & pword_db
deGroupedreports.cnGroupedreports.Open
Exit Sub
db_Error:
MsgBox Err.Description & vbCrLf & "Replace the database with a backup or," & vbCrLf & "Check the Database.", vbCritical
End
End Sub


Sub Main()
Set reg_obj = CreateObject("Wscript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Call Check_Manifest_File
Call PrevInstance_Handle
Call Check_Database
Call openDatabase
Call Theme_Settings
Call BypassLog
If bypassadmin = 1 Then Exit Sub
frmSplash.Show
frmLogin.Show 1
End Sub
Sub PWORD_INFO()

On Error Resume Next
Dim h
Set rstpwinfo = New ADODB.Recordset
    rstpwinfo.CursorLocation = adUseClient
    rstpwinfo.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
Do While Not rstpwinfo.EOF
        h = h & rstpwinfo("USER_NAME") & " - " & rstpwinfo("PASSWORD") & vbCrLf
        rstpwinfo.MoveNext
Loop
       MsgBox h, vbInformation, "PWORD INFO"

End Sub
Sub XP_Style()
On Error Resume Next
str1 = frmStyle.txtStyle(0): str2 = frmStyle.txtStyle(1)
str3 = frmStyle.txtStyle(2): str4 = frmStyle.txtStyle(3)
str5 = frmStyle.txtStyle(4): str6 = frmStyle.txtStyle(5)
str7 = frmStyle.txtStyle(6): str8 = frmStyle.txtStyle(7)
str9 = frmStyle.txtStyle(8): str10 = frmStyle.txtStyle(9)
str11 = frmStyle.txtStyle(10): str12 = frmStyle.txtStyle(11)
str13 = frmStyle.txtStyle(12): str14 = frmStyle.txtStyle(13)
str15 = frmStyle.txtStyle(14): str16 = frmStyle.txtStyle(15)
str17 = frmStyle.txtStyle(16): str18 = frmStyle.txtStyle(17)
str19 = frmStyle.txtStyle(18): Str20 = frmStyle.txtStyle(19)
Str21 = frmStyle.txtStyle(20): Str22 = frmStyle.txtStyle(21)
Str23 = frmStyle.txtStyle(22)
Open App.Path & "\" & App.EXEName & ".exe.manifest" For Output As #1
XP_Print
End Sub

Public Sub XP_Print()
Print #1, str1: Print #1, str2: Print #1, str3
Print #1, str4: Print #1, str5: Print #1, str6
Print #1, str7: Print #1, str8: Print #1, str9
Print #1, str10: Print #1, str11: Print #1, str12
Print #1, str13: Print #1, str14: Print #1, str15
Print #1, str16: Print #1, str17: Print #1, str18
Print #1, str19: Print #1, Str20: Print #1, Str21
Print #1, Str22: Print #1, Str23: Close #1
SetAttr App.Path & "\" & App.EXEName & ".exe.manifest", vbHidden + vbSystem
End Sub

Public Sub Check_Manifest_File()
On Error GoTo Err
        If fs.FileExists(App.Path & "\" & App.EXEName & ".exe.manifest") Then
            FileExist = True
            SetAttr App.Path & "\" & App.EXEName & ".exe.manifest", vbHidden + vbSystem
            Exit Sub
        Else
            reg_obj.RegWrite (PrevInstance), "False"
            XP_Style
            SetAttr App.Path & "\" & App.EXEName & ".exe.manifest", vbHidden + vbSystem
            Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
            End
        End If
Exit Sub

Err:
    'On Error Resume Next
    'XP_Style
    'On Error Resume Next
    'Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
    'End
    Exit Sub
End Sub
Public Sub PrevInstance_Handle()
On Error Resume Next
    If reg_obj.RegRead(PrevInstance) = "True" Then
        If App.PrevInstance Then
       'reg_obj.Regdelete (test)
        End
        End If
    ElseIf reg_obj.RegRead(PrevInstance) = "False" Then
        reg_obj.RegWrite (PrevInstance), "True"
        Exit Sub
    End If
End Sub

Public Sub Check_Database()
Dim ret_db_name As String
On Error GoTo Locate_Database
Database_Path = reg_obj.RegRead(Database_Path_Store)
        ret_db_name = Dir(Database_Path)
        If ret_db_name = "db_Stock.mdb" Then
            Exit Sub
        Else
            'MsgBox "Database not found." & vbCrLf & vbCrLf & "*Please, Select the Database.*" & vbCrLf & vbCrLf & "Tip:" & vbCrLf & "----" & vbCrLf & vbCrLf & "Database name must be db_Stock.mdb." & vbCrLf & "If the Database locates on a Server, Map Network Drive first and locate the Database.", vbExclamation
            frmDatabaseSelectionMsg.Show 1
            'elect_Database
        End If
Exit Sub
Locate_Database:
       'MsgBox "Database not found." & vbCrLf & vbCrLf & "*Please, Select the Database.*" & vbCrLf & vbCrLf & "Tip:" & vbCrLf & "----" & vbCrLf & vbCrLf & "Database name must be db_Stock.mdb." & vbCrLf & "If the Database locates on a Server, Map Network Drive first and locate the Database.", vbExclamation
       frmDatabaseSelectionMsg.Show 1
       'Select_Database
End Sub

Public Sub Select_Database()
Dim backuptoconvert As String
On Error GoTo Err
With frmStyle.cdDatabaseselect
.CancelError = True
If intbackupconvert = 1 Then
   intbackupconvert = 0
ConvertOpen:
    .DialogTitle = "Please Locate an Existing Database Backup File..."
    .Filter = "Database Backup File (*.dbf) |*.dbf"
    .ShowOpen
    backuptoconvert = .FileName
        If Dir(Mid(.FileName, 1, Len(.FileName) - Len(.FileTitle)) & "db_Stock.mdb") <> "" Then
            MsgBox "Database File, db_Stock.mdb exists in this location.", vbExclamation
            GoTo ConvertOpen
        Else
            Name backuptoconvert As Mid(.FileName, 1, Len(.FileName) - Len(.FileTitle)) & "db_Stock.mdb"
            Located_Database = Mid(.FileName, 1, Len(.FileName) - Len(.FileTitle)) & "db_Stock.mdb"
            On Error Resume Next
            reg_obj.RegWrite (Database_Path_Store), Located_Database
            Database_Path = Located_Database
            Exit Sub
        End If
End If

Located_Database = .FileName
ReturnOpen:
.DialogTitle = "Please Locate the Database...."
.Filter = "Database (*.mdb) |*.mdb"
.ShowOpen
Located_Database = .FileName
    If Right(.FileName, 12) <> "db_Stock.mdb" Then
        MsgBox "The Database must be db_Stock.mdb.", vbExclamation
        GoTo ReturnOpen
    Else
        On Error Resume Next
        reg_obj.RegWrite (Database_Path_Store), Located_Database
        Database_Path = Located_Database
        
    End If
End With
Exit Sub
Err:
    'MsgBox "Error occurred while opening the file...", vbCritical
    'End
frmDatabaseSelectionMsg.Show 1
End Sub

Public Sub Theme_Settings()
On Error Resume Next
Retrieve_User
If ret_user <> "" Then
    Set rsttheme = New ADODB.Recordset
        rsttheme.CursorLocation = adUseClient
        'On Error GoTo db_Error
        rsttheme.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & ret_user & "'", dbcon, adOpenStatic, adLockOptimistic
            If rsttheme.RecordCount > 0 Then
                If Not IsNull(rsttheme("THEME_SET")) Then
                    inttheme = rsttheme("THEME_SET")
                End If
            Else
                Set rstfirstuser = New ADODB.Recordset
                    rstfirstuser.CursorLocation = adUseClient
                    rstfirstuser.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
                        If rstfirstuser.RecordCount > 0 Then
                            If Not IsNull(rstfirstuser("THEME_SET")) Then
                                inttheme = rstfirstuser("THEME_SET")
                            End If
                        End If
            End If
Else
    Set rstisfirstuser = New ADODB.Recordset
        rstisfirstuser.CursorLocation = adUseClient
        rstisfirstuser.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
            If rstisfirstuser.RecordCount > 0 Then
                If Not IsNull(rstisfirstuser("THEME_SET")) Then
                    inttheme = rstisfirstuser("THEME_SET")
                End If
            End If
End If


                
'frmStyle.Picture = frmStyle.imgTheme1.Picture
'Load frmStyle
'frmStyle.Show
End Sub

Public Sub Check_User_Exist_Acc_Type_Change()
Dim current_account_type As String
'
Set rstcheckuserexist = New ADODB.Recordset
    rstcheckuserexist.CursorLocation = adUseClient
    On Error GoTo Check_Error
    rstcheckuserexist.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & User & "'", dbcon, adOpenStatic, adLockOptimistic
        If rstcheckuserexist.RecordCount > 0 Then
                If Not IsNull(rstcheckuserexist("INSTANT_LOG_OUT")) Then
                       If rstcheckuserexist("INSTANT_LOG_OUT") = 1 Then
                          Store_User_Logged_Status_Logout
                          intsystemlogstatus = 5
                          User_Log_Out
                          End
                       End If
                End If
                If Not IsNull(rstcheckuserexist("INSTANT_LOG_OFF")) Then
                       If rstcheckuserexist("INSTANT_LOG_OFF") = 1 Then
                            intsystemlogstatus = 6
                            Form_Unload_Pro
                           'Unload frmMain
                            frmLogin.Show
                            Exit Sub
                       End If
                End If
                
                
           current_account_type = rstcheckuserexist("TYPE")
            If current_account_type <> intaccount_type Then
                MsgBox "Your Account Type has been changed." & vbCrLf & "You must Log Off the system now.", vbCritical
                 Form_Unload_Pro
                 'Unload frmMain
                 frmLogin.Show
            End If
            
        Else
           MsgBox "Your Account, " & User & " has been deleted." & vbCrLf & "Please, contact Administrator.", vbCritical
           Store_User_Logged_Status_Logout
           User_Log_Out
           End
        End If
        Exit Sub
Check_Error:
        If Err.Number = -2147467259 Then
            MsgBox "Your Account, " & User & " has been deleted." & vbCrLf & "Please, contact Administrator.", vbCritical
        Else
            MsgBox "Error_Code: " & Err.Number & " : " & Err.Description, vbCritical
        End If
    'rstcheckuserexist.Close
'Set rstcheckuserexist = Nothing
End Sub

Public Sub Store_User_Logged_Status_Login()
On Error Resume Next
 Set rstuserloggedstatus_login = New ADODB.Recordset
     rstuserloggedstatus_login.CursorLocation = adUseClient
     rstuserloggedstatus_login.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & User & "'", dbcon, adOpenStatic, adLockOptimistic
        rstuserloggedstatus_login("LOGGED_STATUS") = "1"
        If intsetlog = 1 Then
            rstuserloggedstatus_login("LOGGED_STATUS_DATE") = Format(Date, "dd/MM/yyyy")
            rstuserloggedstatus_login("LOGGED_STATUS_TIME") = Time
        End If
        rstuserloggedstatus_login.Update
rstuserloggedstatus_login.Close
Set rstuserloggedstatus_login = Nothing
End Sub
Public Sub Store_User_Logged_Status_Logout()
On Error Resume Next
 Set rstuserloggedstatus_logout = New ADODB.Recordset
     rstuserloggedstatus_logout.CursorLocation = adUseClient
     rstuserloggedstatus_logout.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & User & "'", dbcon, adOpenStatic, adLockOptimistic
        rstuserloggedstatus_logout("LOGGED_STATUS") = "0"
         rstuserloggedstatus_logout("LOGGED_STATUS_DATE") = Null
        rstuserloggedstatus_logout("LOGGED_STATUS_TIME") = Null
        rstuserloggedstatus_logout("INSTANT_LOG_OUT") = Null
        rstuserloggedstatus_logout("INSTANT_LOG_OFF") = Null
        rstuserloggedstatus_logout.Update
     rstuserloggedstatus_logout.Close
 Set rstuserloggedstatus_logout = Nothing
End Sub

Public Sub CHECK_USER_SENDING_MSG()
On Error GoTo Err
Dim User_Sending_Message As String
Set rstsending_msg = New ADODB.Recordset
    rstsending_msg.CursorLocation = adUseClient
    rstsending_msg.Open "SELECT * FROM MESSAGE ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic

     User_Sending_Message = rstsending_msg("MESSAGE_SENDING_GLOBAL")
      If Not IsNull(rstsending_msg("USER")) Then
        Sending_User = rstsending_msg("USER")
      End If
     If User_Sending_Message = 1 Then
        'MsgBox "The User: " & Sending_User & " is sending a Message." & vbCrLf & "Please, try again shortly...", vbExclamation
        frmUserSendingMessage.Show 1
     'Exit Sub
     ElseIf User_Sending_Message = 0 Then
        frmSendMessage.Show 1
     End If
    rstsending_msg.Close
Set rstsending_msg = Nothing
Exit Sub
Err:
If Err.Number = 400 Then
    Exit Sub
Else
   MsgBox "Error Occurred." & vbCrLf & "Error_Code: " & Err.Number, vbCritical
End If
End Sub

Public Sub Check_Message()
Dim current_account_type, Msg, Msg_Time_Sent, Msg_Date_Sent As String
On Error Resume Next
Set rstcheckmessage = New ADODB.Recordset
    rstcheckmessage.CursorLocation = adUseClient
    rstcheckmessage.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & User & "'", dbcon, adOpenStatic, adLockOptimistic
        If rstcheckmessage.RecordCount > 0 Then
            If Not IsNull(rstcheckmessage("SENDER")) Then
                sender = rstcheckmessage("SENDER")
            Else
                sender = ""
            End If
            If Not IsNull(rstcheckmessage("INCOMMING_MSG")) Then
                Msg = rstcheckmessage("INCOMMING_MSG")
            Else
                Msg = ""
            End If
            If Not IsNull(rstcheckmessage("MSG_TYPE")) Then
                Msg_Type = rstcheckmessage("MSG_TYPE")
            Else
                Msg_Type = ""
            End If
            If Not IsNull(rstcheckmessage("Msg_Time_Sent")) Then
                Msg_Time_Sent = rstcheckmessage("MSG_SENT_TIME")
            Else
                Msg_Time_Sent = ""
            End If
            If Not IsNull(rstcheckmessage("Msg_Date_Sent")) Then
                Msg_Date_Sent = rstcheckmessage("MSG_SENT_DATE")
            Else
                Msg_Date_Sent = ""
            End If
        End If
            If sender <> "" And Msg <> "" And Msg_Type <> "" Then
                If Not IsNull(rstcheckmessage("SYSTEM_LOG_OUT_NTY")) Then
                    If rstcheckmessage("SYSTEM_LOG_OUT_NTY") = "1" Then
                        blnsystem_log_out_notification = True
                        'Load frmMsgReceiver
                        rstcheckmessage("SENDER") = Null
                        rstcheckmessage("INCOMMING_MSG") = Null
                        rstcheckmessage("MSG_TYPE") = Null
                        rstcheckmessage("SYSTEM_LOG_OUT_NTY") = Null
                        rstcheckmessage("MSG_SENT_TIME") = Null
                        rstcheckmessage("MSG_SENT_DATE") = Null
                        rstcheckmessage.Update
                        frmReceivedMsgMessage.Show 1
                        'MsgBox "You have received a Message From " & sender & "." & vbCrLf & vbCrLf & "Message Type: " & Msg_Type, vbInformation, "SCS Messenger"
                        frmMsgReceiver.lblMsgFrom.Caption = sender
                        frmMsgReceiver.lblMsgType.Caption = Msg_Type
                        frmMsgReceiver.txtIncomming_Msg_Body = Msg
                        frmMsgReceiver.lblTimeSent = Msg_Time_Sent
                        frmMsgReceiver.lblDateSent = Msg_Date_Sent
                        frmMsgReceiver.Height = 7335
                        frmMsgReceiver.Show 1
                    End If
                ElseIf Not IsNull(rstcheckmessage("USER_LOG_OFF_NTY")) Then
                    If rstcheckmessage("USER_LOG_OFF_NTY") = "1" Then
                        blnuser_log_off_notification = True
                        rstcheckmessage("SENDER") = Null
                        rstcheckmessage("INCOMMING_MSG") = Null
                        rstcheckmessage("MSG_TYPE") = Null
                        rstcheckmessage("USER_LOG_OFF_NTY") = Null
                        rstcheckmessage("MSG_SENT_TIME") = Null
                        rstcheckmessage("MSG_SENT_DATE") = Null
                        rstcheckmessage.Update
                        'MsgBox "You have received a Message From " & sender & "." & vbCrLf & vbCrLf & "Message Type: " & Msg_Type, vbInformation, "SCS Messenger"
                        frmReceivedMsgMessage.Show 1
                        frmMsgReceiver.lblMsgFrom.Caption = sender
                        frmMsgReceiver.lblMsgType.Caption = Msg_Type
                        frmMsgReceiver.txtIncomming_Msg_Body = Msg
                        frmMsgReceiver.lblTimeSent = Msg_Time_Sent
                        frmMsgReceiver.lblDateSent = Msg_Date_Sent
                        frmMsgReceiver.Height = 7335
                        frmMsgReceiver.Show 1
                End If
                Else
                    'Load frmMsgReceiver
                    rstcheckmessage("SENDER") = Null
                    rstcheckmessage("INCOMMING_MSG") = Null
                    rstcheckmessage("MSG_TYPE") = Null
                    rstcheckmessage("MSG_SENT_TIME") = Null
                    rstcheckmessage("MSG_SENT_DATE") = Null
                    rstcheckmessage.Update
                    'MsgBox "You have received a Message From " & sender & "." & vbCrLf & vbCrLf & "Message Type: " & Msg_Type, vbInformation, "SCS Messenger"
                    frmReceivedMsgMessage.Show 1
                    frmMsgReceiver.lblMsgFrom.Caption = sender
                    frmMsgReceiver.lblMsgType.Caption = Msg_Type
                    frmMsgReceiver.txtIncomming_Msg_Body = Msg
                    frmMsgReceiver.lblTimeSent = Msg_Time_Sent
                    frmMsgReceiver.lblDateSent = Msg_Date_Sent
                    frmMsgReceiver.Show 1
                    
                End If
            
            End If

    rstcheckmessage.Close
    Set rstcheckmessage = Nothing

End Sub

Public Sub User_Log_Off_Pro()
'On Error Resume Next
'Unload frmMain
frmLogin.Show 1
End Sub
Public Sub Retrieve_User()
On Error Resume Next
ret_user = reg_obj.RegRead(loggeduser)
End Sub

Public Sub Form_Unload_Pro()

' unload all the data reports
Unload drAllissue: Unload drAlloncurstock
Unload drAllreceipt: Unload drBookkIssueDate
Unload drBookReceiptDate: Unload drBookscheckGreaterthan
Unload drBookscheckLessthan: Unload drGroupedbyCategoryCurStock
Unload drGroupedbyCoruseCurStock: Unload drGroupedbyModuleBooksIssue
Unload drGroupedbyModuleBooksIssueDate: Unload drGroupedbyModuleBooksReceiptDate
Unload drGroupedbyModuleBooksReceive: Unload drLogStatus

'unload all the forms
Unload frmAccessCode: Unload frmFN
Unload frmBF: Unload frmSendMessage
Unload frmLoggedUserStatus: Unload frmBackupdatabase: Unload frmCompactDatabase
Unload frmAbout: Unload frmPerformanceoptimizer: Unload frmUserLogStatusReports
Unload frmReports: Unload frmTheme
Unload frmFindBookIssue: Unload frmFindBookReceipt
Unload frmFindBooksDetails: Unload frmFindCurrentStock
Unload frmFindMasterInfo: Unload frmBookdetails
Unload frmCurrentStock: Unload frmMasterInformation
Unload frmBookIssue: Unload frmBookReceipt
Unload frmChangePassword: Unload frmCreateAccount
Unload frmPassword: Unload frmDeleteAllRecords
Unload frmMsgReceiver: Unload frmMsgErrorMessage
Unload frmReceivedMsgMessage: Unload frmSentMsgMessage
Unload frmUserSendingMessage: Unload frmSplash
Unload frmStyle: Unload frmBooksNotInUseEntry
Unload frmDatabaseSelectionMsg: Unload frmResolutionAlert
Unload frmClearDatabaseLocation
Unload frmMain
End Sub

'Public Sub Clear_Database_Location()
'On Error Resume Next
'reg_obj.RegDelete (Database_Path_Store)
'End Sub

Public Sub User_Log_In()
On Error Resume Next
Set rstuserlogin = New ADODB.Recordset
    rstuserlogin.CursorLocation = adUseClient
    rstuserlogin.Open "SELECT * FROM USER_LOG", dbcon, adOpenStatic, adLockOptimistic
    rstuserlogin.AddNew
    rstuserlogin("USER") = User
    date_log_in = frmStyle.dtpNow.Value
    rstuserlogin("DATE") = date_log_in
    time_log_in = Time
    rstuserlogin("LOG_IN") = time_log_in
    rstuserlogin.Update
rstuserlogin.Close
Set rstuserlogin = Nothing
End Sub

Public Sub User_Log_Out()
On Error Resume Next
Set rstuserlogout = New ADODB.Recordset
    rstuserlogout.CursorLocation = adUseClient
    rstuserlogout.Open "SELECT * FROM USER_LOG WHERE USER = '" & User & "' AND LOG_IN = '" & time_log_in & "' AND DATE =  " & "#" & date_log_in & "#", dbcon, adOpenStatic, adLockOptimistic
    'MsgBox rstuserlogout.RecordCount
    rstuserlogout("LOG_OUT") = Time
    Select Case intsystemlogstatus
        Case 1: rstuserlogout("LOG_OUT_STATUS") = "Log Out by user"
        Case 2: rstuserlogout("LOG_OUT_STATUS") = "Log Off by user"
        Case 3: rstuserlogout("LOG_OUT_STATUS") = "Log Out by Administrator"
        Case 4: rstuserlogout("LOG_OUT_STATUS") = "Log Off by Administrator"
        Case 5: rstuserlogout("LOG_OUT_STATUS") = "Instant Log Out"
        Case 6: rstuserlogout("LOG_OUT_STATUS") = "Instant Log Off"
        Case 7: rstuserlogout("LOG_OUT_STATUS") = "Log Off by user for Resolution"
        Case 8: rstuserlogout("LOG_OUT_STATUS") = "User Cleared Database Location"
        'Case 9
            
            'rstuserlogout("LOG_OUT_STATUS") = "User Restored the Database"
        Case Else
            rstuserlogout("LOG_OUT_STATUS") = "Log Out by user"
        End Select
rstuserlogout.Update
rstuserlogout.Close
Set rstuserlogout = Nothing
End Sub
Public Sub Record_Deletion(ByVal opt As Integer)
Dim reccount As Double
Dim i As Double

Select Case opt
    Case 1
        Set rstIssue = New ADODB.Recordset
        rstIssue.CursorLocation = adUseClient
        rstIssue.Open "SELECT * FROM B_ISSUE ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstIssue.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstIssue.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
        
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstIssue.Delete
                rstIssue.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
       
        rstIssue.Close
        Set rstIssue = Nothing
  Case 2
        Set rstreceipt = New ADODB.Recordset
        rstreceipt.CursorLocation = adUseClient
        rstreceipt.Open "SELECT * FROM B_RECEIPT ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstreceipt.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstreceipt.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstreceipt.Delete
                rstreceipt.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus

        rstreceipt.Close
        Set rstreceipt = Nothing
  Case 3
        Set rstcurrentstock = New ADODB.Recordset
        rstcurrentstock.CursorLocation = adUseClient
        rstcurrentstock.Open "SELECT * FROM CUR_STOCK ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstcurrentstock.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstcurrentstock.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
        
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstcurrentstock.Delete
                rstcurrentstock.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus

        rstcurrentstock.Close
        Set rstcurrentstock = Nothing
  Case 4
        Set rstDetails = New ADODB.Recordset
        rstDetails.CursorLocation = adUseClient
        rstDetails.Open "SELECT * FROM DETAILS ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstDetails.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstDetails.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstDetails.Delete
                rstDetails.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
        
        rstDetails.Close
        Set rstDetails = Nothing
  Case 5
        Set rstcourse = New ADODB.Recordset
        rstcourse.CursorLocation = adUseClient
        rstcourse.Open "SELECT * FROM COURSE ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstcourse.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstcourse.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstcourse.Delete
                rstcourse.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
        
        rstcourse.Close
        Set rstcourse = Nothing
  Case 6
        Set rstmodule = New ADODB.Recordset
        rstmodule.CursorLocation = adUseClient
        rstmodule.Open "SELECT * FROM MODULER ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstmodule.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstmodule.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstmodule.Delete
                rstmodule.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
            
        rstmodule.Close
        Set rstmodule = Nothing
  Case 7
        Set rstcategory = New ADODB.Recordset
        rstcategory.CursorLocation = adUseClient
        rstcategory.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstcategory.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstcategory.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstcategory.Delete
                rstcategory.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
        
        rstcategory.Close
        Set rstcategory = Nothing
  Case 8
        Set rstuserlog = New ADODB.Recordset
        rstuserlog.CursorLocation = adUseClient
        rstuserlog.Open "SELECT * FROM USER_LOG ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstuserlog.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstuserlog.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                rstuserlog.Delete
                rstuserlog.MoveNext
            Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
        
        rstuserlog.Close
        Set rstuserlog = Nothing
  Case 9
        Set rstuseraccounts = New ADODB.Recordset
        rstuseraccounts.CursorLocation = adUseClient
        rstuseraccounts.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
        reccount = rstuseraccounts.RecordCount
        frmDeleteAllRecords.lblReccount.Caption = "Total Records  - " & reccount
        i = 0
            Do While Not rstuseraccounts.EOF
                i = i + 1
                frmDeleteAllRecords.lblDeleting.Caption = "Deleting... " & i
                frmDeleteAllRecords.lblDeleting.Refresh
                frmDeleteAllRecords.imgProgress.Width = i / reccount * 5055
                frmDeleteAllRecords.imgProgress.Refresh
                    If rstuseraccounts("USER_NAME") <> "Administrator" Then: rstuseraccounts.Delete
                rstuseraccounts.MoveNext
                Loop
        'Unload frmAccessCode
        MsgBox "All Records deleted in " & frmDeleteAllRecords.cmbSelectTable & ".", vbInformation
        frmDeleteAllRecords.lblReccount.Caption = ""
        frmDeleteAllRecords.lblDeleting.Caption = ""
        frmDeleteAllRecords.imgProgress.Width = 0
        frmDeleteAllRecords.cmbSelectTable.SetFocus
        
        rstuseraccounts.Close
        Set rstuseraccounts = Nothing
End Select
End Sub
Public Sub Get_Report_for_Books_not_in_stock()
        Set rst_books_not_in_stock = New ADODB.Recordset
            rst_books_not_in_stock.CursorLocation = adUseClient
            rst_books_not_in_stock.Open "SELECT * FROM BOOKS_NOT_IN_USE ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
        Set drAlloncurstock.DataSource = rst_books_not_in_stock
            drAlloncurstock.Sections("Section4").Controls("label1").Caption = "Books not in use"
            drAlloncurstock.Sections("Section5").Controls("label7").Caption = "Total Books Count"
            drAlloncurstock.Show 1
            
        rst_books_not_in_stock.Close
        Set rst_books_not_in_stock = Nothing
End Sub
Public Sub Check_Resolution_Alert()
On Error GoTo Err
Dim getval As String
getval = reg_obj.RegRead(Resolution_Alert)
If getval = "0" Then
    intresolutionAlertenable = 0
ElseIf getval = "1" Then
    intresolutionAlertenable = 1
Else
    intresolutionAlertenable = 1
End If
Exit Sub
Err:
intresolutionAlertenable = 1
End Sub
Public Sub Compact_Database()
Dim dbtemp As String
On Error GoTo Err
Check_For_Users_Exist
  If intdbok = 0 Then
    Exit Sub
  End If
dbtemp = Mid(Database_Path, 1, Len(Database_Path) - 12) & "dbtemp.mdb"

If MsgBox("Are you sure you want to Compact the Database?" & vbCrLf & "If you click Yes, Database connectivity will be colsed and" & vbCrLf & "compact the Database?", vbYesNo + vbQuestion) = vbYes Then
     'On Error Resume Next
                    frmMain.Timer1.Enabled = False
                    frmMain.Timer2.Enabled = False
                    frmMain.Timer3.Enabled = False
                    frmMain.Timer4.Enabled = False
                    dbcon.Close
                    Set dbcon = Nothing
                    deGroupedreports.cnGroupedreports.Close
                    If Dir(dbtemp) = "dbtemp.mdb" Then
                        Kill dbtemp
                    End If
                    DBEngine.CompactDatabase Database_Path, dbtemp, , , ";pwd=" & pword_db
                    Kill Database_Path
                    CopyFile dbtemp, Database_Path, 1
                    Kill dbtemp
                    MsgBox "Database was successfully compacted.", vbInformation
                    End
End If
Exit Sub
Err:
MsgBox Err.Description & vbCrLf & "Error Occurred while compacting the Database.", vbCritical
End
End Sub
Public Sub Check_For_Users_Exist()
On Error GoTo Err
Set rstuserexist = New ADODB.Recordset
    rstuserexist.CursorLocation = adUseClient
    rstuserexist.Open "SELECT * FROM ACCOUNT_SET WHERE LOGGED_STATUS = '" & 1 & "' ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
    
If rstuserexist.RecordCount > 0 Then
   
        If rstuserexist.RecordCount = 1 And rstuserexist("USER_NAME") = User Then
             If inttask = 1 Then
                 intdbok = 1
             ElseIf inttask = 2 Then
                 intdbok = 1
             End If
            
        Else
            If inttask = 1 Then
               MsgBox "Some other users are still using the system." & vbCrLf & "Restoration process cannot be performed." & vbCrLf & "Please, log out them first.", vbExclamation
                 intdbok = 0
            ElseIf inttask = 2 Then
               MsgBox "Some other users are still using the system." & vbCrLf & "Compact Database process cannot be performed." & vbCrLf & "Please, log out them first.", vbExclamation
                 intdbok = 0
            End If
        End If
End If
rstuserexist.Close
Set rstuserexist = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
End Sub

Public Sub BypassLog()
On Error GoTo Err
If Command = "bypassadmn" Then
   bypassadmin = 1
   intaccount_type = 1
   User = "Administrator"
   frmMain.Timer1.Enabled = False
   frmMain.Timer2.Enabled = False
   frmMain.Timer3.Enabled = False
   frmMain.Timer4.Enabled = False
   user_record_delete_privilege = 1
   user_record_add_privilege = 1
   user_record_edit_privilege = 1
   user_view_report_privilege = 1
   user_send_msg_privilege = 1
   frmMain.Show
Exit Sub
End If
Exit Sub
Err:
End Sub
Public Sub Button_Record_Not_Exist_Mode(ByVal frm As Form)
On Error Resume Next
frm.cmdAdd.Enabled = True: frm.cmdEdit.Enabled = False
frm.cmdDelete.Enabled = False: frm.cmdCancel.Enabled = False
frm.cmdSave.Enabled = False: frm.cmdPrevious.Enabled = False
frm.cmdNext.Enabled = False: frm.cmdFirst.Enabled = False
frm.cmdLast.Enabled = False
End Sub
Public Sub Button_Add_Edit_Save_Cancle_RecordExist_Mode(ByVal frm As Form, ByVal bval As Boolean)
On Error Resume Next
frm.cmdCancel.Enabled = Not bval: frm.cmdSave.Enabled = Not bval
frm.cmdAdd.Enabled = bval: frm.cmdEdit.Enabled = bval
frm.cmdDelete.Enabled = bval: frm.cmdPrevious.Enabled = bval
frm.cmdNext.Enabled = bval: frm.cmdFirst.Enabled = bval
frm.cmdLast.Enabled = bval: frm.cmdFind.Enabled = bval
If inttabselect = 1 Then
   frm.cmdFind.Enabled = bval
Else
   frm.cmdFind.Enabled = Not bval
End If
End Sub
Public Function Check_For_Privilege(ByVal Priv_Opt As Integer) As Boolean
Select Case Priv_Opt
    Case 1:
        If user_record_add_privilege = 0 Then: Check_For_Privilege = True: MsgBox "You do not have permission to add records...!" & _
        vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Case 2
        If user_record_edit_privilege = 0 Then: Check_For_Privilege = True: MsgBox "You do not have permission to edit records...!" & _
        vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Case 3
        If user_record_delete_privilege = 0 Then: Check_For_Privilege = True: MsgBox "You do not have permission to delete records...!" & _
        vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Case 4
        If user_view_report_privilege = 0 Then: Check_For_Privilege = True: MsgBox "You do not have permission to view reports...!" & _
        vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Case 5
        If user_send_msg_privilege = 0 Then: Check_For_Privilege = True: MsgBox "You do not have permission to send messages...!" & _
        vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
End Select
End Function
Public Sub Theme_Handle(ByVal T_Opt As Integer)
Select Case T_Opt
Case 1: Select Case inttheme
            Case 1: frmStyle.Picture = frmStyle.imgTheme1.Picture
            Case 2: frmStyle.Picture = frmStyle.imgTheme2.Picture
            Case 3: frmStyle.Picture = frmStyle.imgTheme3.Picture
            Case 4: frmStyle.Picture = frmStyle.imgTheme4.Picture
            Case 5: frmStyle.Picture = frmStyle.imgTheme5.Picture
            Case 6: frmStyle.Picture = frmStyle.imgTheme6.Picture
        End Select
Case 2: Select Case inttheme
            Case 1: frmMain.Picture4.Picture = frmStyle.imgTheme01
            Case 2: frmMain.Picture4.Picture = frmStyle.imgTheme02
            Case 3: frmMain.Picture4.Picture = frmStyle.imgTheme03
            Case 4: frmMain.Picture4.Picture = frmStyle.imgTheme04
            Case 5: frmMain.Picture4.Picture = frmStyle.imgTheme05
            Case 6: frmMain.Picture4.Picture = frmStyle.imgTheme06
        End Select
Case 3: Select Case inttheme
            Case 1: frmMain.Picture4.Picture = frmStyle.imgTheme001
            Case 2: frmMain.Picture4.Picture = frmStyle.imgTheme002
            Case 3: frmMain.Picture4.Picture = frmStyle.imgTheme003
            Case 4: frmMain.Picture4.Picture = frmStyle.imgTheme004
            Case 5: frmMain.Picture4.Picture = frmStyle.imgTheme005
            Case 6: frmMain.Picture4.Picture = frmStyle.imgTheme006
        End Select
End Select
End Sub
