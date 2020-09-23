VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCurrentStock 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current Stock "
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   Icon            =   "frmCurrentStock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   11055
      Begin VB.Image Image2 
         Height          =   360
         Left            =   10560
         Picture         =   "frmCurrentStock.frx":000C
         ToolTipText     =   "Application Help"
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain Current Stock Here. These Entries Are Highly Effect For Other Entries Of The System."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   240
         Width           =   8775
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   120
         Picture         =   "frmCurrentStock.frx":0776
         Top             =   130
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   5520
      TabIndex        =   25
      Top             =   610
      Width           =   5670
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4520
         Left            =   40
         ScaleHeight     =   4515
         ScaleWidth      =   5595
         TabIndex        =   26
         Top             =   120
         Width           =   5600
         Begin VB.CommandButton cmdReturntofrecords 
            Caption         =   "Return to F&ull Records"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   940
            MouseIcon       =   "frmCurrentStock.frx":1244
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   4080
            Width           =   3735
         End
         Begin MSFlexGridLib.MSFlexGrid msfgCurStock 
            Height          =   3975
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   7011
            _Version        =   393216
            BackColor       =   16642796
            AllowUserResizing=   1
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   610
      Width           =   5300
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4515
         Left            =   40
         ScaleHeight     =   4515
         ScaleWidth      =   5205
         TabIndex        =   15
         Top             =   120
         Width           =   5200
         Begin VB.TextBox txtQuantity 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   2280
            Width           =   1575
         End
         Begin VB.ComboBox cmbCategory 
            BackColor       =   &H80000007&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   315
            Left            =   1800
            TabIndex        =   24
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox txtModuleDes 
            BackColor       =   &H00000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   600
            Width           =   3255
         End
         Begin VB.ComboBox cmbCourse 
            BackColor       =   &H80000007&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000004&
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   1320
            Width           =   2415
         End
         Begin VB.CommandButton cmdExit 
            Cancel          =   -1  'True
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            MouseIcon       =   "frmCurrentStock.frx":1396
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "&Last >|"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            MouseIcon       =   "frmCurrentStock.frx":14E8
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "|< &First"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MouseIcon       =   "frmCurrentStock.frx":163A
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "< &Previous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MouseIcon       =   "frmCurrentStock.frx":178C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next >"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            MouseIcon       =   "frmCurrentStock.frx":18DE
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            MouseIcon       =   "frmCurrentStock.frx":1A30
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            MouseIcon       =   "frmCurrentStock.frx":1B82
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            MouseIcon       =   "frmCurrentStock.frx":1CD4
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   3120
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            MouseIcon       =   "frmCurrentStock.frx":1E26
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            MouseIcon       =   "frmCurrentStock.frx":1F78
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MouseIcon       =   "frmCurrentStock.frx":20CA
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   3120
            Width           =   1215
         End
         Begin VB.ComboBox cmbModule 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            TabIndex        =   1
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   15
            Left            =   120
            TabIndex        =   21
            Top             =   2880
            Width           =   5055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Course"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Module Code"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Module Description"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmCurrentStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCmforcurstock As ADODB.Recordset
Dim rstMmforcurstock As ADODB.Recordset
Dim rstCgmforcurstock As ADODB.Recordset
Dim rstCurstock As ADODB.Recordset
Dim rstDatatoflex As ADODB.Recordset
Dim blnadd_Click As Boolean
Dim blnedit_Click As Boolean
Dim validity As Boolean
Dim duplicate As Boolean
Dim valid_content As Boolean
Dim find_mode As Boolean
Dim update_find As Boolean
Dim invalid_quantity As Boolean
Private Sub cmbModule_Click()
On Error GoTo Ret_Error
Set rstInfofromdetails = New ADODB.Recordset
    rstInfofromdetails.CursorLocation = adUseClient
    rstInfofromdetails.Open "SELECT * FROM DETAILS WHERE MODULER = '" & cmbModule.Text & " '", dbcon, adOpenStatic, adLockReadOnly
    txtModuleDes.Text = rstInfofromdetails("MODULE_DES")
    cmbCourse.Text = rstInfofromdetails("COURSE")
    cmbCategory.Text = rstInfofromdetails("CATEGORY")
    
    rstInfofromdetails.Close
    Set rstInfofromdetails = Nothing
    
    Exit Sub
Ret_Error:
    MsgBox "Please update the Books Details First.", vbExclamation
    
    rstInfofromdetails.Close
    Set rstInfofromdetails = Nothing
    
    cmdCancel_Click
    frmBookdetails.Show 1

End Sub

Private Sub cmbModule_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQuantity.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
If Check_For_Privilege(1) = True Then: Exit Sub
Enable_Controls
blnadd_Click = True
blnedit_Click = False
Clear_Fields
cmbModule.SetFocus
SendKeys "{F4}"
SendKeys "{DOWN}"
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
rstCurstock.CancelUpdate
Clear_Fields
Inforfield
Disable_Controls
On Error Resume Next
cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err
If Check_For_Privilege(3) = True Then: Exit Sub
If MsgBox("Are you sure you want to delete the current Record ?", vbQuestion + vbYesNo) = vbYes Then
        On Error Resume Next
        rstCurstock.Delete
        rstCurstock.MoveNext
            If rstCurstock.EOF Then
                rstCurstock.MoveLast
            End If
           If blnfind_status = True Then
                update_find = True
                Find_for_Details
           Else
                Clear_Fields
                Inforfield
                Format_Flex
                update_Flex
           End If
End If
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub

Private Sub cmdEdit_Click()
If Check_For_Privilege(2) = True Then: Exit Sub
Enable_Controls
blnedit_Click = True
blnadd_Click = False
cmbModule.SetFocus
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
frmFindCurrentStock.Show 1
End Sub

Private Sub cmdFirst_Click()
    If rstCurstock.BOF = False Then
        rstCurstock.MoveFirst
        Inforfield
        MsgBox "You are on the First Record.", vbInformation
    End If

End Sub

Private Sub cmdLast_Click()
    If rstCurstock.EOF = False Then
        rstCurstock.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If

End Sub

Private Sub cmdNext_Click()
    If rstCurstock.EOF = False Then
        rstCurstock.MoveNext
        Inforfield
    End If
    If rstCurstock.EOF Then
        rstCurstock.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If
End Sub

Private Sub cmdPrevious_Click()
    If rstCurstock.BOF = False Then
        rstCurstock.MovePrevious
        Inforfield
    End If
    If rstCurstock.BOF Then
        rstCurstock.MoveFirst
        Inforfield
        MsgBox "You are on the first Record.", vbInformation
    End If
End Sub


Private Sub cmdReturntofrecords_Click()
Load_Initiate
End Sub

Private Sub cmdSave_Click()
Data_Validity
    If validity = True Then
        If blnadd_Click = True Then
            blnadd_Click = False
            rstCurstock.AddNew
            Save_Data
        End If
        If blnedit_Click = True Then
            blnedit_Click = False
            Save_Data
        End If
        
    End If
 
       
End Sub
Public Sub Data_Validity()
'Valid_contents_check
Module_contents_check

      If valid_content = False Then
            validity = False
      Exit Sub
      ElseIf valid_content = True Then
        validity = True
      End If
          
Duplicate_Data

      If duplicate = True Then
          MsgBox "Module is already exist!!", vbExclamation
          cmbModule.SetFocus
          SendKeys "{F4}"
          validity = False
          Exit Sub
      ElseIf duplicate = False Then
          validity = True
      End If
      
 'quantity_validate
        'If invalid_quantity = False Then
            'MsgBox "Invalid Quantity...!", vbExclamation
            'txtQuantity.SetFocus
           ' SendKeys "{HOME}+{END}"
           ' validity = False
          'Exit Sub
       ' ElseIf invalid_quantity = True Then
           ' validity = True
        'End If
End Sub
 Sub Duplicate_Data()
            Set rstValidatemodule = New ADODB.Recordset
            rstValidatemodule.CursorLocation = adUseClient
            rstValidatemodule.Open "SELECT * FROM CUR_STOCK WHERE MODULER = '" & cmbModule.Text & "'", dbcon, adOpenStatic, adLockReadOnly
                
                If rstValidatemodule.RecordCount > 0 Then
                    If blnadd_Click = True Then
                        'blnadd_Click = False
                        duplicate = True
                       
                    ElseIf blnedit_Click = True Then
                        If UCase(rstCurstock("MODULER")) = UCase(cmbModule.Text) Then
                           'blnedit_Click = False
                           duplicate = False
                        Else
                        duplicate = True
                        End If
                    End If
                Else
                duplicate = False
                End If
             rstValidatemodule.Close
             Set rstValidatemodule = Nothing
                
End Sub

'Sub Valid_contents_check()
'Module_contents_check
'End Sub

Sub Module_contents_check()
Dim i As Integer
    For i = 0 To cmbModule.ListCount - 1
        If UCase(Trim(cmbModule.List(i))) = UCase(Trim(cmbModule.Text)) Then
            valid_content = True
            cmbModule.Text = UCase(Trim(cmbModule.Text))
            Exit For
        Else
            valid_content = False
        End If
    Next i
    
'Get the result from the module list and proceed.

    If valid_content = True Then
        Exit Sub
    ElseIf valid_content = False Then
        MsgBox "Select the Module from the List.", vbExclamation
        cmbModule.SetFocus
        SendKeys "{F4}"
    End If
End Sub

'Sub Course_contents_check()
'Dim i As Integer
    'For i = 0 To cmbCourse.ListCount - 1
       ' If UCase(Trim(cmbCourse.List(i))) = UCase(Trim(cmbCourse.Text)) Then
           ' valid_content = True
            'cmbCourse.Text = UCase(Trim(cmbCourse.Text))
            'Exit For
        'Else
         'valid_content = False
        'End If
    'Next i
    
'Get the result from the course list and proceed.

    'If valid_content = True Then
        'Category_contents_check
    'ElseIf valid_content = False Then
        'MsgBox "Select the Course from the List.", vbExclamation
        'cmbCourse.SetFocus
        'SendKeys "{F4}"
    'End If
'End Sub


'Sub Category_contents_check()
'Dim i As Integer
'For i = 0 To cmbCategory.ListCount - 1
    'If UCase(Trim(cmbCategory.List(i))) = UCase(Trim(cmbCategory.Text)) Then
       'valid_content = True
        'cmbCategory.Text = UCase(Trim(cmbCategory.Text))
     'Exit For
    'Else
        'valid_content = False
   ' End If
'Next i

'Get the result from the module list and proceed.

'If valid_content = True Then
    'Exit Sub
'ElseIf valid_content = False Then
        'MsgBox "Select the Category from the List.", vbExclamation
        'cmbCategory.SetFocus
        'SendKeys "{F4}"
'End If
'End Sub

'Public Sub quantity_validate()
'If Val(txtQuantity) = 0 Then
    'invalid_quantity = False
'Else
    'invalid_quantity = True
'End If
'End Sub


Sub Save_Data()
On Error GoTo Err
     rstCurstock("MODULER") = UCase(cmbModule.Text)
     rstCurstock("MODULE_DES") = txtModuleDes.Text
     rstCurstock("COURSE") = UCase(cmbCourse.Text)
     rstCurstock("CATEGORY") = UCase(cmbCategory.Text)
     rstCurstock("QUANTITY") = Val(txtQuantity.Text)
     rstCurstock.Update
     Me.Caption = "Current Stock   -  Record Count: " & rstCurstock.RecordCount
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    If blnfind_status = True Then
        update_find = True
        Find_for_Details
    Else
    
    Format_Flex
    update_Flex
    
    End If
    Disable_Controls
    On Error Resume Next
    cmdAdd.SetFocus
 Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub


Private Sub Form_Load()
'Set rstcmforcurstock = New ADODB.Recordset
'    rstcmforcurstock.CursorLocation = adUseClient
'    rstcmforcurstock.Open "SELECT * FROM COURSE ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
   'MsgBox rstcoursemaster.RecordCount & "Course"
Set rstMmforcurstock = New ADODB.Recordset
    rstMmforcurstock.CursorLocation = adUseClient
    rstMmforcurstock.Open "SELECT * FROM MODULER ORDER BY MODULER", dbcon, adOpenStatic, adLockReadOnly
   'MsgBox rstcoursemaster.RecordCount & " MODULER"
'Set rstcgmforcurstock = New ADODB.Recordset
'    rstcgmforcurstock.CursorLocation = adUseClient
'    rstcgmforcurstock.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly
   'MsgBox rstcategorymaster.RecordCount & "Category"
Set rstCurstock = New ADODB.Recordset
    rstCurstock.CursorLocation = adUseClient
    rstCurstock.Open "SELECT * FROM CUR_STOCK ORDER BY COURSE", dbcon, adOpenStatic, adLockOptimistic
Set rstDatatoflex = New ADODB.Recordset
    rstDatatoflex.CursorLocation = adUseClient
    rstDatatoflex.Open "SELECT * FROM CUR_STOCK ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
    
    Add_Maser_Data_to_Combo
    Inforfield
    Format_Flex
    Add_Flex
    Disable_Controls
    Me.Picture = frmStyle.Picture
    End Sub

Public Sub Add_Maser_Data_to_Combo()
If rstMmforcurstock.RecordCount > 0 Then
    cmbModule.Clear
    Do While Not rstMmforcurstock.EOF
        cmbModule.AddItem rstMmforcurstock("MODULER")
    rstMmforcurstock.MoveNext
    Loop
Else
 MsgBox "Please update the master information for Module.", vbExclamation
Exit Sub
End If

'If rstcmforcurstock.RecordCount = 0 Then
    'cmbCourse.Clear
    'Do While Not rstcmforcurstock.EOF
    '    cmbCourse.AddItem rstcmforcurstock("COURSE")
    'rstcmforcurstock.MoveNext
    'Loop
'Else
' MsgBox "Please update the master information", vbExclamation
'    Exit Sub
'End If

'If rstcgmforcurstock.RecordCount = 0 Then
    'cmbCategory.Clear
    'Do While Not rstcgmforcurstock.EOF
    '    cmbCategory.AddItem rstcgmforcurstock("CATEGORY")
    'rstcgmforcurstock.MoveNext
    'Loop
'Else
' MsgBox "Please update the master information", vbExclamation
' Exit Sub
'End If
End Sub
Public Sub Inforfield()
On Error Resume Next
If rstCurstock.RecordCount > 0 Then
    cmbModule.Text = rstCurstock("MODULER")
    txtModuleDes.Text = rstCurstock("MODULE_DES")
    cmbCourse.Text = rstCurstock("COURSE")
    cmbCategory.Text = rstCurstock("CATEGORY")
    txtQuantity.Text = rstCurstock("QUANTITY")
    Me.Caption = "Current Stock   -  Record Count: " & rstCurstock.RecordCount
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
 Else
    Button_Record_Not_Exist_Mode Me
    End If
End Sub


Public Sub Clear_Fields()
On Error Resume Next
    cmbCourse.Text = ""
    cmbModule.Text = ""
    txtModuleDes.Text = ""
    cmbCategory.Text = ""
    txtQuantity.Text = ""
  
End Sub
Public Sub Format_Flex()
msfgCurStock.CellAlignment = flexAlignLeftCenter
With msfgCurStock
.Clear
.Cols = 4
.Rows = 1
.ColWidth(0) = 1500
.TextMatrix(0, 0) = "Module"
.ColWidth(1) = 1500
.TextMatrix(0, 1) = "Course"
.ColWidth(2) = 1450
.TextMatrix(0, 2) = "Category"
.ColWidth(3) = 1000
.TextMatrix(0, 3) = "Quantity"
End With
End Sub

Public Sub Add_Flex()
Dim i As Integer
Dim rcount As Integer

    If rstDatatoflex.RecordCount > 0 Then
        rcount = rstDatatoflex.RecordCount
        msfgCurStock.Rows = rcount + 1
        i = 1
        With msfgCurStock
            Do While Not rstDatatoflex.EOF
            .Row = i
            .Col = 0: .Text = rstDatatoflex("MODULER")
            .Row = i
            .Col = 1: .Text = rstDatatoflex("COURSE")
            .Row = i
            .Col = 2: .Text = rstDatatoflex("CATEGORY")
            .Row = i
            .Col = 3: .Text = Val(rstDatatoflex("QUANTITY"))
       i = i + 1
       rstDatatoflex.MoveNext
       
       Loop
       End With
    Else
       Format_Flex
    End If
End Sub
Public Sub update_Flex()
Dim i As Integer
Dim rcount As Integer

    rstDatatoflex.Close
    Set rstDatatoflex = Nothing
    
Set rstDatatoflex = New ADODB.Recordset
    rstDatatoflex.CursorLocation = adUseClient
    rstDatatoflex.Open "SELECT * FROM CUR_STOCK ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly


    If rstDatatoflex.RecordCount > 0 Then
        rcount = rstDatatoflex.RecordCount
        msfgCurStock.Rows = rcount + 1
        i = 1
        With msfgCurStock
            Do While Not rstDatatoflex.EOF
            .Row = i
            .Col = 0: .Text = rstDatatoflex("MODULER")
            .Row = i
            .Col = 1: .Text = rstDatatoflex("COURSE")
            .Row = i
            .Col = 2: .Text = rstDatatoflex("CATEGORY")
            .Row = i
            .Col = 3: .Text = Val(rstDatatoflex("QUANTITY"))
       i = i + 1
       rstDatatoflex.MoveNext
       
       Loop
       End With
    Else
       Format_Flex
    End If
End Sub
Public Sub Find_for_Details()
    rstCurstock.Close
    Set rstCurstock = Nothing
    rstDatatoflex.Close
    Set rstdatetoflex = Nothing
    'MsgBox Find_Val
If detailsfindparameter = 0 Then
    Set rstCurstock = New ADODB.Recordset
        rstCurstock.Open "SELECT* FROM CUR_STOCK WHERE[MODULER] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
    Set rstDatatoflex = New ADODB.Recordset
        rstDatatoflex.Open "SELECT* FROM CUR_STOCK WHERE[MODULER] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        If rstCurstock.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex
            Unload frmFindCurrentStock
            cmdReturntofrecords.Enabled = True
        Else
            If update_find = True Then
                update_find = False
                Load_Initiate
                Exit Sub
            Else
            MsgBox "No records found for this search...!", vbExclamation
            frmFindCurrentStock.txtFindValue.SetFocus
            SendKeys "{Home}+{End}"
            'Load_Initiate
            Exit Sub
            End If
        End If
        
ElseIf detailsfindparameter = 1 Then

     Set rstCurstock = New ADODB.Recordset
        rstCurstock.Open "SELECT* FROM CUR_STOCK WHERE[COURSE] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
     Set rstDatatoflex = New ADODB.Recordset
        rstDatatoflex.Open "SELECT* FROM CUR_STOCK WHERE[COURSE] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        If rstCurstock.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex
            Unload frmFindCurrentStock
            cmdReturntofrecords.Enabled = True
        Else
            If update_find = True Then
                update_find = False
                Load_Initiate
                Exit Sub
            Else
            MsgBox "No records found for this search...!", vbExclamation
            frmFindCurrentStock.txtFindValue.SetFocus
            SendKeys "{Home}+{End}"
            'Load_Initiate
            Exit Sub
            End If
        End If
       
   ElseIf detailsfindparameter = 2 Then

     Set rstCurstock = New ADODB.Recordset
        rstCurstock.Open "SELECT* FROM CUR_STOCK WHERE[CATEGORY] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
     Set rstDatatoflex = New ADODB.Recordset
        rstDatatoflex.Open "SELECT* FROM CUR_STOCK WHERE[CATEGORY] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        If rstCurstock.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex
            Unload frmFindCurrentStock
            cmdReturntofrecords.Enabled = True
        Else
             If update_find = True Then
                update_find = False
                Load_Initiate
                Exit Sub
            Else
            MsgBox "No records found for this search...!", vbExclamation
            frmFindCurrentStock.txtFindValue.SetFocus
            SendKeys "{Home}+{End}"
            'Load_Initiate
            Exit Sub
            End If
         End If
   End If
End Sub

Public Sub Load_Initiate()
cmbCourse.Clear
cmbModule.Clear
cmbCategory.Clear
Form_Load
cmdReturntofrecords.Enabled = False
blnfind_status = False
update_find = False
End Sub


Public Sub Enable_Controls()
cmbModule.Enabled = True
txtQuantity.Enabled = True
End Sub
Public Sub Disable_Controls()
cmbModule.Enabled = False
txtQuantity.Enabled = False
End Sub

Private Sub txtQuantity_Change()
If Not IsNumeric(txtQuantity) Then: txtQuantity = ""
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: cmdSave.SetFocus
End Sub
