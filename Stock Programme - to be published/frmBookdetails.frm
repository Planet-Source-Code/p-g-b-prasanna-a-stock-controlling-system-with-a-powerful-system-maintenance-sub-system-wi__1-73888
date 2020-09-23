VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBookdetails 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books Details Entry"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "frmBookdetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   9975
      Begin VB.Image Image3 
         Height          =   360
         Left            =   120
         Picture         =   "frmBookdetails.frx":000C
         Top             =   160
         Width           =   360
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Books Details Here. These Entries Are Required For Other Entries Of The System. "
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
         Left            =   600
         TabIndex        =   29
         Top             =   240
         Width           =   8775
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   9480
         Picture         =   "frmBookdetails.frx":0776
         ToolTipText     =   "Application Help"
         Top             =   160
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   5520
      TabIndex        =   25
      Top             =   610
      Width           =   4575
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4575
         Left            =   30
         ScaleHeight     =   4575
         ScaleWidth      =   4485
         TabIndex        =   26
         Top             =   120
         Width           =   4490
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
            Left            =   380
            MouseIcon       =   "frmBookdetails.frx":0EE0
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   4200
            Width           =   3735
         End
         Begin MSFlexGridLib.MSFlexGrid msfgBooksDetails 
            Height          =   4095
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   7223
            _Version        =   393216
            BackColor       =   16642796
            BackColorFixed  =   14737632
            HighLight       =   0
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   610
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4635
         Left            =   40
         ScaleHeight     =   4635
         ScaleWidth      =   5145
         TabIndex        =   17
         Top             =   120
         Width           =   5140
         Begin VB.TextBox txtRemarks 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   2300
            Width           =   3255
         End
         Begin VB.TextBox txtModuleDes 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00FFFFFF&
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   600
            Width           =   3255
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
            MouseIcon       =   "frmBookdetails.frx":1032
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   4200
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
            MouseIcon       =   "frmBookdetails.frx":1184
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   4200
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
            MouseIcon       =   "frmBookdetails.frx":12D6
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   4200
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
            MouseIcon       =   "frmBookdetails.frx":1428
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   3720
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
            MouseIcon       =   "frmBookdetails.frx":157A
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   3720
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
            MouseIcon       =   "frmBookdetails.frx":16CC
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   3720
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
            MouseIcon       =   "frmBookdetails.frx":181E
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3240
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
            MouseIcon       =   "frmBookdetails.frx":1970
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   3240
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
            MouseIcon       =   "frmBookdetails.frx":1AC2
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   3720
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
            MouseIcon       =   "frmBookdetails.frx":1C14
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3240
            Width           =   1095
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Default         =   -1  'True
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
            MouseIcon       =   "frmBookdetails.frx":1D66
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   3240
            Width           =   1215
         End
         Begin VB.ComboBox cmbCourse 
            Appearance      =   0  'Flat
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
            TabIndex        =   2
            Text            =   "cmbCourse"
            Top             =   1300
            Width           =   2295
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
            Text            =   "cmbModule"
            Top             =   120
            Width           =   2895
         End
         Begin VB.ComboBox cmbCategory 
            Appearance      =   0  'Flat
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
            TabIndex        =   3
            Text            =   "cmbCategory"
            Top             =   1800
            Width           =   2055
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
            TabIndex        =   24
            Top             =   1300
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            Top             =   2280
            Width           =   765
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   15
            Left            =   0
            TabIndex        =   19
            Top             =   3000
            Width           =   5175
         End
      End
   End
End
Attribute VB_Name = "frmBookdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCmfordetails As ADODB.Recordset
Dim rstMmfordetails As ADODB.Recordset
Dim rstCgmfordetails As ADODB.Recordset
Dim rstDetails As ADODB.Recordset
Dim rstDatatoflex As ADODB.Recordset
Dim blnadd_Click As Boolean
Dim blnedit_Click As Boolean
Dim validity As Boolean
Dim duplicate As Boolean
Dim valid_content As Boolean
Dim find_mode As Boolean
Dim update_find As Boolean
Private Sub cmbCategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: txtRemarks.SetFocus
End Sub

Private Sub cmbCourse_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: cmbCategory.SetFocus
End Sub

Private Sub cmbModule_Click()
Set rstSelectmoduledes = New ADODB.Recordset
    rstSelectmoduledes.CursorLocation = adUseClient
    rstSelectmoduledes.Open "SELECT * FROM MODULER WHERE MODULER = '" & cmbModule.Text & " '", dbcon, adOpenStatic, adLockReadOnly
    txtModuleDes.Text = rstSelectmoduledes("MODULE_DES")
rstSelectmoduledes.Close
Set rstSelectmoduledes = Nothing
End Sub

Private Sub cmbModule_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then: cmbCourse.SetFocus
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
rstDetails.CancelUpdate
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
     rstDetails.Delete
     rstDetails.MoveNext
        If rstDetails.EOF Then
            rstDetails.MoveLast
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
'rstdetails.Close
'Set rstdetails = Nothing
'Unload Me
Form_Load
End Sub

Private Sub cmdEdit_Click()
If Check_For_Privilege(2) = True Then: Exit Sub
Enable_Controls
blnedit_Click = True
blnadd_Click = False
cmbCourse.SetFocus
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
frmFindBooksDetails.Show 1
End Sub

Private Sub cmdFirst_Click()
    If rstDetails.BOF = False Then
        rstDetails.MoveFirst
        Inforfield
        MsgBox "You are on the First Record.", vbInformation
    End If

End Sub

Private Sub cmdLast_Click()
    If rstDetails.EOF = False Then
        rstDetails.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If

End Sub

Private Sub cmdNext_Click()
    If rstDetails.EOF = False Then
        rstDetails.MoveNext
        Inforfield
    End If
    If rstDetails.EOF Then
        rstDetails.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If

End Sub

Private Sub cmdPrevious_Click()
    If rstDetails.BOF = False Then
        rstDetails.MovePrevious
        Inforfield
    End If
    If rstDetails.BOF Then
        rstDetails.MoveFirst
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
            rstDetails.AddNew
            Save_Data
        End If
        If blnedit_Click = True Then
            blnedit_Click = False
            Save_Data
        End If
    End If
End Sub
Public Sub Data_Validity()
Valid_contents_check

      If valid_content = False Then
            validity = False
      Exit Sub
      ElseIf valid_content = True Then
        validity = True
      End If
          
Duplicate_Data

      If duplicate = True Then
          MsgBox "Module is already exist...", vbExclamation
          cmbModule.SetFocus
          SendKeys "{F4}"
          validity = False
          Exit Sub
      ElseIf duplicate = False Then
          validity = True
      End If
        
End Sub

Sub Duplicate_Data()
            Set rstValidatemodule = New ADODB.Recordset
            rstValidatemodule.CursorLocation = adUseClient
            rstValidatemodule.Open "SELECT * FROM DETAILS WHERE MODULER = '" & cmbModule.Text & "'", dbcon, adOpenStatic, adLockReadOnly
                
                If rstValidatemodule.RecordCount > 0 Then
                    If blnadd_Click = True Then
                       'blnadd_Click = False
                       duplicate = True
                       
                    ElseIf blnedit_Click = True Then
                        If UCase(rstDetails("MODULER")) = UCase(cmbModule.Text) Then
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
Sub Valid_contents_check()
Course_contents_check
End Sub

Sub Course_contents_check()
Dim i As Integer
    For i = 0 To cmbCourse.ListCount - 1
        If UCase(Trim(cmbCourse.List(i))) = UCase(Trim(cmbCourse.Text)) Then
            valid_content = True
            cmbCourse.Text = UCase(Trim(cmbCourse.Text))
            Exit For
        Else
         valid_content = False
        End If
    Next i
    
'Get the result from the course list and proceed.

    If valid_content = True Then
        Module_contents_check
    ElseIf valid_content = False Then
        MsgBox "Select the Course from the List.", vbExclamation
        cmbCourse.SetFocus
        SendKeys "{F4}"
    End If
End Sub

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
        Category_contents_check
    ElseIf valid_content = False Then
        MsgBox "Select the Module from the List.", vbExclamation
        cmbModule.SetFocus
        SendKeys "{F4}"
    End If
End Sub

Sub Category_contents_check()
Dim i As Integer
For i = 0 To cmbCategory.ListCount - 1
    If UCase(Trim(cmbCategory.List(i))) = UCase(Trim(cmbCategory.Text)) Then
       valid_content = True
        cmbCategory.Text = UCase(Trim(cmbCategory.Text))
     Exit For
    Else
        valid_content = False
    End If
Next i

'Get the result from the module list and proceed.

If valid_content = True Then
    Exit Sub
ElseIf valid_content = False Then
        MsgBox "Select the Category from the List.", vbExclamation
        cmbCategory.SetFocus
        SendKeys "{F4}"
End If
End Sub
Sub Save_Data()
On Error GoTo Err
rstDetails("COURSE") = UCase(cmbCourse.Text)
rstDetails("MODULER") = UCase(cmbModule.Text)
rstDetails("MODULE_DES") = txtModuleDes.Text
rstDetails("CATEGORY") = UCase(cmbCategory.Text)
rstDetails("REMARKS") = txtRemarks.Text
rstDetails.Update
Me.Caption = "Books Details Entry   -  Record Count: " & rstDetails.RecordCount
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    If blnfind_status = True Then
        update_find = True
        Find_for_Details
    Else
       Format_Flex
       update_Flex
       Disable_Controls
       On Error Resume Next
       cmdAdd.SetFocus
    End If
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'rstdetails.Close
'Set rstissue = Nothing
'Unload Me
Form_Load
End Sub


Private Sub Form_Load()
Set rstCmfordetails = New ADODB.Recordset
    rstCmfordetails.CursorLocation = adUseClient
    rstCmfordetails.Open "SELECT * FROM COURSE ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstcoursemaster.RecordCount & "Course"
Set rstMmfordetails = New ADODB.Recordset
    rstMmfordetails.CursorLocation = adUseClient
    rstMmfordetails.Open "SELECT * FROM MODULER ORDER BY MODULER", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstcoursemaster.RecordCount & " MODULER"
Set rstCgmfordetails = New ADODB.Recordset
    rstCgmfordetails.CursorLocation = adUseClient
    rstCgmfordetails.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstcategorymaster.RecordCount & "Category"
Set rstDetails = New ADODB.Recordset
    rstDetails.CursorLocation = adUseClient
    rstDetails.Open "SELECT * FROM DETAILS ORDER BY COURSE", dbcon, adOpenStatic, adLockOptimistic
Set rstDatatoflex = New ADODB.Recordset
    rstDatatoflex.CursorLocation = adUseClient
    rstDatatoflex.Open "SELECT * FROM DETAILS ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
    
    Add_Maser_Data_to_Combo
    Inforfield
    Format_Flex
    Add_Flex
    Disable_Controls
    Me.Picture = frmStyle.Picture
    End Sub

Public Sub Add_Maser_Data_to_Combo()
If rstCmfordetails.RecordCount > 0 Then
    cmbCourse.Clear
    Do While Not rstCmfordetails.EOF
        cmbCourse.AddItem rstCmfordetails("COURSE")
    rstCmfordetails.MoveNext
    Loop
Else
 MsgBox "Please update the master information for Course.", vbExclamation
End If
If rstMmfordetails.RecordCount > 0 Then
    cmbModule.Clear
    Do While Not rstMmfordetails.EOF
        cmbModule.AddItem rstMmfordetails("MODULER")
    rstMmfordetails.MoveNext
    Loop
Else
 MsgBox "Please update the master information for Module.", vbExclamation
End If
If rstCgmfordetails.RecordCount > 0 Then
    cmbCategory.Clear
    Do While Not rstCgmfordetails.EOF
        cmbCategory.AddItem rstCgmfordetails("CATEGORY")
    rstCgmfordetails.MoveNext
    Loop
Else
 MsgBox "Please update the master information for Category.", vbExclamation
End If
End Sub
Public Sub Inforfield()
On Error Resume Next
If rstDetails.RecordCount > 0 Then
    cmbCourse.Text = rstDetails("COURSE")
    cmbModule.Text = rstDetails("MODULER")
    txtModuleDes.Text = rstDetails("MODULE_DES")
    cmbCategory.Text = rstDetails("CATEGORY")
    txtRemarks.Text = rstDetails("REMARKS")
    Me.Caption = "Books Details Entry   -  Record Count: " & rstDetails.RecordCount
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
 Else
    Button_Record_Not_Exist_Mode Me
    End If
End Sub


Public Sub Clear_Fields()
On Error Resume Next
cmbCourse.Text = "": cmbModule.Text = ""
txtModuleDes.Text = "": cmbCategory.Text = ""
txtRemarks.Text = ""
  
End Sub
Public Sub Format_Flex()
msfgBooksDetails.CellAlignment = flexAlignLeftCenter
With msfgBooksDetails
.Clear
.Cols = 3
.Rows = 1
.ColWidth(0) = 1500
.TextMatrix(0, 0) = ("Module")
.ColWidth(1) = 1500
.TextMatrix(0, 1) = "Course"
.ColWidth(2) = 1450
.TextMatrix(0, 2) = "Category"
End With
End Sub

Public Sub Add_Flex()
Dim i As Integer
Dim rcount As Integer

    If rstDatatoflex.RecordCount > 0 Then
        rcount = rstDatatoflex.RecordCount
        msfgBooksDetails.Rows = rcount + 1
        i = 1
        With msfgBooksDetails
            Do While Not rstDatatoflex.EOF
            .Row = i
            .Col = 0: .Text = rstDatatoflex("MODULER")
            .Row = i
            .Col = 1: .Text = rstDatatoflex("COURSE")
            .Row = i
            .Col = 2: .Text = rstDatatoflex("CATEGORY")
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
    rstDatatoflex.Open "SELECT * FROM DETAILS ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly


    If rstDatatoflex.RecordCount > 0 Then
        rcount = rstDatatoflex.RecordCount
        msfgBooksDetails.Rows = rcount + 1
        i = 1
        With msfgBooksDetails
            Do While Not rstDatatoflex.EOF
            .Row = i
            .Col = 0: .Text = rstDatatoflex("MODULER")
            .Row = i
            .Col = 1: .Text = rstDatatoflex("COURSE")
            .Row = i
            .Col = 2: .Text = rstDatatoflex("CATEGORY")
       i = i + 1
       rstDatatoflex.MoveNext
       
       Loop
       End With
    Else
       Format_Flex
    End If
'rstdatatoflex.Close
'Set rstdatatoflex = Nothing
End Sub
Public Sub Find_for_Details()
    rstDetails.Close
    Set rstDetails = Nothing
    rstDatatoflex.Close
    Set rstdatetoflex = Nothing
    'MsgBox Find_Val
If detailsfindparameter = 0 Then
    Set rstDetails = New ADODB.Recordset
        rstDetails.Open "SELECT* FROM DETAILS WHERE[MODULER] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
    Set rstDatatoflex = New ADODB.Recordset
        rstDatatoflex.Open "SELECT* FROM DETAILS WHERE[MODULER] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        If rstDetails.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex
            Unload frmFindBooksDetails
            cmdReturntofrecords.Enabled = True
        Else
            If update_find = True Then
                update_find = False
                Load_Initiate
                Exit Sub
            Else
            MsgBox "No records found for this search...!", vbExclamation
            frmFindBooksDetails.txtFindValue.SetFocus
            SendKeys "{Home}+{End}"
            'Load_Initiate
            Exit Sub
            End If
        End If
        
ElseIf detailsfindparameter = 1 Then

     Set rstDetails = New ADODB.Recordset
        rstDetails.Open "SELECT* FROM DETAILS WHERE[COURSE] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
     Set rstDatatoflex = New ADODB.Recordset
        rstDatatoflex.Open "SELECT* FROM DETAILS WHERE[COURSE] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        If rstDetails.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex
            Unload frmFindBooksDetails
            cmdReturntofrecords.Enabled = True
            cmdAdd.Enabled = False
        Else
            If update_find = True Then
                update_find = False
                Load_Initiate
                Exit Sub
            Else
            MsgBox "No records found for this search...!", vbExclamation
            frmFindBooksDetails.txtFindValue.SetFocus
            SendKeys "{Home}+{End}"
            'Load_Initiate
            Exit Sub
            End If
        End If
       
   ElseIf detailsfindparameter = 2 Then

     Set rstDetails = New ADODB.Recordset
        rstDetails.Open "SELECT* FROM DETAILS WHERE[CATEGORY] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
     Set rstDatatoflex = New ADODB.Recordset
        rstDatatoflex.Open "SELECT* FROM DETAILS WHERE[CATEGORY] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        If rstDetails.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex
            Unload frmFindBooksDetails
            cmdReturntofrecords.Enabled = True
            cmdAdd.Enabled = False
        Else
             If update_find = True Then
                update_find = False
                Load_Initiate
                Exit Sub
            Else
            MsgBox "No records found for this search...!", vbExclamation
            frmFindBooksDetails.txtFindValue.SetFocus
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
cmbCourse.Enabled = True
cmbModule.Enabled = True
cmbCategory.Enabled = True
txtRemarks.Enabled = True
End Sub
Public Sub Disable_Controls()
cmbCourse.Enabled = False
cmbModule.Enabled = False
cmbCategory.Enabled = False
txtRemarks.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    rstCmfordetails.Close
    Set rstCmfordetails = Nothing
    
    rstMmfordetails.Close
    Set rstMmfordetails = Nothing
    
    rstCgmfordetails.Close
    Set rstCgmfordetails = Nothing
    
    rstDetails.Close
    Set rstDetails = Nothing
    
    rstDatatoflex.Close
    Set rstDatatoflex = Nothing

End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub
