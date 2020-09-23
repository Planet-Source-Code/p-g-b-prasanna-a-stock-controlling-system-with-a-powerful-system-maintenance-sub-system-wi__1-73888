VERSION 5.00
Begin VB.Form frmBooksNotInUseEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books Not In Use Entry"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmBooksNotInUseEntry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4620
      Left            =   120
      TabIndex        =   18
      Top             =   610
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4395
         Left            =   40
         ScaleHeight     =   4395
         ScaleWidth      =   5205
         TabIndex        =   19
         Top             =   120
         Width           =   5205
         Begin VB.TextBox txtCategory 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox txtCourse 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox txtModule 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            TabIndex        =   1
            Top             =   120
            Width           =   2895
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   3000
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   3000
            Width           =   1095
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   3480
            Width           =   1215
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton cmdReport 
            Caption         =   "&Report"
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":0554
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   3000
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":06A6
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   3480
            Width           =   1095
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":07F8
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   3480
            Width           =   1095
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":094A
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   3480
            Width           =   1215
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":0A9C
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   3960
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":0BEE
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   3960
            Width           =   1575
         End
         Begin VB.CommandButton cmdExit 
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
            MouseIcon       =   "frmBooksNotInUseEntry.frx":0D40
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   3960
            Width           =   1575
         End
         Begin VB.TextBox txtModuleDes 
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
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   600
            Width           =   3255
         End
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
            TabIndex        =   5
            Top             =   2160
            Width           =   1575
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
            TabIndex        =   25
            Top             =   2160
            Width           =   720
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
            TabIndex        =   24
            Top             =   1680
            Width           =   795
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
            TabIndex        =   23
            Top             =   720
            Width           =   1620
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
            TabIndex        =   22
            Top             =   120
            Width           =   1110
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
            TabIndex        =   21
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label6 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   15
            Left            =   120
            TabIndex        =   20
            Top             =   2760
            Width           =   4935
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.Image Image1 
         Height          =   360
         Left            =   120
         Picture         =   "frmBooksNotInUseEntry.frx":0E92
         Top             =   150
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entries Of Books Which Are Not In Use."
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
         Left            =   570
         TabIndex        =   17
         Top             =   240
         Width           =   3375
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   4800
         Picture         =   "frmBooksNotInUseEntry.frx":1594
         ToolTipText     =   "Application Help"
         Top             =   165
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmBooksNotInUseEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstBooksnotinuse As ADODB.Recordset
Dim rstMmforcurstock As ADODB.Recordset
Dim rstCgmforcurstock As ADODB.Recordset
Dim rstDatatoflex As ADODB.Recordset
Dim blnadd_Click As Boolean
Dim blnedit_Click As Boolean
Dim validity As Boolean
Dim duplicate As Boolean
Dim valid_content As Boolean
Dim find_mode As Boolean
Dim update_find As Boolean
Dim invalid_quantity As Boolean

Private Sub cmdAdd_Click()
If Check_For_Privilege(1) = True Then: Exit Sub
Enable_Controls
blnadd_Click = True
blnedit_Click = False
Clear_Fields
txtModule.SetFocus
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
rstBooksnotinuse.CancelUpdate
Clear_Fields
Inforfield
Disable_Controls
On Error Resume Next
cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Err GoTo Err
If Check_For_Privilege(3) = True Then: Exit Sub
If MsgBox("Are you sure you want to delete the current Record ?", vbQuestion + vbYesNo) = vbYes Then
        On Error Resume Next
        rstBooksnotinuse.Delete
        rstBooksnotinuse.MoveNext
            If rstBooksnotinuse.EOF Then
                rstBooksnotinuse.MoveLast
            End If
        Clear_Fields
        Inforfield
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
txtModule.SetFocus
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFirst_Click()
    If rstBooksnotinuse.BOF = False Then
        rstBooksnotinuse.MoveFirst
        Inforfield
        MsgBox "You are on the First Record.", vbInformation
    End If

End Sub

Private Sub cmdLast_Click()
    If rstBooksnotinuse.EOF = False Then
        rstBooksnotinuse.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If

End Sub

Private Sub cmdNext_Click()
    If rstBooksnotinuse.EOF = False Then
        rstBooksnotinuse.MoveNext
        Inforfield
    End If
    If rstBooksnotinuse.EOF Then
        rstBooksnotinuse.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If

End Sub

Private Sub cmdPrevious_Click()
    If rstBooksnotinuse.BOF = False Then
        rstBooksnotinuse.MovePrevious
        Inforfield
    End If
    If rstBooksnotinuse.BOF Then
        rstBooksnotinuse.MoveFirst
        Inforfield
        MsgBox "You are on the first Record.", vbInformation
    End If
End Sub

Private Sub cmdReport_Click()
If Check_For_Privilege(4) = True Then: Exit Sub
Get_Report_for_Books_not_in_stock
End Sub

Private Sub cmdSave_Click()
Data_Validity
    If validity = True Then
        If blnadd_Click = True Then
            blnadd_Click = False
            rstBooksnotinuse.AddNew
            Save_Data
        End If
        If blnedit_Click = True Then
            blnedit_Click = False
            Save_Data
        End If
    ElseIf validity = False Then
        Exit Sub
    End If
End Sub
Public Sub Data_Validity()
If txtQuantity = "" Then
    MsgBox "You must enter a value for Quantity.", vbExclamation
        txtQuantity.SetFocus
        validity = False
    Exit Sub
End If
Duplicate_Data

      If duplicate = True Then
          MsgBox "Module is already exist!!", vbExclamation
          txtModule.SetFocus
          SendKeys "{Home}+{End}"
          validity = False
          Exit Sub
      ElseIf duplicate = False Then
          validity = True
      End If
      
End Sub
 Sub Duplicate_Data()
            Set rstValidatemodule = New ADODB.Recordset
            rstValidatemodule.CursorLocation = adUseClient
            rstValidatemodule.Open "SELECT * FROM BOOKS_NOT_IN_USE WHERE MODULER = '" & txtModule & "'", dbcon, adOpenStatic, adLockReadOnly
                
                If rstValidatemodule.RecordCount > 0 Then
                    If blnadd_Click = True Then
                        duplicate = True
                       
                    ElseIf blnedit_Click = True Then
                        If UCase(rstBooksnotinuse("MODULER")) = UCase(txtModule.Text) Then
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
Sub Save_Data()
On Error GoTo Err
     rstBooksnotinuse("MODULER") = UCase(txtModule.Text)
     rstBooksnotinuse("MODULE_DES") = txtModuleDes.Text
     rstBooksnotinuse("COURSE") = UCase(txtCourse.Text)
     rstBooksnotinuse("CATEGORY") = UCase(txtCategory.Text)
     rstBooksnotinuse("QUANTITY") = Val(txtQuantity.Text)
     rstBooksnotinuse.Update
     Me.Caption = "Books Not In Use Entry   -  Record Count: " & rstBooksnotinuse.RecordCount
     Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
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
    Set rstBooksnotinuse = New ADODB.Recordset
    rstBooksnotinuse.Open "SELECT * FROM BOOKS_NOT_IN_USE ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
    Inforfield
    Disable_Controls
    Me.Picture = frmStyle.Picture
    End Sub
End Sub
Public Sub Inforfield()
On Error Resume Next
If rstBooksnotinuse.RecordCount > 0 Then
    txtModule.Text = rstBooksnotinuse("MODULER")
    txtModuleDes.Text = rstBooksnotinuse("MODULE_DES")
    txtCourse.Text = rstBooksnotinuse("COURSE")
    txtCategory.Text = rstBooksnotinuse("CATEGORY")
    txtQuantity.Text = rstBooksnotinuse("QUANTITY")
    Me.Caption = "Books Not In Use Entry   -  Record Count: " & rstBooksnotinuse.RecordCount
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
 Else
    Button_Record_Not_Exist_Mode Me
    End If
End Sub


Public Sub Clear_Fields()
On Error Resume Next
txtModule = ""
txtModuleDes = ""
txtCourse = ""
txtQuantity = ""
txtCategory = ""
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
txtModule.Enabled = True
txtModuleDes.Enabled = True
txtCourse.Enabled = True
txtQuantity.Enabled = True
txtCategory.Enabled = True
End Sub
Public Sub Disable_Controls()
txtModule.Enabled = False
txtModuleDes.Enabled = False
txtCourse.Enabled = False
txtQuantity.Enabled = False
txtCategory.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
rstBooksnotinuse.Close
Set rstBooksnotinuse = Nothing
End Sub

Private Sub txtCategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQuantity.SetFocus
End If
End Sub

Private Sub txtCourse_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCategory.SetFocus
End If
End Sub

Private Sub txtModule_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtModuleDes.SetFocus
End If
End Sub

Private Sub txtModuleDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCourse.SetFocus
End If
End Sub

Private Sub txtQuantity_Change()
If Not IsNumeric(txtQuantity) Then
    txtQuantity = ""
End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub
