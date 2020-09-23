VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBookReceipt 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books Receipt Entry"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   Icon            =   "frmBookReceipt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   10995
      Begin VB.Image Image3 
         Height          =   360
         Left            =   120
         Picture         =   "frmBookReceipt.frx":000C
         Top             =   160
         Width           =   360
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter All The Books Receipt Entries Here, When Books Are Received, This Should Be Updated."
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
         TabIndex        =   36
         Top             =   240
         Width           =   8775
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   10440
         Picture         =   "frmBookReceipt.frx":0776
         ToolTipText     =   "Application Help"
         Top             =   165
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   5520
      TabIndex        =   32
      Top             =   610
      Width           =   5600
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5850
         Left            =   40
         ScaleHeight     =   5850
         ScaleWidth      =   5505
         TabIndex        =   33
         Top             =   120
         Width           =   5500
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
            Left            =   200
            MouseIcon       =   "frmBookReceipt.frx":0EE0
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   5400
            Width           =   3735
         End
         Begin VB.CommandButton cmdOptimize 
            Caption         =   "Optimize..."
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
            Left            =   4080
            MouseIcon       =   "frmBookReceipt.frx":1032
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   5400
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid msfgBooksReceipt 
            Height          =   5295
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   9340
            _Version        =   393216
            BackColor       =   16642796
            AllowUserResizing=   1
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   610
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5800
         Left            =   40
         ScaleHeight     =   5805
         ScaleWidth      =   5190
         TabIndex        =   18
         Top             =   120
         Width           =   5190
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
            MouseIcon       =   "frmBookReceipt.frx":1184
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   5400
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
            MouseIcon       =   "frmBookReceipt.frx":12D6
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   5400
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
            MouseIcon       =   "frmBookReceipt.frx":1428
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   5400
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
            MouseIcon       =   "frmBookReceipt.frx":157A
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   4920
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
            MouseIcon       =   "frmBookReceipt.frx":16CC
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   4920
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
            MouseIcon       =   "frmBookReceipt.frx":181E
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   4920
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
            MouseIcon       =   "frmBookReceipt.frx":1970
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   4440
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
            MouseIcon       =   "frmBookReceipt.frx":1AC2
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   4440
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
            MouseIcon       =   "frmBookReceipt.frx":1C14
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   4920
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
            MouseIcon       =   "frmBookReceipt.frx":1D66
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   4440
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
            MouseIcon       =   "frmBookReceipt.frx":1EB8
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   4440
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
            Width           =   2895
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
            TabIndex        =   3
            Top             =   2880
            Width           =   1575
         End
         Begin VB.ComboBox cmbCourse 
            BackColor       =   &H80000008&
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
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   1320
            Width           =   2415
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
            TabIndex        =   20
            Top             =   600
            Width           =   3255
         End
         Begin VB.ComboBox cmbCategory 
            BackColor       =   &H80000008&
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
            Height          =   315
            Left            =   1800
            TabIndex        =   19
            Top             =   1800
            Width           =   2295
         End
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
            Top             =   3480
            Width           =   3255
         End
         Begin MSComCtl2.DTPicker dtpIDate 
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   48627715
            CurrentDate     =   39668
         End
         Begin VB.Label Label11 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            Height          =   15
            Left            =   120
            TabIndex        =   31
            Top             =   4200
            Width           =   5055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Current Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblcurbal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            Top             =   2925
            Width           =   60
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Date"
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
            TabIndex        =   28
            Top             =   2280
            Width           =   1095
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
            TabIndex        =   27
            Top             =   2880
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
            TabIndex        =   26
            Top             =   1800
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label8 
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
            TabIndex        =   22
            Top             =   3480
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frmBookReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCmforreceipt As ADODB.Recordset
Dim rstMmforreceipt As ADODB.Recordset
Dim rstCgmforreceipt As ADODB.Recordset
Dim rstReceipt As ADODB.Recordset
Dim rstDatatoflex As ADODB.Recordset
Dim blnadd_Click As Boolean
Dim blnedit_Click As Boolean
Dim blndelete_Click As Boolean
Dim blndeletion  As Boolean
Dim validity As Boolean
Dim duplicate As Boolean
Dim valid_content As Boolean
Dim invalid_quantity As Boolean
Dim find_mode As Boolean
Dim update_find As Boolean
Dim update_cur_stock As Boolean
Dim old_qty, new_qty As Integer
Dim f_module, s_module, old_module, new_module As String

Private Sub cmbModule_Click()
On Error GoTo Ret_Error
Set rstInfofromdetails = New ADODB.Recordset
    rstInfofromdetails.CursorLocation = adUseClient
    rstInfofromdetails.Open "SELECT * FROM DETAILS WHERE MODULER = '" & cmbModule.Text & " '", dbcon, adOpenStatic, adLockReadOnly
    txtModuleDes.Text = rstInfofromdetails("MODULE_DES")
    cmbCourse.Text = rstInfofromdetails("COURSE")
    cmbCategory.Text = rstInfofromdetails("CATEGORY")
    Get_Cur_Stock
    
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
    'dtpIDate.SetFocus
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
blnedit_Click = False
new_module = ""
old_module = ""
rstReceipt.CancelUpdate
Clear_Fields
Inforfield
Get_Cur_Stock
Disable_Controls
On Error Resume Next
cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err
If Check_For_Privilege(3) = True Then: Exit Sub
blndelete_Click = True
If MsgBox("Are you sure you want to delete the current Record ?", vbQuestion + vbYesNo) = vbYes Then
    Check_Cur_Stock
        If blndeletion = False Then
            frmCurrentStock.Show 1
            Exit Sub
        ElseIf blndeletion = True Then
            On Error Resume Next
            'MsgBox "TEstingngdgdgd"
            rstReceipt.Delete
            rstReceipt.MoveNext
                If rstReceipt.EOF Then
                    rstReceipt.MoveLast
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
End If
Get_Cur_Stock
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
'rstreceipt.Close
'Set rstreceipt = Nothing
Form_Load
End Sub

Private Sub cmdEdit_Click()
If Check_For_Privilege(2) = True Then: Exit Sub
Enable_Controls
blnedit_Click = True
blnadd_Click = False
cmbModule.SetFocus
old_qty = Val(txtQuantity)
f_module = UCase(cmbModule)
old_module = cmbModule
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
Get_Cur_Stock
'lblcurbal.Caption = ""
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
lblcurbal.Caption = ""
blnedit_Click = False
new_module = ""
old_module = ""
frmFindBookReceipt.Show 1
End Sub

Private Sub cmdFirst_Click()
    If rstReceipt.BOF = False Then
        rstReceipt.MoveFirst
        Inforfield
        MsgBox "You are on the First Record.", vbInformation
    End If
Get_Cur_Stock
End Sub

Private Sub cmdLast_Click()
    If rstReceipt.EOF = False Then
        rstReceipt.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If
Get_Cur_Stock
End Sub

Private Sub cmdNext_Click()
    If rstReceipt.EOF = False Then
        rstReceipt.MoveNext
        Inforfield
    End If
    If rstReceipt.EOF Then
        rstReceipt.MoveLast
        Inforfield
        MsgBox "You are on the Last Record.", vbInformation
    End If
Get_Cur_Stock
End Sub

Private Sub cmdOptimize_Click()
If MsgBox("Do you need to go to Optimize Dialog?", vbQuestion + vbYesNo) = vbYes Then
    Unload Me
    frmPerformanceoptimizer.Show 1
End If
End Sub

Private Sub cmdPrevious_Click()
    If rstReceipt.BOF = False Then
        rstReceipt.MovePrevious
        Inforfield
    End If
    If rstReceipt.BOF Then
        rstReceipt.MoveFirst
        Inforfield
        MsgBox "You are on the first Record.", vbInformation
    End If
Get_Cur_Stock
End Sub

Private Sub cmdReturntofrecords_Click()
Load_Initiate
End Sub

Private Sub cmdSave_Click()
Data_Validity
    If validity = True Then
        If blnadd_Click = True Then
            'blnadd_Click = False
            rstReceipt.AddNew
            Save_Data
        End If
        If blnedit_Click = True Then
            'blnedit_Click = False
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
      
      quantity_validate
        If invalid_quantity = False Then
            MsgBox "Invalid Quantity...!", vbExclamation
            txtQuantity.SetFocus
            SendKeys "{HOME}+{END}"
            validity = False
          Exit Sub
        ElseIf invalid_quantity = True Then
            validity = True
        End If
        
      Check_Cur_Stock
        If update_cur_stock = False Then
            validity = False
            Exit Sub
        ElseIf update_cur_stock = True Then
            validity = True
        End If
  
End Sub
Public Sub quantity_validate()
If Val(txtQuantity) = 0 Then
    invalid_quantity = False
Else
    invalid_quantity = True
End If
End Sub
Sub Valid_contents_check()
Module_contents_check
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
        Course_contents_check
    ElseIf valid_content = False Then
        MsgBox "Select the Module from the List.", vbExclamation
        cmbModule.SetFocus
        SendKeys "{F4}"
    End If
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
        Category_contents_check
    ElseIf valid_content = False Then
        MsgBox "Select the Course from the List.", vbExclamation
        cmbCourse.SetFocus
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
Calculate_Cur_Stock
On Error GoTo Err
    rstReceipt("MODULER") = UCase(cmbModule.Text)
    rstReceipt("MODULE_DES") = txtModuleDes.Text
    rstReceipt("COURSE") = UCase(cmbCourse.Text)
    rstReceipt("CATEGORY") = UCase(cmbCategory.Text)
    rstReceipt("DATE_RECEIVED") = dtpIDate.Value
    rstReceipt("QUANTITY") = Val(txtQuantity.Text)
    rstReceipt("REMARKS") = txtRemarks.Text
    rstReceipt.Update
    Me.Caption = "Books Receipt Entry   -  Record Count: " & rstReceipt.RecordCount
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    If blnfind_status = True Then
        update_find = True
        Find_for_Details
    Else
    
    Format_Flex
    update_Flex
    End If
    blnadd_Click = False
    blnedit_Click = False
    Get_Cur_Stock
    Disable_Controls
    On Error Resume Next
    cmdAdd.SetFocus
Exit Sub
Err:
 MsgBox Err.Description & " _ " & Err.Number, vbCritical
 'Unload Me
 'rstreceipt.Close
 'Set rstreceipt = Nothing
 Form_Load
End Sub



Private Sub dtpIDate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQuantity.SetFocus
End If
End Sub

Private Sub Form_Load()
'Set rstcmforreceipt = New ADODB.Recordset
'    rstcmforreceipt.CursorLocation = adUseClient
'    rstcmforreceipt.Open "SELECT * FROM COURSE ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstcoursemaster.RecordCount & "Course"
Set rstMmforreceipt = New ADODB.Recordset
    rstMmforreceipt.CursorLocation = adUseClient
    rstMmforreceipt.Open "SELECT * FROM MODULER ORDER BY MODULER", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstcoursemaster.RecordCount & " MODULER"
'Set rstcgmforreceipt = New ADODB.Recordset
'    rstcgmforreceipt.CursorLocation = adUseClient
'    rstcgmforreceipt.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstcategorymaster.RecordCount & "Category"
Set rstReceipt = New ADODB.Recordset
    rstReceipt.CursorLocation = adUseClient
    rstReceipt.Open "SELECT * FROM B_RECEIPT ORDER BY COURSE", dbcon, adOpenStatic, adLockOptimistic
If enabledataviewbooksreceipt = 1 Then
Set rstDatatoflex = New ADODB.Recordset
    rstDatatoflex.CursorLocation = adUseClient
    rstDatatoflex.Open "SELECT * FROM B_RECEIPT ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
End If
    
    Add_Maser_Data_to_Combo
    Inforfield
    Format_Flex
    Add_Flex
    Disable_Controls
    Me.Picture = frmStyle.Picture
    End Sub

Public Sub Add_Maser_Data_to_Combo()
If rstMmforreceipt.RecordCount > 0 Then
    cmbModule.Clear
    Do While Not rstMmforreceipt.EOF
        cmbModule.AddItem rstMmforreceipt("MODULER")
    rstMmforreceipt.MoveNext
    Loop
Else
 MsgBox "Please update the master information for Module.", vbExclamation
End If

'If rstcmforreceipt.RecordCount > 0 Then
'    cmbCourse.Clear
'    Do While Not rstcmforreceipt.EOF
'        cmbCourse.AddItem rstcmforreceipt("COURSE")
'    rstcmforreceipt.MoveNext
'    Loop
'Else
' MsgBox "Please update the master information", vbExclamation
'End If
'If rstcgmforreceipt.RecordCount > 0 Then
'    cmbCategory.Clear
'    Do While Not rstcgmforreceipt.EOF
'        cmbCategory.AddItem rstcgmforreceipt("CATEGORY")
'    rstcgmforreceipt.MoveNext
'    Loop
'Else
' MsgBox "Please update the master information", vbExclamation
'End If
End Sub
Public Sub Inforfield()
On Error Resume Next
If rstReceipt.RecordCount > 0 Then
    cmbModule.Text = rstReceipt("MODULER")
    txtModuleDes.Text = rstReceipt("MODULE_DES")
    cmbCourse.Text = rstReceipt("COURSE")
    cmbCategory.Text = rstReceipt("CATEGORY")
    dtpIDate.Value = rstReceipt("DATE_RECEIVED")
    txtQuantity.Text = rstReceipt("QUANTITY")
    txtRemarks.Text = rstReceipt("REMARKS")
    Me.Caption = "Books Receipt Entry   -  Record Count: " & rstReceipt.RecordCount
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    Get_Cur_Stock
 Else
    Button_Record_Not_Exist_Mode Me
    dtpIDate.Value = Date
    End If
End Sub


Public Sub Clear_Fields()
On Error Resume Next
    cmbModule.Text = ""
    txtModuleDes.Text = ""
    cmbCourse.Text = ""
    cmbCategory.Text = ""
    dtpIDate.Value = Date
    txtQuantity.Text = ""
    txtRemarks.Text = ""
  
End Sub
Public Sub Format_Flex()
If enabledataviewbooksreceipt = 1 Then
msfgBooksReceipt.CellAlignment = flexAlignLeftCenter
With msfgBooksReceipt
.Clear
.Cols = 5
.Rows = 1
.ColWidth(0) = 1000
.TextMatrix(0, 0) = "Date Received"
.ColWidth(1) = 1500
.TextMatrix(0, 1) = "Module"
.ColWidth(2) = 1500
.TextMatrix(0, 2) = "Course"
.ColWidth(3) = 700
.TextMatrix(0, 3) = "Quantity"
.ColWidth(3) = 1200
.TextMatrix(0, 4) = "Category"
End With
End If
End Sub

Public Sub Add_Flex()
If enabledataviewbooksreceipt = 1 Then
Dim i As Integer
Dim rcount As Integer

    If rstDatatoflex.RecordCount > 0 Then
        rcount = rstDatatoflex.RecordCount
        msfgBooksReceipt.Rows = rcount + 1
        i = 1
        With msfgBooksReceipt
            Do While Not rstDatatoflex.EOF
            .Row = i
            .Col = 0: .Text = Format(rstDatatoflex("DATE_RECEIVED"), "dd/mm/yyyy")
            .Row = i
            .Col = 1: .Text = rstDatatoflex("MODULER")
            .Row = i
            .Col = 2: .Text = rstDatatoflex("COURSE")
            .Row = i
            .Col = 3: .Text = rstDatatoflex("QUANTITY")
            .Row = i
            .Col = 4: .Text = rstDatatoflex("CATEGORY")
       i = i + 1
       rstDatatoflex.MoveNext
       
       Loop
       End With
    Else
       Format_Flex
    End If
End If
End Sub
Public Sub update_Flex()
If enabledataviewbooksreceipt = 1 Then
Dim i As Integer
Dim rcount As Integer

rstDatatoflex.Close
Set rstDatatoflex = Nothing

Set rstDatatoflex = New ADODB.Recordset
    rstDatatoflex.CursorLocation = adUseClient
    rstDatatoflex.Open "SELECT * FROM B_RECEIPT ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly


    If rstDatatoflex.RecordCount > 0 Then
        rcount = rstDatatoflex.RecordCount
        msfgBooksReceipt.Rows = rcount + 1
        i = 1
        With msfgBooksReceipt
            Do While Not rstDatatoflex.EOF
           .Row = i
            .Col = 0: .Text = Format(rstDatatoflex("DATE_RECEIVED"), "dd/mm/yyyy")
            .Row = i
            .Col = 1: .Text = rstDatatoflex("MODULER")
            .Row = i
            .Col = 2: .Text = rstDatatoflex("COURSE")
            .Row = i
            .Col = 3: .Text = rstDatatoflex("QUANTITY")
            .Row = i
            .Col = 4: .Text = rstDatatoflex("CATEGORY")
       i = i + 1
       rstDatatoflex.MoveNext
       
       Loop
       End With
    Else
       Format_Flex
    End If
End If
End Sub
Public Sub Find_for_Details()
If enabledataviewbooksreceipt = 1 Then
    rstReceipt.Close
    Set rstReceipt = Nothing
    rstDatatoflex.Close
    Set rstdatetoflex = Nothing
End If

If detailsfindparameter = 0 Then
    If blnenabledate = False Then
        Set rstReceipt = New ADODB.Recordset
            rstReceipt.Open "SELECT* FROM B_RECEIPT WHERE[MODULER] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
        If enabledataviewbooksreceipt = 1 Then
            Set rstDatatoflex = New ADODB.Recordset
                rstDatatoflex.Open "SELECT* FROM B_RECEIPT WHERE[MODULER] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        End If
            If rstReceipt.RecordCount > 0 Then
                Inforfield
                Format_Flex
                Add_Flex
                Unload frmFindBookReceipt
                cmdReturntofrecords.Enabled = True
            Else
                If update_find = True Then
                    update_find = False
                    Load_Initiate
                 Exit Sub
                Else
                    MsgBox "No records found for this search...!", vbExclamation
                    frmFindBookReceipt.txtFindValue.SetFocus
                    SendKeys "{Home}+{End}"
                    'Load_Initiate
                    Exit Sub
                End If
            End If
    ElseIf blnenabledate = True Then

        Set rstReceipt = New ADODB.Recordset
          rstReceipt.Open "SELECT* FROM B_RECEIPT WHERE[MODULER] Like '" & Find_Val & "%' AND DATE_RECEIVED >= " & "#" & Format(fromdate, "dd/mm/yyyy") & "#" & " AND DATE_RECEIVED <= " & "#" & Format(todate, "dd/mm/yyyy") & "#", dbcon, adOpenStatic, adLockOptimistic
        Set rstDatatoflex = New ADODB.Recordset
        If enabledataviewbooksreceipt = 1 Then
            rstDatatoflex.Open "SELECT* FROM B_RECEIPT WHERE[MODULER] Like '" & Find_Val & "%' AND DATE_RECEIVED >= " & "#" & Format(fromdate, "dd/mm/yyyy") & "#" & " AND DATE_RECEIVED <= " & "#" & Format(todate, "dd/mm/yyyy") & "#", dbcon, adOpenStatic, adLockReadOnly
        End If
            If rstReceipt.RecordCount > 0 Then
                Inforfield
                Format_Flex
                Add_Flex
                Unload frmFindBookReceipt
                cmdReturntofrecords.Enabled = True
            Else
                If update_find = True Then
                    update_find = False
                    Load_Initiate
                 Exit Sub
                Else
                    MsgBox "No records found for this search...!", vbExclamation
                    frmFindBookReceipt.txtFindValue.SetFocus
                    SendKeys "{Home}+{End}"
                    'Load_Initiate
                    Exit Sub
                End If
            End If
    End If
ElseIf detailsfindparameter = 1 Then
    If blnenabledate = False Then
        Set rstReceipt = New ADODB.Recordset
            rstReceipt.Open "SELECT* FROM B_RECEIPT WHERE[COURSE] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
        If enabledataviewbooksreceipt = 1 Then
            Set rstDatatoflex = New ADODB.Recordset
                rstDatatoflex.Open "SELECT* FROM B_RECEIPT WHERE[COURSE] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
        End If
            If rstReceipt.RecordCount > 0 Then
                Inforfield
                Format_Flex
                Add_Flex
                Unload frmFindBookReceipt
                cmdReturntofrecords.Enabled = True
            Else
                If update_find = True Then
                    update_find = False
                    Load_Initiate
                    Exit Sub
                Else
                    MsgBox "No records found for this search...!", vbExclamation
                    frmFindBookReceipt.txtFindValue.SetFocus
                    SendKeys "{Home}+{End}"
                    'Load_Initiate
                    Exit Sub
                End If
             End If
    ElseIf blnenabledate = True Then
        Set rstReceipt = New ADODB.Recordset
            rstReceipt.Open "SELECT* FROM B_RECEIPT WHERE[COURSE] Like '" & Find_Val & "%' AND DATE_RECEIVED >= " & "#" & Format(fromdate, "dd/mm/yyyy") & "#" & " AND DATE_RECEIVED <= " & "#" & Format(todate, "dd/mm/yyyy") & "#", dbcon, adOpenStatic, adLockOptimistic
       If enabledataviewbooksreceipt = 1 Then
        Set rstDatatoflex = New ADODB.Recordset
            rstDatatoflex.Open "SELECT* FROM B_RECEIPT WHERE[COURSE] Like '" & Find_Val & "%' AND DATE_RECEIVED >= " & "#" & Format(fromdate, "dd/mm/yyyy") & "#" & " AND DATE_RECEIVED <= " & "#" & Format(todate, "dd/mm/yyyy") & "#", dbcon, adOpenStatic, adLockReadOnly
       End If
            If rstReceipt.RecordCount > 0 Then
                Inforfield
                Format_Flex
                Add_Flex
                Unload frmFindBookReceipt
                cmdReturntofrecords.Enabled = True
            Else
                If update_find = True Then
                    update_find = False
                    Load_Initiate
                    Exit Sub
                Else
                    MsgBox "No records found for this search...!", vbExclamation
                    frmFindBookReceipt.txtFindValue.SetFocus
                    SendKeys "{Home}+{End}"
                    'Load_Initiate
                    Exit Sub
                End If
             End If
    End If
ElseIf detailsfindparameter = 2 Then
    If blnenabledate = False Then
        Set rstReceipt = New ADODB.Recordset
            rstReceipt.Open "SELECT* FROM B_RECEIPT WHERE[CATEGORY] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockOptimistic
      If enabledataviewbooksreceipt = 1 Then
        Set rstDatatoflex = New ADODB.Recordset
            rstDatatoflex.Open "SELECT* FROM B_RECEIPT WHERE[CATEGORY] Like '" & Find_Val & "%'", dbcon, adOpenStatic, adLockReadOnly
      End If
            If rstReceipt.RecordCount > 0 Then
                Inforfield
                Format_Flex
                Add_Flex
                Unload frmFindBookReceipt
                cmdReturntofrecords.Enabled = True
            Else
                If update_find = True Then
                    update_find = False
                    Load_Initiate
                    Exit Sub
                Else
                    MsgBox "No records found for this search...!", vbExclamation
                    frmFindBookReceipt.txtFindValue.SetFocus
                    SendKeys "{Home}+{End}"
                    'Load_Initiate
                    Exit Sub
                End If
            End If
        ElseIf blnenabledate = True Then
            Set rstReceipt = New ADODB.Recordset
                rstReceipt.Open "SELECT* FROM B_RECEIPT WHERE[CATEGORY] Like '" & Find_Val & "%' AND DATE_RECEIVED >= " & "#" & Format(fromdate, "dd/mm/yyyy") & "#" & " AND DATE_RECEIVED <= " & "#" & Format(todate, "dd/mm/yyyy") & "#", dbcon, adOpenStatic, adLockOptimistic
          If enabledataviewbooksreceipt = 1 Then
            Set rstDatatoflex = New ADODB.Recordset
                rstDatatoflex.Open "SELECT* FROM B_RECEIPT WHERE[CATEGORY] Like '" & Find_Val & "%' AND DATE_RECEIVED >= " & "#" & Format(fromdate, "dd/mm/yyyy") & "#" & " AND DATE_RECEIVED <= " & "#" & Format(todate, "dd/mm/yyyy") & "#", dbcon, adOpenStatic, adLockReadOnly
          End If
                If rstReceipt.RecordCount > 0 Then
                    Inforfield
                    Format_Flex
                    Add_Flex
                    Unload frmFindBookReceipt
                    cmdReturntofrecords.Enabled = True
                Else
                    If update_find = True Then
                        update_find = False
                        Load_Initiate
                        Exit Sub
                    Else
                        MsgBox "No records found for this search...!", vbExclamation
                        frmFindBookReceipt.txtFindValue.SetFocus
                        SendKeys "{Home}+{End}"
                        'Load_Initiate
                        Exit Sub
                    End If
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

Public Sub Calculate_Cur_Stock()
On Error GoTo Err
Dim current_qty, receipt_qty, calculate_qty As Integer
Set rstcur_item_qty = New ADODB.Recordset
    rstcur_item_qty.Open "SELECT* FROM CUR_STOCK WHERE[MODULER] = '" & cmbModule.Text & "'", dbcon, adOpenStatic, adLockOptimistic
        If rstcur_item_qty.RecordCount > 0 Then
            If blnedit_Click = True Then
                new_qty = Val(txtQuantity)
                    If new_qty > old_qty Then
                        current_qty = Val(rstcur_item_qty("QUANTITY"))
                        calculate_qty = current_qty + (new_qty - old_qty)
                        rstcur_item_qty("QUANTITY") = Val(calculate_qty)
                        rstcur_item_qty.Update
                    ElseIf old_qty > new_qty Then
                        current_qty = Val(rstcur_item_qty("QUANTITY"))
                        calculate_qty = current_qty - (old_qty - new_qty)
                        rstcur_item_qty("QUANTITY") = Val(calculate_qty)
                        rstcur_item_qty.Update
                    End If
            ElseIf blnadd_Click = True Then
                current_qty = Val(rstcur_item_qty("QUANTITY"))
                receipt_qty = Val(txtQuantity)
                calculate_qty = current_qty + receipt_qty
                rstcur_item_qty("QUANTITY") = calculate_qty
                rstcur_item_qty.Update
            ElseIf blndelete_Click = True Then
                current_qty = Val(rstcur_item_qty("QUANTITY"))
                receipt_qty = Val(txtQuantity)
                calculate_qty = current_qty - receipt_qty
                rstcur_item_qty("QUANTITY") = calculate_qty
                rstcur_item_qty.Update
                blndeletion = True
            End If
        Else
            Add_New_Item_to_Cur_Stock
        End If
rstcur_item_qty.Close
Set rstcur_item_qty = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'rstcur_item_qty.Close
'Set rstcur_item_qty = Nothing
'Unload Me
Form_Load
End Sub

Public Sub Add_New_Item_to_Cur_Stock()
On Error GoTo Err
Set rstcurstock = New ADODB.Recordset
    rstcurstock.CursorLocation = adUseClient
    rstcurstock.Open "SELECT * FROM CUR_STOCK", dbcon, adOpenStatic, adLockOptimistic
    rstcurstock.AddNew
    rstcurstock("MODULER") = UCase(cmbModule.Text)
    rstcurstock("MODULE_DES") = txtModuleDes.Text
    rstcurstock("COURSE") = UCase(cmbCourse.Text)
    rstcurstock("CATEGORY") = UCase(cmbCategory.Text)
    rstcurstock("QUANTITY") = Val(txtQuantity.Text)
    rstcurstock.Update
    
    rstcurstock.Close
    Set rstcursotck = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
'rstcurstock.Close
'Set rstcurstock = Nothing
Form_Load
End Sub
Public Sub Check_Cur_Stock()
'Dim new_module As String
Set rstCur_Check_qty = New ADODB.Recordset
    rstCur_Check_qty.Open "SELECT* FROM CUR_STOCK WHERE[MODULER] = '" & cmbModule.Text & "'", dbcon, adOpenStatic, adLockReadOnly
    
'--------------------------------------------------------------------------
new_module = cmbModule
  If blnedit_Click = True Then
      
    If UCase(new_module) <> UCase(old_module) Then
       'blnedit_Click = False
       'new_module = ""
       'old_module = ""
       MsgBox "This updation cannot be done...!" & vbCrLf & _
       "Please delete the record with " & old_module & " first and" & _
       " enter a new record for " & new_module & ".", vbExclamation
       update_cur_stock = False
       cmbModule.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
   End If
'-------------------------------------------------------------------------------
   
    If blnedit_Click = True Then
        If rstCur_Check_qty.RecordCount > 0 Then
          new_qty = txtQuantity
            If Val(rstCur_Check_qty("QUANTITY")) < (old_qty - new_qty) Then
                MsgBox "Please check the Current Stock...!", vbExclamation
                    update_cur_stock = False
                    cmdCancel_Click
                    frmCurrentStock.Show 1
            Else
                update_cur_stock = True
            End If
        Else
            MsgBox "Please update the Current Stock...!", vbExclamation
            update_cur_stock = False
            cmdCancel_Click
            frmCurrentStock.Show 1
        End If
        
    ElseIf blnadd_Click = True Then
         update_cur_stock = True
         
    ElseIf blndelete_Click = True Then
      If rstCur_Check_qty.RecordCount > 0 Then
          new_qty = txtQuantity
            If Val(rstCur_Check_qty("QUANTITY")) < new_qty Then
                If MsgBox("Current Stock is not enough to update." & vbCrLf & "Click Yes to update the Current Stock manually." & vbCrLf & "Click No to ignore the updation & Delete the record.", vbExclamation + vbYesNo) = vbYes Then
                    blndeletion = False
                Else
                    blndeletion = True
                End If
           Else
                 Calculate_Cur_Stock
           End If
      Else
        If blndelete_Click = True Then
            If MsgBox("This item is not in Current Stock." & vbCrLf & "Click Yes to update the Current Stock manually." & vbCrLf & "Click No to ignore the updation & Delete the record.", vbExclamation + vbYesNo) = vbYes Then
                blndeletion = False
            Else
                blndeletion = True
            End If
        End If
      End If
   End If
   
rstCur_Check_qty.Close
Set rstCur_Check_qty = Nothing
End Sub

Public Sub Get_Cur_Stock()
Set rstCur_get_qty = New ADODB.Recordset
    rstCur_get_qty.Open "SELECT* FROM CUR_STOCK WHERE[MODULER] = '" & cmbModule.Text & "'", dbcon, adOpenStatic, adLockReadOnly
    If rstCur_get_qty.RecordCount > 0 Then
        lblcurbal.Caption = Val(rstCur_get_qty("QUANTITY"))
    Else
        lblcurbal.Caption = "0"
    End If
rstCur_get_qty.Close
Set rstCur_get_qty = Nothing
End Sub

Public Sub Enable_Controls()
cmbModule.Enabled = True
dtpIDate.Enabled = True
txtQuantity.Enabled = True
txtRemarks.Enabled = True
End Sub
Public Sub Disable_Controls()
cmbModule.Enabled = False
dtpIDate.Enabled = False
txtQuantity.Enabled = False
txtRemarks.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    rstMmforreceipt.Close
    Set rstMmforreceipt = Nothing
    
    rstReceipt.Close
    Set rstReceipt = Nothing
    
    rstDatatoflex.Close
    Set rstDatatoflex = Nothing
End Sub

Private Sub txtQuantity_Change()
If Not IsNumeric(txtQuantity) Then
    txtQuantity = ""
End If

End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRemarks.SetFocus
End If
End Sub


Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub

