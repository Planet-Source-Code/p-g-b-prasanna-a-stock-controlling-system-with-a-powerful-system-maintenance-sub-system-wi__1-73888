VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMasterInformation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master Information Entry"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   Icon            =   "frmMasterInformation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5400
      MouseIcon       =   "frmMasterInformation.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   450
         Left            =   120
         Picture         =   "frmMasterInformation.frx":015E
         Top             =   100
         Width           =   435
      End
      Begin VB.Image Image2 
         Height          =   360
         Left            =   6600
         Picture         =   "frmMasterInformation.frx":0BF0
         ToolTipText     =   "Application Help"
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Master Information Of The System Here."
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
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   4005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5650
      Left            =   120
      TabIndex        =   0
      Top             =   610
      Width           =   7095
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3550
         Left            =   40
         LinkTimeout     =   80
         ScaleHeight     =   3555
         ScaleWidth      =   6975
         TabIndex        =   31
         Top             =   2040
         Width           =   6975
         Begin VB.CommandButton cmdReturntofrecords 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Return to F&ull Records"
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
            Left            =   4200
            MouseIcon       =   "frmMasterInformation.frx":135A
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CommandButton cmdFind 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5280
            MouseIcon       =   "frmMasterInformation.frx":14AC
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdLast 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2040
            MouseIcon       =   "frmMasterInformation.frx":15FE
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   1200
            Width           =   2055
         End
         Begin VB.CommandButton cmdFirst 
            BackColor       =   &H00FFFFFF&
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
            MouseIcon       =   "frmMasterInformation.frx":1750
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CommandButton cmdPrevious 
            BackColor       =   &H00FFFFFF&
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
            MouseIcon       =   "frmMasterInformation.frx":18A2
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdNext 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1680
            MouseIcon       =   "frmMasterInformation.frx":19F4
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5280
            MouseIcon       =   "frmMasterInformation.frx":1B46
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3480
            MouseIcon       =   "frmMasterInformation.frx":1C98
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3480
            MouseIcon       =   "frmMasterInformation.frx":1DEA
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   720
            Width           =   1695
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1680
            MouseIcon       =   "frmMasterInformation.frx":1F3C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFFFF&
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
            MouseIcon       =   "frmMasterInformation.frx":208E
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid MSFGMaster 
            Height          =   1815
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   3201
            _Version        =   393216
            BackColor       =   16642796
            BackColorFixed  =   14737632
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
      Begin TabDlg.SSTab SSTab 
         Height          =   1935
         Left            =   40
         TabIndex        =   19
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         _Version        =   393216
         Style           =   1
         Tab             =   2
         TabHeight       =   520
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Course Info"
         TabPicture(0)   =   "frmMasterInformation.frx":21E0
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame5"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Module Info"
         TabPicture(1)   =   "frmMasterInformation.frx":21FC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Category Info"
         TabPicture(2)   =   "frmMasterInformation.frx":2218
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Height          =   1550
            Left            =   -75000
            TabIndex        =   26
            Top             =   340
            Width           =   6950
            Begin VB.TextBox txtCourse 
               Appearance      =   0  'Flat
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
               Left            =   1920
               TabIndex        =   1
               Top             =   240
               Width           =   3855
            End
            Begin VB.TextBox txtCourseDesctiption 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   2
               Top             =   720
               Width           =   4575
            End
            Begin VB.Label Label1 
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
               Height          =   375
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Course Desctiption"
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
               Top             =   720
               Width           =   1620
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   1550
            Left            =   -75000
            TabIndex        =   23
            Top             =   340
            Width           =   6950
            Begin VB.TextBox txtModuleDescription 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   1920
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   720
               Width           =   4575
            End
            Begin VB.TextBox txtModule 
               Appearance      =   0  'Flat
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
               Left            =   1920
               TabIndex        =   3
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Module Desctiption"
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
               Width           =   1605
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Module"
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
               Left            =   120
               TabIndex        =   24
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   1550
            Left            =   0
            TabIndex        =   20
            Top             =   340
            Width           =   6950
            Begin VB.TextBox txtCategoryDescription 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2040
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox txtCategory 
               Appearance      =   0  'Flat
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
               Left            =   2040
               TabIndex        =   5
               Top             =   240
               Width           =   3735
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Category  Desctiption"
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
               Width           =   1860
            End
            Begin VB.Label Label9 
               BackColor       =   &H00FFFFFF&
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
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   240
               Width           =   975
            End
         End
      End
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   4680
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4680
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4800
      Picture         =   "frmMasterInformation.frx":2234
      Top             =   6300
      Width           =   510
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4680
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4680
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "frmMasterInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCoursemaster As ADODB.Recordset
Dim rstModulemaster As ADODB.Recordset
Dim rstCategorymaster As ADODB.Recordset
'Dim inttabselect As Integer
Dim blnadd_Click As Boolean
Dim blnedit_Click As Boolean
Dim validity As Boolean
Dim duplicate As Boolean

Private Sub cmdAdd_Click()
If Check_For_Privilege(1) = True Then: Exit Sub
Enable_Controls
    blnadd_Click = True
    blnedit_Click = False
If inttabselect = 0 Then
    Empty_Course_Fields
    txtCourse.SetFocus
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
ElseIf inttabselect = 1 Then
    Empty_Module_Fields
    txtModule.SetFocus
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
ElseIf inttabselect = 2 Then
    Empty_Category_Fields
    txtCategory.SetFocus
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
If inttabselect = 0 Then
    rstCoursemaster.CancelUpdate
    Empty_Course_Fields
    Rec_Set_Info
ElseIf inttabselect = 1 Then
    rstModulemaster.CancelUpdate
    Empty_Module_Fields
    Rec_Set_Info
ElseIf inttabselect = 2 Then
    rstCategorymaster.CancelUpdate
    Empty_Category_Fields
    Rec_Set_Info
End If
Disable_Controls
On Error Resume Next
cmdAdd.SetFocus
End Sub

Private Sub cmdDelete_Click()
If Check_For_Privilege(3) = True Then: Exit Sub

If MsgBox("Are you sure you want to delete the current Record ?", vbQuestion + vbYesNo) = vbYes Then
    If inttabselect = 0 Then
        On Error Resume Next
        rstCoursemaster.Delete
        rstCoursemaster.MoveNext
            If rstCoursemaster.EOF Then
                rstCoursemaster.MoveLast
            End If
           Empty_Course_Fields
        Rec_Set_Info
    ElseIf inttabselect = 1 Then
        On Error Resume Next
        rstModulemaster.Delete
        rstModulemaster.MoveNext
            If rstModulemaster.EOF Then
                rstModulemaster.MoveLast
            End If
           Empty_Module_Fields
        Rec_Set_Info
    ElseIf inttabselect = 2 Then
        On Error Resume Next
        rstCategorymaster.Delete
        rstCategorymaster.MoveNext
            If rstCategorymaster.EOF Then
               rstCategorymaster.MoveLast
            End If
        Empty_Category_Fields
        Rec_Set_Info
    End If
Else
Exit Sub
End If
End Sub

Private Sub cmdEdit_Click()
If Check_For_Privilege(2) = True Then: Exit Sub

Enable_Controls
    blnedit_Click = True
    blnadd_Click = False
If inttabselect = 0 Then
    'Empty_Course_Fields
    txtCourse.SetFocus
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
ElseIf inttabselect = 1 Then
    'Empty_Module_Fields
    txtModule.SetFocus
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
ElseIf inttabselect = 2 Then
    'Empty_Category_Fields
    txtCategory.SetFocus
    Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, False
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
frmFindMasterInfo.Show 1
End Sub

Private Sub cmdFirst_Click()
If inttabselect = 0 Then
    If rstCoursemaster.BOF = False Then
        rstCoursemaster.MoveFirst
        Rec_Set_Info
    'ElseIf rstcoursemaster.EOF Then
        'rstcoursemaster.MoveLast
        'Rec_Set_Info
        MsgBox "You are on the First Record.", vbInformation
    End If
ElseIf inttabselect = 1 Then
    If rstModulemaster.BOF = False Then
        rstModulemaster.MoveFirst
        Rec_Set_Info
    'ElseIf rstmodulemaster.EOF Then
        'rstmodulemaster.MoveLast
        'Rec_Set_Info
        MsgBox "You are on the First Record.", vbInformation
    End If
    
ElseIf inttabselect = 2 Then
    If rstCategorymaster.BOF = False Then
        rstCategorymaster.MoveFirst
        Rec_Set_Info
    'ElseIf rstcategorymaster.BOF Then
       ' rstcategorymaster.MoveFirst
        Rec_Set_Info
        MsgBox "You are on the First Record.", vbInformation
    End If
End If

End Sub

Private Sub cmdLast_Click()
If inttabselect = 0 Then
    If rstCoursemaster.EOF = False Then
        rstCoursemaster.MoveLast
        Rec_Set_Info
    'ElseIf rstcoursemaster.EOF Then
        'rstcoursemaster.MoveLast
        'Rec_Set_Info
        MsgBox "You are on the Last Record.", vbInformation
    End If
ElseIf inttabselect = 1 Then
    If rstModulemaster.EOF = False Then
        rstModulemaster.MoveLast
        Rec_Set_Info
    'ElseIf rstmodulemaster.EOF Then
        'rstmodulemaster.MoveLast
        'Rec_Set_Info
        MsgBox "You are on the Last Record.", vbInformation
    End If
    
ElseIf inttabselect = 2 Then
    If rstCategorymaster.EOF = False Then
        rstCategorymaster.MoveLast
        Rec_Set_Info
    'ElseIf rstcategorymaster.EOF Then
       ' rstcategorymaster.MoveLast
        Rec_Set_Info
        MsgBox "You are on the Last Record.", vbInformation
    End If
End If

End Sub

Private Sub cmdNext_Click()
If inttabselect = 0 Then
    
    If rstCoursemaster.EOF = False Then
        rstCoursemaster.MoveNext
        Rec_Set_Info
    End If
    If rstCoursemaster.EOF Then
        rstCoursemaster.MoveLast
        Rec_Set_Info
        MsgBox "You are on the Last Record.", vbInformation
    End If
ElseIf inttabselect = 1 Then
    
    If rstModulemaster.EOF = False Then
        rstModulemaster.MoveNext
        Rec_Set_Info
    End If
    If rstModulemaster.EOF Then
        rstModulemaster.MoveLast
        Rec_Set_Info
        MsgBox "You are on the Last Record.", vbInformation
    End If
    
ElseIf inttabselect = 2 Then
    
    If rstCategorymaster.EOF = False Then
        rstCategorymaster.MoveNext
        Rec_Set_Info
    End If
    If rstCategorymaster.EOF Then
        rstCategorymaster.MoveLast
        Rec_Set_Info
        MsgBox "You are on the Last Record.", vbInformation
    End If
End If
End Sub

Private Sub cmdPrevious_Click()
If inttabselect = 0 Then
    If rstCoursemaster.BOF = False Then
        rstCoursemaster.MovePrevious
        Rec_Set_Info
    ElseIf rstCoursemaster.BOF Then
        rstCoursemaster.MoveFirst
        Rec_Set_Info
        MsgBox "You are on the first Record.", vbInformation
    End If
ElseIf inttabselect = 1 Then
    If rstModulemaster.BOF = False Then
        rstModulemaster.MovePrevious
        Rec_Set_Info
    ElseIf rstModulemaster.BOF Then
        rstModulemaster.MoveFirst
        Rec_Set_Info
        MsgBox "You are on the First Record.", vbInformation
    End If
    
ElseIf inttabselect = 2 Then
    If rstCategorymaster.BOF = False Then
        rstCategorymaster.MovePrevious
        Rec_Set_Info
    ElseIf rstCategorymaster.BOF Then
        rstCategorymaster.MoveFirst
        Rec_Set_Info
        MsgBox "You are on the first Record.", vbInformation
    End If
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
            If inttabselect = 0 Then
                rstCoursemaster.AddNew
                    Save_Data
            ElseIf inttabselect = 1 Then
                rstModulemaster.AddNew
                    Save_Data
            ElseIf inttabselect = 2 Then
                rstCategorymaster.AddNew
                    Save_Data
            End If
         End If
         If blnedit_Click = True Then
            blnedit_Click = False
            If inttabselect = 0 Then
                Save_Data
            ElseIf inttabselect = 1 Then
                Save_Data
            ElseIf inttabselect = 2 Then
                Save_Data
            End If
        End If
    End If
If validity = True Then
Disable_Controls
End If
End Sub

Private Sub Command1_Click()
MSFGMaster.Rows = 0
MSFGMaster.Cols = 0
End Sub

Private Sub Form_Load()
Set rstCoursemaster = New ADODB.Recordset
    rstCoursemaster.CursorLocation = adUseClient
    rstCoursemaster.Open "SELECT * FROM COURSE ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
    'MsgBox rstcoursemaster.RecordCount & "Course"
Set rstModulemaster = New ADODB.Recordset
    rstModulemaster.CursorLocation = adUseClient
    rstModulemaster.Open "SELECT * FROM MODULER ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
    'MsgBox rstcoursemaster.RecordCount & " MODULER"
Set rstCategorymaster = New ADODB.Recordset
    rstCategorymaster.CursorLocation = adUseClient
    rstCategorymaster.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockOptimistic
    'MsgBox rstcategorymaster.RecordCount & "Category"
  
    If find_check = True Then
        find_check = False
        SSTab.Tab = 1
        inttabselect = 1
    Else
        SSTab.Tab = 0
        'inttabselect = 0
    End If
    Disable_Controls
    Rec_Set_Info
    Me.Picture = frmStyle.Picture
    End Sub

Private Sub Form_Unload(Cancel As Integer)
blnfind_status = False
update_find = False
    On Error Resume Next
    rstCoursemaster.Close
    Set rstCoursemaster = Nothing
    
    rstModulemaster.Close
    Set rstModulemaster = Nothing
    
    rstCategorymaster.Close
    Set rstCategorymaster = Nothing
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
If SSTab.Tab = 0 Then
    inttabselect = 0
    Rec_Set_Info
ElseIf SSTab.Tab = 1 Then
    inttabselect = 1
     Rec_Set_Info
ElseIf SSTab.Tab = 2 Then
    inttabselect = 2
     Rec_Set_Info
End If
Disable_Controls
End Sub
Public Sub Rec_Set_Info()
If inttabselect = 0 Then
   cmdFind.Enabled = False
   cmdReturntofrecords.Enabled = False
    If rstCoursemaster.RecordCount > 0 Then
        Inforfield
        'Format_Flex
        'Add_Flex
        Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    Else
        'MsgBox "No Records available!!", vbExclamation
        Button_Record_Not_Exist_Mode Me
    End If
ElseIf inttabselect = 1 Then
    
    cmdFind.Enabled = True
    'If blnfind_status = True Then
        
        'cmdReturntofrecords.Enabled = True
    'Else
    'cmdReturntofrecords.Enabled = False
    'End If
    If rstModulemaster.RecordCount > 0 Then
        
        Inforfield
        'Format_Flex
        'Add_Flex
        Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    Else
        'MsgBox "No records available!!", vbExclamation
        Button_Record_Not_Exist_Mode Me
    End If
ElseIf inttabselect = 2 Then
    cmdFind.Enabled = False
    cmdReturntofrecords.Enabled = False
    If rstCategorymaster.RecordCount > 0 Then
        Inforfield
        Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    Else
        'MsgBox "No recoprds available!!", vbExclamation
        Button_Record_Not_Exist_Mode Me
    End If
End If
    If blnfind_status = True Then
        'blnfind_status = False
        Exit Sub
    End If
    
  Format_Flex
  Add_Flex
End Sub

Public Sub Inforfield()
On Error Resume Next
If inttabselect = 0 Then
    txtCourse.Text = rstCoursemaster("COURSE")
    txtCourseDesctiption.Text = rstCoursemaster("COURSE_DES")
    Me.Caption = "Master Information Entry   -  Record Count: " & rstCoursemaster.RecordCount
ElseIf inttabselect = 1 Then
    txtModule.Text = rstModulemaster("MODULER")
    txtModuleDescription = rstModulemaster("MODULE_DES")
    Me.Caption = "Master Information Entry   -  Record Count: " & rstModulemaster.RecordCount
ElseIf inttabselect = 2 Then
    txtCategory.Text = rstCategorymaster("CATEGORY")
    txtCategoryDescription.Text = rstCategorymaster("CATEGORY_DES")
    Me.Caption = "Master Information Entry   -  Record Count: " & rstCategorymaster.RecordCount
End If
'Format_Flex
End Sub
Public Sub Empty_Course_Fields()
txtCourse.Text = ""
txtCourseDesctiption.Text = ""
End Sub
Public Sub Empty_Module_Fields()
txtModule.Text = ""
txtModuleDescription.Text = ""
End Sub
Public Sub Empty_Category_Fields()
txtCategory.Text = ""
txtCategoryDescription.Text = ""
End Sub

Public Sub Save_Data()
On Error GoTo Err
If inttabselect = 0 Then
    rstCoursemaster("COURSE") = UCase(txtCourse.Text)
    rstCoursemaster("COURSE_DES") = txtCourseDesctiption.Text
    rstCoursemaster.Update
    Me.Caption = "Master Information Entry   -  Record Count: " & rstCoursemaster.RecordCount
ElseIf inttabselect = 1 Then
    rstModulemaster("MODULER") = UCase(txtModule.Text)
    rstModulemaster("MODULE_DES") = txtModuleDescription.Text
    rstModulemaster.Update
    Me.Caption = "Master Information Entry   -  Record Count: " & rstModulemaster.RecordCount
ElseIf inttabselect = 2 Then
    rstCategorymaster("CATEGORY") = UCase(txtCategory.Text)
    rstCategorymaster("CATEGORY_DES") = txtCategoryDescription.Text
    rstCategorymaster.Update
    Me.Caption = "Master Information Entry   -  Record Count: " & rstCategorymaster.RecordCount
End If
Button_Add_Edit_Save_Cancle_RecordExist_Mode Me, True
    If blnfind_status = True Then
        update_find = True
        Find_for_Details
    Else
    
    Format_Flex
    Add_Flex
    On Error Resume Next
    cmdAdd.SetFocus
    End If
Exit Sub

Err:
    MsgBox Err.Description & " - " & Err.Number & ".", vbCritical
    Form_Load
End Sub

Public Sub Data_Validity()
If inttabselect = 0 Then
     If txtCourse.Text = "" Then
         MsgBox "Course Field Cannot be blank!!!", vbExclamation
         validity = False
         txtCourse.SetFocus
         Exit Sub
     Else
         validity = True
     End If
     Duplicate_Data
        If duplicate = True Then
            MsgBox "Course is already exist!!", vbExclamation
            txtCourse.SetFocus
            SendKeys "{Home}+{End}"
            validity = False
            Exit Sub
        ElseIf duplicate = False Then
            validity = True
        End If
 ElseIf inttabselect = 1 Then
    'If txtModule.Text = "" Then
    '    MsgBox "Module Field cannot be blank!!!", vbExclamation
    '        txtModule.SetFocus
    '    validity = False
    '    Exit Sub
    'Else
    '    validity = True
    'End If
    'If txtModuleDescription.Text = "" Then
    '    MsgBox "Module Description Field cannot be blank!!!", vbExclamation
     '       txtModuleDescription.SetFocus
    '    validity = False
    '    Exit Sub
    'Else
    '    validity = True
    'End If
    If txtModule.Text = "" Or txtModuleDescription.Text = "" Then
        MsgBox "Module Field and Module Description Field cannot be blank!!!", vbExclamation
            If txtModule.Text = "" And txtModuleDescription.Text <> "" Then
                txtModule.SetFocus
            ElseIf txtModuleDescription.Text = "" And txtModule.Text <> "" Then
                txtModuleDescription.SetFocus
            ElseIf txtModule.Text = "" And txtModuleDescription.Text = "" Then
                txtModule.SetFocus
            End If
         validity = False
         Exit Sub
     Else
        validity = True
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
     
 ElseIf inttabselect = 2 Then
    If txtCategory.Text = "" Then
        MsgBox "Category Field Cannot be blank!!!", vbExclamation
        txtCategory.SetFocus
         validity = False
         Exit Sub
     Else
        validity = True
     End If
    Duplicate_Data
        If duplicate = True Then
            MsgBox "Category is already exist!!", vbExclamation
                txtCategory.SetFocus
                SendKeys "{Home}+{End}"
                validity = False
        ElseIf duplicate = False Then
            validity = True
        End If
     
End If
End Sub

Public Sub Format_Flex()
With MSFGMaster
.Clear
.Cols = 2
.Rows = 1
.CellAlignment = flexAlignLeftCenter
If inttabselect = 0 Then
    .ColWidth(0) = 1500
    .TextMatrix(0, 0) = "Course"
    .ColWidth(1) = 5000
    .TextMatrix(0, 1) = "Description"
ElseIf inttabselect = 1 Then
   .ColWidth(0) = 2000
    .TextMatrix(0, 0) = "Module"
    .ColWidth(1) = 4500
    .TextMatrix(0, 1) = "Description"
ElseIf inttabselect = 2 Then
    .ColWidth(0) = 2000
    .TextMatrix(0, 0) = "Category"
    .ColWidth(1) = 4500
    .TextMatrix(0, 1) = "Description"
End If
End With

End Sub

Public Sub Add_Flex()
Dim i As Integer
Dim rcount As Integer
If inttabselect = 0 Then
    Set rstcourseflex = New ADODB.Recordset
        rstcourseflex.CursorLocation = adUseClient
        rstcourseflex.Open "SELECT * FROM COURSE ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly

    If rstcourseflex.RecordCount > 0 Then
        rcount = rstcourseflex.RecordCount
        MSFGMaster.Rows = rcount + 1
        i = 1
        With MSFGMaster
            Do While Not rstcourseflex.EOF
            .Row = i
            .Col = 0: .Text = rstcourseflex("COURSE")
            .Row = i
            .Col = 1: .Text = rstcourseflex("COURSE_DES")
       i = i + 1
       rstcourseflex.MoveNext
       Loop
       End With
    'Else
       ' Format_Flex
    End If
ElseIf inttabselect = 1 Then

Set rstmoduleflex = New ADODB.Recordset
    rstmoduleflex.CursorLocation = adUseClient
    rstmoduleflex.Open "SELECT * FROM MODULER ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly
    'MsgBox rstmoduleflex.RecordCount
    If rstmoduleflex.RecordCount > 0 Then
        rcount = rstmoduleflex.RecordCount
        MSFGMaster.Rows = rcount + 1
        i = 1
        With MSFGMaster
            Do While Not rstmoduleflex.EOF
            .Row = i
            .Col = 0: .Text = rstmoduleflex("MODULER")
            .Row = i
            .Col = 1: .Text = rstmoduleflex("MODULE_DES")
       i = i + 1
       rstmoduleflex.MoveNext
       Loop
       End With
    'Else
        'Format_Flex
    End If
ElseIf inttabselect = 2 Then
Set rstcategoryflex = New ADODB.Recordset
    rstcategoryflex.CursorLocation = adUseClient
    rstcategoryflex.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly

    
    If rstcategoryflex.RecordCount > 0 Then
        rcount = rstcategoryflex.RecordCount
        MSFGMaster.Rows = rcount + 1
        i = 1
        With MSFGMaster
            Do While Not rstcategoryflex.EOF
            .Row = i
            .Col = 0: .Text = rstcategoryflex("CATEGORY")
            .Row = i
            .Col = 1: .Text = rstcategoryflex("CATEGORY_DES")
       i = i + 1
       rstcategoryflex.MoveNext
       Loop
       End With
    'Else
        'Format_Flex
    End If
End If
End Sub

Public Sub Duplicate_Data()
        If inttabselect = 0 Then
            Set rstvalidatecourse = New ADODB.Recordset
            rstvalidatecourse.CursorLocation = adUseClient
            rstvalidatecourse.Open "SELECT * FROM COURSE WHERE COURSE = '" & txtCourse.Text & "'", dbcon, adOpenStatic, adLockReadOnly
                If rstvalidatecourse.RecordCount > 0 Then
                    If blnadd_Click = True Then
                        'blnadd_Click = False
                        duplicate = True
                    ElseIf blnedit_Click = True Then
                        If UCase(rstCoursemaster("COURSE")) = UCase(txtCourse.Text) Then
                           'blnedit_Click = False
                           duplicate = False
                        Else
                        duplicate = True
                        End If
                    End If
                Else
                duplicate = False
                End If
                rstvalidatecourse.Close
                Set rstvalidatecourse = Nothing
        ElseIf inttabselect = 1 Then
            Set rstvalidatemoduler = New ADODB.Recordset
            rstvalidatemoduler.CursorLocation = adUseClient
            rstvalidatemoduler.Open "SELECT * FROM MODULER WHERE MODULER = '" & txtModule.Text & "'", dbcon, adOpenStatic, adLockReadOnly
                If rstvalidatemoduler.RecordCount > 0 Then
                    If blnadd_Click = True Then
                        'blnadd_Click = False
                        duplicate = True
                    ElseIf blnedit_Click = True Then
                        If UCase(rstModulemaster("MODULER")) = UCase(txtModule.Text) Then
                           'blnedit_Click = False
                           duplicate = False
                        Else
                        duplicate = True
                        End If
                    End If
                Else
                duplicate = False
                End If
                rstvalidatemoduler.Close
                Set rstvalidatemoduler = Nothing
            
        ElseIf inttabselect = 2 Then
             Set rstvalidatecategory = New ADODB.Recordset
            rstvalidatecategory.CursorLocation = adUseClient
            rstvalidatecategory.Open "SELECT * FROM CATEGORY WHERE CATEGORY = '" & txtCategory.Text & "'", dbcon, adOpenStatic, adLockReadOnly
                If rstvalidatecategory.RecordCount > 0 Then
                    If blnadd_Click = True Then
                        'blnadd_Click = False
                        duplicate = True
                    ElseIf blnedit_Click = True Then
                        If UCase(rstCategorymaster("CATEGORY")) = UCase(txtCategory.Text) Then
                           'blnedit_Click = False
                           duplicate = False
                        Else
                        duplicate = True
                        End If
                    End If
                Else
                duplicate = False
                End If
                rstvalidatecategory.Close
                Set rstvalidatecategory = Nothing
        End If
        
End Sub
Public Sub Find_for_Details()
    rstModulemaster.Close
    Set rstModulemaster = Nothing
    'rstmoduleflex.Close
    'Set rstmoduleflex = Nothing
    'MsgBox Find_Val
    Set rstModulemaster = New ADODB.Recordset
        rstModulemaster.Open "SELECT* FROM MODULER WHERE[MODULER] = '" & Find_Val & "'", dbcon, adOpenStatic, adLockOptimistic
        If rstModulemaster.RecordCount > 0 Then
            Inforfield
            Format_Flex
            Add_Flex_Find_Data
            Unload frmFindMasterInfo
            cmdReturntofrecords.Enabled = True
            SSTab.TabEnabled(0) = False
            SSTab.TabEnabled(2) = False
         Else
            If update_find = True Then
                
                update_find = False
                Load_Initiate
                Exit Sub
            Else
                
            MsgBox "No records found for this search...!", vbExclamation
            On Error Resume Next
            frmFindMasterInfo.txtFindValue.SetFocus
            SendKeys "{Home}+{End}"
            Load_Initiate
            Exit Sub
            End If
        End If
End Sub

Public Sub Add_Flex_Find_Data()
Dim i As Integer
Dim rcount As Integer
    Set rstmoduleflexforfind = New ADODB.Recordset
        rstmoduleflexforfind.Open "SELECT* FROM MODULER WHERE[MODULER] = '" & Find_Val & "'", dbcon, adOpenStatic, adLockReadOnly

  If rstmoduleflexforfind.RecordCount > 0 Then
        rcount = rstmoduleflexforfind.RecordCount
        MSFGMaster.Rows = rcount + 1
        i = 1
        With MSFGMaster
            Do While Not rstmoduleflexforfind.EOF
            .Row = i
            .Col = 0: .Text = rstmoduleflexforfind("MODULER")
            .Row = i
            .Col = 1: .Text = rstmoduleflexforfind("MODULE_DES")
       i = i + 1
       rstmoduleflexforfind.MoveNext
       Loop
       End With
    Else
        Format_Flex
    End If
    
End Sub
Public Sub Load_Initiate()
blnfind_status = False
update_find = False
Form_Load
cmdReturntofrecords.Enabled = False
SSTab.TabEnabled(0) = True
SSTab.TabEnabled(2) = True
End Sub

Public Sub Enable_Controls()
If inttabselect = 0 Then
    txtCourse.Enabled = True
    txtCourseDesctiption.Enabled = True
ElseIf inttabselect = 1 Then
    txtModule.Enabled = True
    txtModuleDescription.Enabled = True
ElseIf inttabselect = 2 Then
    txtCategory.Enabled = True
    txtCategoryDescription.Enabled = True
End If
End Sub
Public Sub Disable_Controls()
If inttabselect = 0 Then
    txtCourse.Enabled = False
    txtCourseDesctiption.Enabled = False
ElseIf inttabselect = 1 Then
    txtModule.Enabled = False
    txtModuleDescription.Enabled = False
ElseIf inttabselect = 2 Then
    txtCategory.Enabled = False
    txtCategoryDescription.Enabled = False
End If
End Sub
Private Sub txtCategory_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtCategoryDescription.SetFocus
End If
End Sub

Private Sub txtCategoryDescription_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub txtCourse_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtCourseDesctiption.SetFocus
End If
End Sub

Private Sub txtCourseDesctiption_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub
Private Sub txtModule_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    txtModuleDescription.SetFocus
End If
End Sub

Private Sub txtModuleDescription_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub
