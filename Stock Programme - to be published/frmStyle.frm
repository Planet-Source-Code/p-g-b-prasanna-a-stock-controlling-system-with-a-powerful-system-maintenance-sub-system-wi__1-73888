VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStyle 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "frmStyle.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStyle.frx":000C
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpNow 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   70844419
      CurrentDate     =   40026
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   22
      Left            =   -120
      TabIndex        =   22
      Text            =   "</assembly>"
      Top             =   5280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   21
      Left            =   -120
      TabIndex        =   21
      Text            =   "</dependency>"
      Top             =   5040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   20
      Left            =   -120
      TabIndex        =   20
      Text            =   "   </dependentAssembly>"
      Top             =   4800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   19
      Left            =   -120
      TabIndex        =   19
      Text            =   "     />"
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   18
      Left            =   -120
      TabIndex        =   18
      Text            =   "       language=""*"""
      Top             =   4320
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   17
      Left            =   -120
      TabIndex        =   17
      Text            =   "       publicKeyToken=""6595b64144ccf1df"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   16
      Left            =   -120
      TabIndex        =   16
      Text            =   "       processorArchitecture=""X86"""
      Top             =   3840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   15
      Left            =   -120
      TabIndex        =   15
      Text            =   "       version=""6.0.0.0"""
      Top             =   3600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   14
      Left            =   -120
      TabIndex        =   14
      Text            =   "       name=""Microsoft.Windows.Common-Controls"""
      Top             =   3360
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   13
      Left            =   -120
      TabIndex        =   13
      Text            =   "       type=""win32"""
      Top             =   3120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   12
      Left            =   -120
      TabIndex        =   12
      Text            =   "     <assemblyIdentity"
      Top             =   2880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   11
      Left            =   -120
      TabIndex        =   11
      Text            =   "   <dependentAssembly>"
      Top             =   2640
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   10
      Left            =   -120
      TabIndex        =   10
      Text            =   "<dependency>"
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   9
      Left            =   -120
      TabIndex        =   9
      Text            =   "<description>Stock Controlling System</description>"
      Top             =   2160
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   8
      Left            =   -120
      TabIndex        =   8
      Text            =   "/>"
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   7
      Left            =   -120
      TabIndex        =   7
      Text            =   "   type=""win32"""
      Top             =   1680
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   6
      Left            =   -120
      TabIndex        =   6
      Text            =   "   name=""Microsoft.VB6.VBnetStyles"""
      Top             =   1440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   5
      Left            =   -120
      TabIndex        =   5
      Text            =   "   processorArchitecture=""X86"""
      Top             =   1200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   4
      Left            =   -120
      TabIndex        =   4
      Text            =   "   version=""1.0.0.0"""
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   3
      Left            =   -120
      TabIndex        =   3
      Text            =   "<assemblyIdentity"
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   2
      Left            =   -120
      TabIndex        =   2
      Text            =   "manifestVersion=""1.0"">"
      Top             =   480
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   1
      Left            =   -120
      TabIndex        =   1
      Text            =   "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"""
      Top             =   240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Index           =   0
      Left            =   -120
      TabIndex        =   0
      Text            =   "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComDlg.CommonDialog cdDatabaseselect 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgTheme6 
      Height          =   435
      Left            =   3720
      Picture         =   "frmStyle.frx":4110
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   795
   End
   Begin VB.Image imgTheme5 
      Height          =   5595
      Left            =   0
      Picture         =   "frmStyle.frx":7994
      Top             =   960
      Width           =   11715
   End
   Begin VB.Image imgTheme4 
      Height          =   5595
      Left            =   5280
      Picture         =   "frmStyle.frx":B968
      Top             =   3120
      Width           =   11715
   End
   Begin VB.Image imgTheme006 
      Height          =   615
      Left            =   1320
      Picture         =   "frmStyle.frx":FA74
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   4050
   End
   Begin VB.Image imgTheme06 
      Height          =   615
      Left            =   1320
      Picture         =   "frmStyle.frx":12A56
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   7050
   End
   Begin VB.Image imgTheme005 
      Height          =   615
      Left            =   1320
      Picture         =   "frmStyle.frx":17978
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   4050
   End
   Begin VB.Image imgTheme05 
      Height          =   615
      Left            =   1320
      Picture         =   "frmStyle.frx":1A95A
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   7050
   End
   Begin VB.Image imgTheme004 
      Height          =   615
      Left            =   1320
      Picture         =   "frmStyle.frx":1F87C
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   4050
   End
   Begin VB.Image imgTheme04 
      Height          =   615
      Left            =   1320
      Picture         =   "frmStyle.frx":2781E
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   7050
   End
   Begin VB.Image imgTheme03 
      Height          =   255
      Left            =   6480
      Picture         =   "frmStyle.frx":35620
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1410
   End
   Begin VB.Image imgTheme003 
      Height          =   375
      Left            =   6720
      Picture         =   "frmStyle.frx":43422
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Image imgTheme02 
      Height          =   495
      Left            =   0
      Picture         =   "frmStyle.frx":4B3C4
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   330
   End
   Begin VB.Image imgTheme002 
      Height          =   615
      Left            =   2280
      Picture         =   "frmStyle.frx":591BE
      Top             =   2280
      Width           =   4050
   End
   Begin VB.Image imgTheme01 
      Height          =   360
      Left            =   6240
      Picture         =   "frmStyle.frx":61160
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1530
   End
   Begin VB.Image imgTheme3 
      Height          =   5595
      Left            =   2040
      Picture         =   "frmStyle.frx":72F22
      Top             =   1680
      Width           =   11715
   End
   Begin VB.Image imgTheme2 
      Height          =   795
      Left            =   2880
      Picture         =   "frmStyle.frx":7702E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1755
   End
   Begin VB.Image imgTheme1 
      Height          =   555
      Left            =   3000
      Picture         =   "frmStyle.frx":7B13A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   675
   End
   Begin VB.Image Image3 
      Height          =   1965
      Left            =   1080
      Picture         =   "frmStyle.frx":7F23E
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   -960
      Picture         =   "frmStyle.frx":D1C2C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   4080
   End
   Begin VB.Image imgTheme001 
      Height          =   360
      Left            =   6480
      Picture         =   "frmStyle.frx":14D646
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1530
   End
End
Attribute VB_Name = "frmStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
Unload Me
End Sub

Public Sub Form_Load()
Theme_Handle 1
dtpNow.Value = Date
End Sub

