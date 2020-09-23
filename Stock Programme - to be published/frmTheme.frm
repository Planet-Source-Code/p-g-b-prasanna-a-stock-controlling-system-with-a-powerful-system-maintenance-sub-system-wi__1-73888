VERSION 5.00
Begin VB.Form frmTheme 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Theme Settings"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   3960
   Icon            =   "frmTheme.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Theme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   3495
         TabIndex        =   8
         Top             =   240
         Width           =   3495
         Begin VB.CommandButton cmdOk 
            Cancel          =   -1  'True
            Caption         =   "OK"
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
            Left            =   2520
            MouseIcon       =   "frmTheme.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   2520
            Width           =   975
         End
         Begin VB.CommandButton cmdApplyTheme 
            Caption         =   "&Apply Theme"
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
            Left            =   960
            MouseIcon       =   "frmTheme.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   2520
            Width           =   1455
         End
         Begin VB.OptionButton optTheme5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Theme 5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optTheme4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Theme 4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton optTheme3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Theme 3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optTheme2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Theme 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optTheme1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Theme 1 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optDefaultTheme 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Default Theme"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   0
            Width           =   1695
         End
         Begin VB.Line Line1 
            X1              =   980
            X2              =   3500
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Image imgSample 
            Height          =   615
            Left            =   120
            Stretch         =   -1  'True
            Top             =   2280
            Width           =   735
         End
         Begin VB.Image Image6 
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmTheme.frx":02B0
            MousePointer    =   99  'Custom
            Picture         =   "frmTheme.frx":0402
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Image Image5 
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmTheme.frx":1204
            MousePointer    =   99  'Custom
            Picture         =   "frmTheme.frx":1356
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Image Image4 
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmTheme.frx":2158
            MousePointer    =   99  'Custom
            Picture         =   "frmTheme.frx":22AA
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Image Image3 
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmTheme.frx":30AC
            MousePointer    =   99  'Custom
            Picture         =   "frmTheme.frx":31FE
            Top             =   720
            Width           =   1455
         End
         Begin VB.Image Image2 
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmTheme.frx":4000
            MousePointer    =   99  'Custom
            Picture         =   "frmTheme.frx":4152
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   1920
            MouseIcon       =   "frmTheme.frx":4F54
            MousePointer    =   99  'Custom
            Picture         =   "frmTheme.frx":50A6
            Top             =   0
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApplyTheme_Click()
On Error Resume Next
Set rstthemeset = New ADODB.Recordset
    rstthemeset.CursorLocation = adUseClient
        rstthemeset.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & User & "'", dbcon, adOpenStatic, adLockOptimistic
            If optDefaultTheme.Value = True Then
                rstthemeset("THEME_SET") = "1"
                rstthemeset.Update
            ElseIf optTheme1.Value = True Then
                rstthemeset("THEME_SET") = "2"
                rstthemeset.Update
            ElseIf optTheme2.Value = True Then
                rstthemeset("THEME_SET") = "3"
                rstthemeset.Update
            ElseIf optTheme3.Value = True Then
                rstthemeset("THEME_SET") = "4"
                rstthemeset.Update
            ElseIf optTheme4.Value = True Then
                rstthemeset("THEME_SET") = "5"
                rstthemeset.Update
            ElseIf optTheme5.Value = True Then
                rstthemeset("THEME_SET") = "6"
                rstthemeset.Update
            End If
   inttheme = rstthemeset("THEME_SET")
   Unload frmStyle
   'Theme_Settings
   Load frmStyle
   blntheme_apply = True
   frmMain.Resolution_Set
   frmMain.Picture = frmStyle.Picture
   Me.Picture = frmStyle.Picture
   
   rstthemeset.Close
   Set rstthemeset = Nothing
   
    End Sub
Private Sub cmdOk_Click()
cmdApplyTheme_Click
Unload Me
End Sub

Private Sub Form_Load()
Select Case inttheme
    Case 1: optDefaultTheme.Value = True: imgSample.Picture = Image1.Picture
    Case 2: optTheme1.Value = True: imgSample.Picture = Image2.Picture
    Case 3: optTheme2.Value = True: imgSample.Picture = Image3.Picture
    Case 4: optTheme3.Value = True: imgSample.Picture = Image4.Picture
    Case 5: optTheme4.Value = True: imgSample.Picture = Image5.Picture
    Case 6: optTheme5.Value = True: imgSample.Picture = Image6.Picture
    Case Else: optDefaultTheme.Value = True: imgSample.Picture = Image1.Picture
End Select
Me.Picture = frmStyle.Picture
End Sub

Private Sub Image1_Click()
optDefaultTheme.Value = True
imgSample.Picture = Image1.Picture
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image1_Click
End Sub

Private Sub Image2_Click()
optTheme1.Value = True
imgSample.Picture = Image2.Picture
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image2_Click
End Sub

Private Sub Image3_Click()
optTheme2.Value = True
imgSample.Picture = Image3.Picture
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image3_Click
End Sub

Private Sub Image4_Click()
optTheme3.Value = True
imgSample.Picture = Image4.Picture
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image4_Click
End Sub

Private Sub Image5_Click()
optTheme4.Value = True
imgSample.Picture = Image5.Picture
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image5_Click
End Sub

Private Sub Image6_Click()
optTheme5.Value = True
imgSample.Picture = Image6.Picture
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image6_Click
End Sub

Private Sub optDefaultTheme_Click()
imgSample.Picture = Image1.Picture
End Sub

Private Sub optTheme1_Click()
imgSample.Picture = Image2.Picture
End Sub

Private Sub optTheme2_Click()
imgSample.Picture = Image3.Picture
End Sub

Private Sub optTheme3_Click()
imgSample.Picture = Image4.Picture
End Sub

Private Sub optTheme4_Click()
imgSample.Picture = Image5.Picture
End Sub

Private Sub optTheme5_Click()
imgSample.Picture = Image6.Picture
End Sub
