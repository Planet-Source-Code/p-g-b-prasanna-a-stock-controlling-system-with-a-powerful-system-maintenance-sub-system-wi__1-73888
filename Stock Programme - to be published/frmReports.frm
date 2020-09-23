VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReports 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Reports"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancle 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   3345
      MouseIcon       =   "frmReports.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&View Report"
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
      Left            =   1800
      MouseIcon       =   "frmReports.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   8040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7860
      Left            =   120
      ScaleHeight     =   7860
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   120
      Width           =   4580
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report Parameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   4575
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   6015
            Left            =   120
            ScaleHeight     =   6015
            ScaleWidth      =   4335
            TabIndex        =   21
            Top             =   180
            Width           =   4335
            Begin VB.ComboBox cmbCategorySelect 
               BackColor       =   &H00FBF4F4&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1000
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   840
               Width           =   3175
            End
            Begin VB.OptionButton optBookIssue 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Books Issue"
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
               Left            =   1080
               TabIndex        =   3
               Top             =   1560
               Width           =   1455
            End
            Begin VB.OptionButton optBookReceipt 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Books Receipt"
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
               Left            =   1080
               TabIndex        =   7
               Top             =   3360
               Width           =   1575
            End
            Begin VB.OptionButton optCurrentStock 
               BackColor       =   &H00FFFFFF&
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
               Height          =   375
               Left            =   1080
               TabIndex        =   1
               Top             =   120
               Width           =   2175
            End
            Begin VB.CheckBox chkBIDateRange 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Date Range"
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
               Left            =   1440
               TabIndex        =   4
               Top             =   1920
               Width           =   1455
            End
            Begin VB.CheckBox chkBRDateRange 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Date Range"
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
               Left            =   1440
               TabIndex        =   8
               Top             =   3720
               Width           =   1455
            End
            Begin VB.OptionButton optBooksCheck 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Books Status"
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
               Left            =   1080
               TabIndex        =   11
               Top             =   5040
               Width           =   1935
            End
            Begin VB.TextBox txtNumber 
               BackColor       =   &H00FBF4F4&
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
               Left            =   2880
               TabIndex        =   14
               Top             =   5520
               Width           =   1300
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   1440
               TabIndex        =   22
               Top             =   5400
               Width           =   1335
               Begin VB.PictureBox Picture3 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   795
                  Left            =   40
                  ScaleHeight     =   795
                  ScaleWidth      =   1260
                  TabIndex        =   23
                  Top             =   -120
                  Width           =   1260
                  Begin VB.OptionButton optGreaterthan 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Greater than"
                     Enabled         =   0   'False
                     Height          =   195
                     Left            =   0
                     TabIndex        =   13
                     Top             =   480
                     Width           =   1215
                  End
                  Begin VB.OptionButton optLessthan 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Less than"
                     Enabled         =   0   'False
                     Height          =   255
                     Left            =   0
                     TabIndex        =   12
                     Top             =   120
                     Width           =   1335
                  End
               End
            End
            Begin MSComCtl2.DTPicker dtpBIFrom 
               Height          =   375
               Left            =   960
               TabIndex        =   5
               Top             =   2520
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   48365571
               CurrentDate     =   40026
            End
            Begin MSComCtl2.DTPicker dtpBITo 
               Height          =   375
               Left            =   2760
               TabIndex        =   6
               Top             =   2520
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   48365571
               CurrentDate     =   40026
            End
            Begin MSComCtl2.DTPicker dtpBRFrom 
               Height          =   375
               Left            =   1080
               TabIndex        =   9
               Top             =   4320
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   48365571
               CurrentDate     =   40026
            End
            Begin MSComCtl2.DTPicker dtpBRTo 
               Height          =   375
               Left            =   2760
               TabIndex        =   10
               Top             =   4320
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   48365571
               CurrentDate     =   40026
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Select Category"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2800
               TabIndex        =   30
               Top             =   480
               Width           =   1365
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "From:"
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
               Left            =   960
               TabIndex        =   27
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
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
               Left            =   2760
               TabIndex        =   26
               Top             =   2160
               Width           =   615
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   4200
               Y1              =   3120
               Y2              =   3120
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   4200
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "From:"
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
               Left            =   960
               TabIndex        =   25
               Top             =   3960
               Width           =   615
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "To:"
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
               Left            =   2760
               TabIndex        =   24
               Top             =   3960
               Width           =   615
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   4200
               Y1              =   4920
               Y2              =   4920
            End
            Begin VB.Image Image1 
               Height          =   750
               Left            =   0
               Picture         =   "frmReports.frx":02B0
               Top             =   3480
               Width           =   675
            End
            Begin VB.Image Image2 
               Height          =   705
               Left            =   0
               Picture         =   "frmReports.frx":1D82
               Top             =   1920
               Width           =   720
            End
            Begin VB.Image Image3 
               Height          =   840
               Left            =   -120
               Picture         =   "frmReports.frx":3834
               Top             =   5040
               Width           =   765
            End
            Begin VB.Image Image4 
               Height          =   600
               Left            =   0
               Picture         =   "frmReports.frx":5A96
               Top             =   360
               Width           =   675
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   0
         TabIndex        =   28
         Top             =   6405
         Width           =   4575
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   4335
            TabIndex        =   29
            Top             =   240
            Width           =   4335
            Begin VB.OptionButton optAll 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Non Grouped Report"
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
               TabIndex        =   17
               Top             =   720
               Width           =   2775
            End
            Begin VB.OptionButton OptCoursewise 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Course wise Report"
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
               TabIndex        =   15
               Top             =   0
               Width           =   2775
            End
            Begin VB.OptionButton optCatetorywise 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Catetory wise Report"
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
               TabIndex        =   16
               Top             =   360
               Width           =   2775
            End
            Begin VB.Image Image5 
               Height          =   855
               Left            =   0
               Picture         =   "frmReports.frx":7018
               Top             =   120
               Width           =   1245
            End
         End
      End
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   1680
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   1680
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   1680
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   1680
      Y1              =   8040
      Y2              =   8040
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dateselect As Boolean
Private Sub checkDateselect()
If chkBIDateRange.Value = 1 Or chkBRDateRange.Value = 1 Then
    dateselect = True
ElseIf chkBIDateRange.Value = 0 Or chkBRDateRange.Value = 0 Then
    dateselect = False
End If
End Sub

Private Sub chkBIDateRange_Click()
If chkBIDateRange.Value = 1 Then
    dtpBIFrom.Enabled = True
    dtpBITo.Enabled = True
ElseIf chkBIDateRange.Value = 0 Then
    dtpBIFrom.Enabled = False
    dtpBITo.Enabled = False
End If
End Sub

Private Sub chkBRDateRange_Click()
If chkBRDateRange.Value = 1 Then
    dtpBRFrom.Enabled = True
    dtpBRTo.Enabled = True
ElseIf chkBRDateRange.Value = 0 Then
    dtpBRFrom.Enabled = False
    dtpBRTo.Enabled = False
End If

End Sub



Private Sub cmbCategorySelect_Click()
optCurrentStock.Value = True
If cmbCategorySelect <> "----- All Stock -----" Then
    OptCoursewise.Enabled = False
    optCatetorywise.Enabled = False
    optAll.Enabled = True
    optAll.Value = True
Else
    OptCoursewise.Enabled = True
    optCatetorywise.Enabled = True
    optAll.Enabled = True
    OptCoursewise.Value = True
End If
End Sub

Private Sub cmdCancle_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
If Check_For_Privilege(4) = True Then: Exit Sub
If optBooksCheck.Value = True Then
    If txtNumber = "" Then
        MsgBox "Value is empty.", vbCritical
        txtNumber.SetFocus
        Exit Sub
    End If
End If
checkDateselect

'----------------------------------------------------------------------------------
' For Current Stock Reports
If optCurrentStock.Value = True And optAll.Value = True Then
    If cmbCategorySelect = "----- All Stock -----" Then
        Set rst_cur_stock_all = New ADODB.Recordset
            rst_cur_stock_all.CursorLocation = adUseClient
            rst_cur_stock_all.Open "SELECT * FROM CUR_STOCK ORDER BY COURSE", dbcon, adOpenStatic, adLockReadOnly
        Set drAlloncurstock.DataSource = rst_cur_stock_all
            drAlloncurstock.Show 1
    Else
        Set rst_cur_stock_category = New ADODB.Recordset
            rst_cur_stock_category.CursorLocation = adUseClient
            rst_cur_stock_category.Open "SELECT * FROM CUR_STOCK WHERE CATEGORY = '" & cmbCategorySelect & "' ORDER BY MODULER", dbcon, adOpenStatic, adLockReadOnly
        Set drAlloncurstock.DataSource = rst_cur_stock_category
            drAlloncurstock.Sections("Section4").Controls("Label1").Caption = "Current Stock for - " & cmbCategorySelect
            drAlloncurstock.Sections("Section5").Controls("Label7").Caption = "Total Books for - " & cmbCategorySelect
            drAlloncurstock.Show 1
    End If
ElseIf optCurrentStock.Value = True And OptCoursewise.Value = True Then
    If deGroupedreports.rscmdGroupedbyCorseCurStock_Grouping.State = adStateOpen Then
       deGroupedreports.rscmdGroupedbyCorseCurStock_Grouping.Close
    End If
    
    drGroupedbyCoruseCurStock.Show 1
    
ElseIf optCurrentStock.Value = True And optCatetorywise.Value = True Then
    If deGroupedreports.rscmdGroupedbyCategoryCurStock_Grouping.State = adStateOpen Then
       deGroupedreports.rscmdGroupedbyCategoryCurStock_Grouping.Close
    End If
    
    drGroupedbyCategoryCurStock.Show 1

' End of Current Stock Reports
'-------------------------------------------------------------------------------


'----------------------------------------------------------------------------------
' For Books Receipt Reports

ElseIf optBookReceipt.Value = True And optAll.Value = True Then
    If dateselect = False Then
        Set rst_receipt_all = New ADODB.Recordset
            rst_receipt_all.CursorLocation = adUseClient
            rst_receipt_all.Open "SELECT * FROM B_RECEIPT ORDER BY DATE_RECEIVED", dbcon, adOpenStatic, adLockReadOnly
        Set drAllreceipt.DataSource = rst_receipt_all
            drAllreceipt.Show 1
    ElseIf dateselect = True Then
        If dtpBRFrom.Value > dtpBRTo.Value Then
            MsgBox "Invalid Date Range.", vbExclamation
            dtpBRFrom.SetFocus
            Exit Sub
        End If
        Set rst_receipt_Date = New ADODB.Recordset
            rst_receipt_Date.CursorLocation = adUseClient
            rst_receipt_Date.Open "SELECT* FROM B_RECEIPT WHERE DATE_RECEIVED >= " & "#" & dtpBRFrom & "#" & " AND DATE_RECEIVED <= " & "#" & dtpBRTo & "#" & "ORDER BY DATE_RECEIVED", dbcon, adOpenStatic, adLockReadOnly
        Set drBookReceiptDate.DataSource = rst_receipt_Date
            drBookReceiptDate.Sections("section4").Controls("Label5").Caption = Format(dtpBRFrom, "dd/mm/yyyy")
            drBookReceiptDate.Sections("section4").Controls("Label8").Caption = Format(dtpBRTo, "dd/mm/yyyy")
            drBookReceiptDate.Show 1
    End If
    
 ElseIf optBookReceipt.Value = True And OptCoursewise.Value = True Then
    If dateselect = False Then
        If deGroupedreports.rscmdGroupedbyModuleBooksReceipt_Grouping.State = adStateOpen Then
           deGroupedreports.rscmdGroupedbyModuleBooksReceipt_Grouping.Close
        End If
        
    drGroupedbyModuleBooksReceive.Show 1
    
    ElseIf dateselect = True Then
            If dtpBRFrom.Value > dtpBRTo.Value Then
            MsgBox "Invalid Date Range.", vbExclamation
            dtpBRFrom.SetFocus
            Exit Sub
        End If

        If deGroupedreports.rscmdGroupedbyModuleBooksR_Grouping.State = adStateOpen Then
            deGroupedreports.rscmdGroupedbyModuleBooksR_Grouping.Close
        End If
            deGroupedreports.cmdGroupedbyModuleBooksR_Grouping dtpBRFrom, dtpBRTo
            Load drGroupedbyModuleBooksReceiptDate
            drGroupedbyModuleBooksReceiptDate.Sections("section4").Controls("Label13").Caption = Format(dtpBRFrom, "dd/mm/yyyy")
            drGroupedbyModuleBooksReceiptDate.Sections("section4").Controls("Label14").Caption = Format(dtpBRTo, "dd/mm/yyyy")

            drGroupedbyModuleBooksReceiptDate.Show 1
   End If
   
 ' End of Books Receipt Reports
 '-------------------------------------------------------------------------------
   
    
 '----------------------------------------------------------------------------------
 ' For Books Issue Reports
     
ElseIf optBookIssue.Value = True And optAll.Value = True Then
    If dateselect = False Then
        Set rst_issued_all = New ADODB.Recordset
            rst_issued_all.CursorLocation = adUseClient
            rst_issued_all.Open "SELECT * FROM B_ISSUE ORDER BY DATE_ISSUED", dbcon, adOpenStatic, adLockReadOnly
        Set drAllissue.DataSource = rst_issued_all
            drAllissue.Show 1
    ElseIf dateselect = True Then
        If dtpBIFrom.Value > dtpBITo.Value Then
            MsgBox "Invalid Date Range.", vbExclamation
            dtpBIFrom.SetFocus
            Exit Sub
        End If
        Set rst_issue_Date = New ADODB.Recordset
            rst_issue_Date.CursorLocation = adUseClient
            rst_issue_Date.Open "SELECT* FROM B_ISSUE WHERE DATE_ISSUED >= " & "#" & dtpBIFrom & "#" & " AND DATE_ISSUED <= " & "#" & dtpBITo & "#" & "ORDER BY DATE_ISSUED", dbcon, adOpenStatic, adLockReadOnly
        Set drBookkIssueDate.DataSource = rst_issue_Date
            drBookkIssueDate.Sections("section4").Controls("Label5").Caption = Format(dtpBIFrom, "dd/mm/yyyy")
            drBookkIssueDate.Sections("section4").Controls("Label8").Caption = Format(dtpBITo, "dd/mm/yyyy")
            drBookkIssueDate.Show 1

    End If
    
ElseIf optBookIssue.Value = True And OptCoursewise.Value = True Then
    If dateselect = False Then
        If deGroupedreports.rscmdGroupedbyCourseBooksIssue_Grouping.State = adStateOpen Then
           deGroupedreports.rscmdGroupedbyCourseBooksIssue_Grouping.Close
        End If
        
        drGroupedbyModuleBooksIssue.Show 1
        
    ElseIf dateselect = True Then
            If dtpBIFrom.Value > dtpBITo.Value Then
            MsgBox "Invalid Date Range.", vbExclamation
            dtpBIFrom.SetFocus
            Exit Sub
            End If
    
        If deGroupedreports.rscmdGroupedbyModuleBook_Grouping.State = adStateOpen Then
            deGroupedreports.rscmdGroupedbyModuleBook_Grouping.Close
        End If
            deGroupedreports.cmdGroupedbyModuleBook_Grouping dtpBIFrom, dtpBITo
            Load drGroupedbyModuleBooksIssueDate
            drGroupedbyModuleBooksIssueDate.Sections("section4").Controls("Label14").Caption = Format(dtpBIFrom, "dd/mm/yyyy")
            drGroupedbyModuleBooksIssueDate.Sections("section4").Controls("Label13").Caption = Format(dtpBITo, "dd/mm/yyyy")
            drGroupedbyModuleBooksIssueDate.Show 1
   End If
  
   ' End of Books Issue Reports
   '-------------------------------------------------------------------------------

  
   '----------------------------------------------------------------------------------
   ' For Books Status Reports

  
    
ElseIf optBooksCheck.Value = True And optAll.Value = True Then
        If optLessthan.Value = True Then
            If deGroupedreports.rscmdBooksCheckLessthan.State = adStateOpen Then
                deGroupedreports.rscmdBooksCheckLessthan.Close
            End If
            
                deGroupedreports.cmdBooksCheckLessthan txtNumber
                Load drBookscheckLessthan
                drBookscheckLessthan.Sections("section4").Controls("Label7").Caption = txtNumber
                drBookscheckLessthan.Show 1
       ElseIf optGreaterthan.Value = True Then
            If deGroupedreports.rscmdBooksCheckGreaterthan.State = adStateOpen Then
                deGroupedreports.rscmdBooksCheckGreaterthan.Close
            End If
                deGroupedreports.cmdBooksCheckGreaterthan txtNumber
                Load drBookscheckGreaterthan
                drBookscheckGreaterthan.Sections("section4").Controls("Label7").Caption = txtNumber
                drBookscheckGreaterthan.Show 1
      End If
End If

   ' End of Books Status Reports
   '-------------------------------------------------------------------------------
End Sub

Private Sub Form_Load()
optCurrentStock.Value = True
OptCoursewise.Value = True
Me.Picture = frmStyle.Picture
'dtpBIFrom.Value = Format(Date, "dd/MM/yyyy")
'dtpBITo.Value = Format(Date, "dd/MM/yyyy")
'dtpBRFrom.Value = Format(Date, "dd/MM/yyyy")
'dtpBRTo.Value = Format(Date, "dd/MM/yyyy")
dtpBIFrom.Value = Date
dtpBITo.Value = Date
dtpBRFrom.Value = Date
dtpBRTo.Value = Date
Add_User_Names_to_Combo
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
deGroupedforcurstock.rscmdCurStock_Grouping.Close
dateselect = False
'deGroupedreports.cnGroupedreports.Close
End Sub

Private Sub optBookIssue_Click()

cmbCategorySelect.Enabled = False

chkBIDateRange.Value = 0
chkBIDateRange.Enabled = True

chkBRDateRange.Value = 0
chkBRDateRange.Enabled = False

optGreaterthan.Value = False
optLessthan.Value = False
optGreaterthan.Enabled = False
optLessthan.Enabled = False
txtNumber = ""
txtNumber.Enabled = False

OptCoursewise.Enabled = True
OptCoursewise.Caption = "Module wise Report"
OptCoursewise.Value = True
optCatetorywise.Enabled = False

dtpBRFrom.Value = Date
dtpBRTo.Value = Date

'Add_User_Names_to_Combo
End Sub

Private Sub optBookReceipt_Click()
cmbCategorySelect.Enabled = False

chkBRDateRange.Value = 0
chkBRDateRange.Enabled = True

chkBIDateRange.Value = 0
chkBIDateRange.Enabled = False

optGreaterthan.Value = False
optLessthan.Value = False
optGreaterthan.Enabled = False
optLessthan.Enabled = False
txtNumber = ""
txtNumber.Enabled = False

OptCoursewise.Enabled = True
OptCoursewise.Caption = "Module wise Report"
OptCoursewise.Value = True
optCatetorywise.Enabled = False

dtpBIFrom.Value = Date
dtpBITo.Value = Date

'Add_User_Names_to_Combo
End Sub

Private Sub optBooksCheck_Click()
cmbCategorySelect.Enabled = False

chkBIDateRange.Value = 0
chkBRDateRange.Value = 0

chkBIDateRange.Enabled = False
chkBRDateRange.Enabled = False

optLessthan.Enabled = True
optGreaterthan.Enabled = True
optLessthan.Value = True
txtNumber.Enabled = True
txtNumber = ""
txtNumber.SetFocus

OptCoursewise.Enabled = False
optCatetorywise.Enabled = False
optAll.Value = True

dtpBIFrom.Value = Date
dtpBITo.Value = Date
dtpBRFrom.Value = Date
dtpBRTo.Value = Date

'Add_User_Names_to_Combo
End Sub

Private Sub optCurrentStock_Click()

If optCurrentStock.Value = True Then
    cmbCategorySelect.Enabled = True
ElseIf optCurrentStock.Value = False Then
    cmbCategorySelect.Enabled = False
End If

cmbCategorySelect.Enabled = True
Add_User_Names_to_Combo

chkBIDateRange.Value = 0
chkBIDateRange.Enabled = False

chkBRDateRange.Value = 0
chkBRDateRange.Enabled = False

optLessthan.Value = False
optGreaterthan.Value = False
optGreaterthan.Enabled = False
optLessthan.Enabled = False
txtNumber = ""
txtNumber.Enabled = False

OptCoursewise.Enabled = True
OptCoursewise.Caption = "Course wise Report"
optCatetorywise.Enabled = True
optAll.Enabled = True
OptCoursewise.Value = True

dtpBIFrom.Value = Date
dtpBITo.Value = Date
dtpBRFrom.Value = Date
dtpBRTo.Value = Date

End Sub

Private Sub optGreaterthan_Click()
On Error Resume Next
txtNumber.SetFocus
End Sub

Private Sub optLessthan_Click()
On Error Resume Next
txtNumber.SetFocus
End Sub

Private Sub txtNumber_Change()
If Not IsNumeric(txtNumber) Then
    txtNumber = ""
End If
End Sub
Public Sub Add_User_Names_to_Combo()
On Error GoTo Err
Set rstgetcategory = New ADODB.Recordset
    rstgetcategory.CursorLocation = adUseClient
    rstgetcategory.Open "SELECT * FROM CATEGORY ORDER BY RECORD_ID", dbcon, adOpenStatic, adLockReadOnly
cmbCategorySelect.Clear
cmbCategorySelect.AddItem "----- All Stock -----"
If rstgetcategory.RecordCount > 0 Then
    Do While Not rstgetcategory.EOF
        cmbCategorySelect.AddItem rstgetcategory("CATEGORY")
    rstgetcategory.MoveNext
    Loop
End If

rstgetcategory.Close
Set rstgetcategory = Nothing

cmbCategorySelect = cmbCategorySelect.List(0)
Exit Sub
Err:
MsgBox Err.Description & "_" & Err.Number, vbCritical
End Sub
