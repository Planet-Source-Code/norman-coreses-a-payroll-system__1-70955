VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View payslip."
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   10455
      Begin MSComctlLib.ListView ListView1 
         Height          =   3615
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6376
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   19
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EMPLOYEE ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "DESIGNATION"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "STATUS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "DAYS WORK"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "DAYS LEAVE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "DAYS TOTAL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "RATE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "OTHER DEDUCTION"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "SSS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "PAG - IBIG"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "PHIL - HEALTH"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "TOTAL DEDUCTION"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "COMMISSION"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "BASIC PAY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   15
            Text            =   "CASH ADVANCED"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "GROSS PAY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   17
            Text            =   "NET PAY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   18
            Text            =   "DATE TRANSACTION"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "&Total Employee:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   3960
         Width           =   75
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Total Net pay:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   7200
         TabIndex        =   11
         Top             =   3960
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8520
         TabIndex        =   10
         Top             =   3960
         Width           =   75
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Print "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Indv Print"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   65280
      CalendarTrailingForeColor=   16777215
      Format          =   3801089
      CurrentDate     =   38939
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "&Payslip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Enter Date and print, Delete existing payslip"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   480
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   -240
      Picture         =   "Form16.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Enter Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1170
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Private Sub Command1_Click()
On Error Resume Next

Dim cnsearch As New ADODB.connection
Dim rssearch As New ADODB.recordset
Dim x
Dim tot As Double

Call connection(cnsearch, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rssearch, cnsearch, "SELECT * FROM Paygen WHERE date='" & Text1.Text & "'")

If rssearch.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Engcore Merchandising"
Label3.Caption = ""
Label5.Caption = ""
ListView1.ListItems.clear
Text1.SetFocus
Call hlfocus(Text1)
Exit Sub
End If

ListView1.ListItems.clear

With rssearch

    While Not .EOF
        Set x = ListView1.ListItems.Add(, , .Fields!emp_id)
            x.SubItems(1) = .Fields!f_name
            x.SubItems(2) = .Fields!Designation
            x.SubItems(3) = .Fields!Status
            x.SubItems(4) = .Fields!d_wrk
            x.SubItems(5) = .Fields!d_leave
            x.SubItems(6) = .Fields!d_total
            x.SubItems(7) = .Fields!d_rate
            x.SubItems(8) = .Fields!ot_ded
            x.SubItems(9) = .Fields!sss
            x.SubItems(10) = .Fields!p_ibig
            x.SubItems(11) = .Fields!p_health
            x.SubItems(12) = .Fields!tot_ded
            x.SubItems(13) = .Fields!commission
            x.SubItems(14) = .Fields!b_pay
            x.SubItems(15) = .Fields!c_advanced
            x.SubItems(16) = .Fields!g_pay
            x.SubItems(17) = .Fields!n_pay
            x.SubItems(18) = .Fields!date_trans
          
            .MoveNext
            tot = tot + Val(x.SubItems(17))
            
    Wend
End With

Label3.Caption = ListView1.ListItems.Count
Label5.Caption = tot

Set cnsearch = Nothing
Set rssearch = Nothing


End Sub

Private Sub Command2_Click()
Dim cnindvprnt As New ADODB.connection
Dim rsindvprnt As New ADODB.recordset

Call connection(cnindvprnt, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsindvprnt, cnindvprnt, "SELECT * FROM Paygen WHERE date='" & Text1.Text & "'")

If rsindvprnt.RecordCount = 0 Then
MsgBox "Select Search button first.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

Set RPT2.DataSource = rsindvprnt

RPT2.Show 1

Set cnindvprnt = Nothing
Set rsindvprnt = Nothing

End Sub

Private Sub Command3_Click()
Dim cnprint As New ADODB.connection
Dim rsprint As New ADODB.recordset

Call connection(cnprint, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsprint, cnprint, "SELECT * FROM Paygen WHERE date ='" & Text1.Text & "'")

If rsprint.RecordCount = 0 Then
MsgBox "Select search button first.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

Set RPT3.DataSource = rsprint

RPT3.Show 1

Set cnprint = Nothing
Set rsprint = Nothing

End Sub

Private Sub Command4_Click()
On Error Resume Next

Dim cndelete As New ADODB.connection
Dim rsdelete As New ADODB.recordset
Dim reply As String

If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

Call connection(cndelete, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdelete, cndelete, "SELECT date FROM Paygen WHERE date='" & Text1.Text & "'")

reply = MsgBox("Are you sure you want to delete this all record.", vbInformation + vbYesNo, "Engcore Merchandising")
If reply = vbNo Then
Exit Sub
End If

With rsdelete
.Delete
End With

MsgBox "Successfully deleted.", vbInformation, "Engcore Merchandising"

Call Command1_Click

Set cndelete = Nothing
Set rsdelete = Nothing

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub DTPicker1_Change()
Text1.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
Text1.Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
Text1.Text = Date
End Sub


Private Sub Label5_Change()
Label5.Caption = Format(Label5, "#,##0.00")
End Sub

Private Sub Text1_Change()
Call Command1_Click
End Sub
