VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payslip."
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   Icon            =   "Form15.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   53
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   48
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   47
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Regular Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   480
      TabIndex        =   25
      Top             =   3000
      Width           =   3855
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   46
         Text            =   "0"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   45
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   44
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   43
         Text            =   "0"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   42
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   41
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   40
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   38
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3000
         Picture         =   "Form15.frx":57E2
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "&Cash Advanced:"
         Height          =   195
         Left            =   480
         TabIndex        =   36
         Top             =   4080
         Width           =   1185
      End
      Begin VB.Label Label22 
         Caption         =   "&Net pay:"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "&Total Deduction:"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "&Other Deduction:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   33
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "&P - Health:"
         Height          =   195
         Left            =   480
         TabIndex        =   32
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "&P - Ibig:"
         Height          =   195
         Left            =   480
         TabIndex        =   31
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "&SSS:"
         Height          =   195
         Left            =   480
         TabIndex        =   30
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Deduction:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   1440
         Width           =   2580
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "&Grosspay:"
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "&Commission:"
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "&Basic pay:"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Contractual Setting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   480
      TabIndex        =   8
      Top             =   3000
      Width           =   3855
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Text            =   "0"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "0"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3000
         Picture         =   "Form15.frx":8524
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Cash Advanced:"
         Height          =   195
         Left            =   480
         TabIndex        =   23
         Top             =   3000
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Net pay:"
         Height          =   195
         Left            =   480
         TabIndex        =   21
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "&Other Deduction:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Grosspay:"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Rate:"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Days total:"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Leave:"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Days worked:"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
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
      Left            =   120
      TabIndex        =   56
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "&Select drop down button before Enter Employee ID. payslip (15 to 30) days period. Enter Leave, Otherdeducion optional."
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
      Left            =   720
      TabIndex        =   55
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label28 
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
      Left            =   120
      TabIndex        =   54
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "30/Aug/2006"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3360
      TabIndex        =   52
      Top             =   960
      Width           =   1290
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "&to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3120
      TabIndex        =   51
      Top             =   960
      Width           =   210
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2760
      TabIndex        =   50
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "&Payroll for the Month of:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   240
      TabIndex        =   49
      Top             =   960
      Width           =   2445
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Status:"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Designation:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Employee Name:"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Enter Employee ID:"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1365
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   0
      Picture         =   "Form15.frx":B266
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5100
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Private Sub Combo1_Change()
Call clear
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Combo1_Click
End If
End Sub

Private Sub Command1_Click()
Dim cnconaddnew As New ADODB.connection
Dim rsconaddnew As New ADODB.recordset

If sempty(Combo1) = True Then Exit Sub
If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text3) = True Then Exit Sub
If sempty(Text4) = True Then Exit Sub
If isempty(Text5) = True Then Exit Sub
If snumber(Text5) = True Then Call hlfocus(Text5): Exit Sub
If sempty(Text6) = True Then Exit Sub
If sempty(Text7) = True Then Exit Sub
If sempty(Text8) = True Then Exit Sub
If isempty(Text9) = True Then Exit Sub
If snumber(Text9) = True Then Call hlfocus(Text9): Exit Sub
If sempty(Text10) = True Then Exit Sub
If sempty(Text11) = True Then Exit Sub


Call connection(cnconaddnew, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsconaddnew, cnconaddnew, "SELECT * FROM Paygen")

With rsconaddnew
.AddNew
.Fields!emp_id = Combo1.Text
.Fields!f_name = Text1.Text
.Fields!Designation = Text2.Text
.Fields!Status = Text3.Text
.Fields!d_wrk = Text4.Text
.Fields!d_leave = Text5.Text
.Fields!d_total = Text6.Text
.Fields!d_rate = Text7.Text
.Fields!g_pay = Text8.Text
.Fields!tot_ded = Text9.Text
.Fields!ot_ded = Text9.Text
.Fields!n_pay = Text10.Text
.Fields!c_advanced = Text11.Text
.Fields!date_trans = Label25.Caption & " to " & Label27.Caption
.Fields!Date = Date
.Update
End With
                    
Call delcashcon
Call clear1
Combo1.SetFocus

MsgBox "Payslip successfully saved", vbInformation, "Engcore Merchandising"
                              
Set cnconaddnew = Nothing
Set rsconaddnew = Nothing

End Sub

Private Sub Command2_Click()
Dim cnregaddnew As New ADODB.connection
Dim rsregaddnew As New ADODB.recordset

If sempty(Combo1) = True Then Exit Sub
If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text3) = True Then Exit Sub
If sempty(Text12) = True Then Exit Sub
If sempty(Text13) = True Then Exit Sub
If sempty(Text14) = True Then Exit Sub
If sempty(Text15) = True Then Exit Sub
If sempty(Text16) = True Then Exit Sub
If sempty(Text17) = True Then Exit Sub
If isempty(Text18) = True Then Exit Sub
If snumber(Text18) = True Then Call hlfocus(Text18): Exit Sub
If sempty(Text19) = True Then Exit Sub
If sempty(Text20) = True Then Exit Sub
If sempty(Text21) = True Then Exit Sub

Call connection(cnregaddnew, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsregaddnew, cnregaddnew, "SELECT * FROM Paygen")

With rsregaddnew
.AddNew
.Fields!emp_id = Combo1.Text
.Fields!f_name = Text1.Text
.Fields!Designation = Text2.Text
.Fields!Status = Text3.Text
.Fields!b_pay = Text12.Text
.Fields!commission = Text13.Text
.Fields!g_pay = Text14.Text
.Fields!sss = Text15.Text
.Fields!p_ibig = Text16.Text
.Fields!p_health = Text17.Text
.Fields!ot_ded = Text18.Text
.Fields!tot_ded = Text19.Text
.Fields!n_pay = Text20.Text
.Fields!c_advanced = Text21.Text
.Fields!date_trans = Label25.Caption & " to " & Label27.Caption
.Fields!Date = Date
.Update
End With

Call delcashreg
Call clear2
Combo1.SetFocus

MsgBox "Payslip successfully saved.", vbInformation, "Engcore Merchandising"

Set cnregaddnew = Nothing
Set rsregaddnew = Nothing

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Visible = False
Command2.Visible = False

Frame1.Visible = False
Frame2.Visible = False
Call viewemployee


'Label25.Caption = Format(Now, Date - 14)
Label25.Caption = Format(Now, "dd") - 14
Label27.Caption = Format(Now, "dd/mmm/yyyy")
End Sub

Public Sub viewemployee()
Dim cnviewemp As New ADODB.connection
Dim rsviewemp As New ADODB.recordset
                                                                             
Call connection(cnviewemp, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsviewemp, cnviewemp, "SELECT emp_id FROM EmployeeInfo ORDER BY emp_id ASC")
                          
Combo1.clear
    With rsviewemp
        While Not .EOF
            Combo1.AddItem .Fields!emp_id
            .MoveNext
        Wend
    End With
    
Set cnviewemp = Nothing
Set rsviewemp = Nothing

End Sub

Private Sub Combo1_Click()
Dim cnview As New ADODB.connection
Dim rsview As New ADODB.recordset

Call connection(cnview, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsview, cnview, "SELECT * FROM EmployeeInfo WHERE emp_id ='" & Combo1.Text & "'")

If rsview.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

clear

With rsview
Text1.Text = .Fields!f_name
Text2.Text = .Fields!Designation
Text3.Text = .Fields!Status
End With

Call view

Set cnview = Nothing
Set rsview = Nothing

End Sub

Public Sub workingdays()
Dim cnwkd As New ADODB.connection
Dim rswkd As New ADODB.recordset

Call connection(cnwkd, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rswkd, cnwkd, "SELECT * FROM Workingdays")

With rswkd
Text4.Text = .Fields!no_of_wkd
End With

Set cnwkd = Nothing
Set rswkd = Nothing

End Sub

Public Sub cashcon()
Dim cnadv As New ADODB.connection
Dim rsadv As New ADODB.recordset

Call connection(cnadv, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsadv, cnadv, "SELECT c_adv FROM Cashadvanced WHERE emp_id='" & Combo1.Text & "'")

If rsadv.RecordCount = 0 Then
Exit Sub
End If

With rsadv
Text11.Text = .Fields!c_adv
End With

Set cnadv = Nothing
Set rsadv = Nothing
End Sub

Public Sub ratecon()
Dim cnrate As New ADODB.connection
Dim rsrate As New ADODB.recordset

Call connection(cnrate, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsrate, cnrate, "SELECT * FROM Designation WHERE d_designation='" & Text2.Text & "'")

With rsrate
Text7.Text = .Fields!r_con
End With

Set cnrate = Nothing
Set rsrate = Nothing

End Sub

Public Sub delcashcon()
Dim cndelcon As New ADODB.connection
Dim rsdelcon As New ADODB.recordset

Call connection(cndelcon, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdelcon, cndelcon, "SELECT * FROM Cashadvanced WHERE emp_id ='" & Combo1.Text & "'")

If rsdelcon.RecordCount = 0 Then
Exit Sub
End If

With rsdelcon
.Fields!c_adv = Val(.Fields!c_adv) - Val(Text9.Text)
.Update
End With

Call deletecon

Set cndelcon = Nothing
Set rsdelcon = Nothing
End Sub

Public Sub deletecon()
Dim cndelcon As New ADODB.connection
Dim rsdelcon As New ADODB.recordset

Call connection(cndelcon, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdelcon, cndelcon, "SELECT * FROM Cashadvanced WHERE emp_id='" & Combo1.Text & "'")

If rsdelcon.RecordCount = 0 Then
Exit Sub
End If

If Text9.Text = Val(Text11.Text) Then
With rsdelcon
.Delete
End With
End If

Set cndelcon = Nothing
Set rsdelcon = Nothing

End Sub

Public Sub computecon()
Text6.Text = Val(Text4.Text) - Val(Text5.Text)
Text8.Text = Val(Text6.Text) * Val(Text7.Text)
Text10.Text = Val(Text8.Text) - Val(Text9.Text)

End Sub

Sub clear1()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = "0"
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = "0"
Text10.Text = ""
Text11.Text = "0"
End Sub

Sub clear2()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text12.Text = ""
Text13.Text = "0"
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = "0"
Text19.Text = ""
Text20.Text = ""
Text21.Text = "0"
End Sub

Public Sub clear()
If Frame1.Visible = True Then
Call clear1

ElseIf Frame2.Visible = True Then
Call clear2
End If

End Sub
Public Sub view()
If Text3.Text = "Contractual" Then
Frame1.Visible = True
Frame2.Visible = False
Command1.Visible = True
Command2.Visible = False

Call workingdays
Call cashcon
Call ratecon


ElseIf Text3.Text = "Regular" Then
Frame2.Visible = True
Frame1.Visible = False
Command2.Visible = True
Command1.Visible = False

Call basepay
Call fillcom
Call ttax
Call cashreg

End If


End Sub
'********************************************************************
'********************************************************************

Public Sub basepay()
Dim cnbase As New ADODB.connection
Dim rsbase As New ADODB.recordset

Call connection(cnbase, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsbase, cnbase, "SELECT * FROM Designation WHERE d_designation='" & Text2.Text & "'")

With rsbase
Text12.Text = .Fields!r_reg
End With

Set cnbase = Nothing
Set rsbase = Nothing

End Sub

Public Sub commission()
Dim cncom As New ADODB.connection
Dim rscom As New ADODB.recordset

Call connection(cncom, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rscom, cncom, "SELECT * FROM Commission")

If rscom.RecordCount = 0 Then
Exit Sub
End If

With rscom
Text13.Text = .Fields!comm
End With

Set cncom = Nothing
Set rscom = Nothing

End Sub

Public Sub fillcom()
If Text2.Text = "Salesman" Or Text2.Text = "Driver" Or Text2.Text = "Helper" Or Text2.Text = "Segregator" Or Text2.Text = "Refiner" Then
Call commission
End If
End Sub

Public Sub ttax()
Dim cntax As New ADODB.connection
Dim rstax As New ADODB.recordset
Dim dss, dpi, dph As Double

Call connection(cntax, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rstax, cntax, "SELECT * FROM Tax")

With rstax
dss = .Fields!sss
dpi = .Fields!p_ibig
dph = .Fields!p_health
End With

Text15.Text = Val(Text12.Text) * dss
Text16.Text = Val(Text12.Text) * dpi
Text17.Text = Val(Text12.Text) * dph

Set cntax = Nothing
Set rstax = Nothing

End Sub

Public Sub computereg()
Text14.Text = Val(Text12.Text) + Val(Text13.Text)
Text19.Text = Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text)
Text20.Text = Val(Text14.Text) - Val(Text19.Text)

End Sub

Public Sub cashreg()
Dim cncadv As New ADODB.connection
Dim rscadv As New ADODB.recordset

Call connection(cncadv, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rscadv, cncadv, "SELECT * FROM Cashadvanced WHERE emp_id ='" & Combo1.Text & "'")

If rscadv.RecordCount = 0 Then
Exit Sub
End If

With rscadv
Text21.Text = .Fields!c_adv
End With

Set cncadv = Nothing
Set rscadv = Nothing
End Sub

Public Sub delcashreg()
Dim cndel As New ADODB.connection
Dim rsdel As New ADODB.recordset

Call connection(cndel, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdel, cndel, "SELECT * FROM Cashadvanced WHERE emp_id ='" & Combo1.Text & "'")

If rsdel.RecordCount = 0 Then
Exit Sub
End If

With rsdel
.Fields!c_adv = Val(.Fields!c_adv) - Val(Text18.Text)
.Update
End With

Call deletereg

Set cndel = Nothing
Set rsdel = Nothing

End Sub

Public Sub deletereg()
Dim cndelreg As New ADODB.connection
Dim rsdelreg As New ADODB.recordset

Call connection(cndelreg, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdelreg, cndelreg, "SELECT * FROM Cashadvanced WHERE emp_id='" & Combo1.Text & "'")

If rsdelreg.RecordCount = 0 Then
Exit Sub
End If

If Text18.Text = Val(Text21.Text) Then
With rsdelreg
.Delete
End With
End If

Set cndelreg = Nothing
Set rsdelreg = Nothing

End Sub


Private Sub Image1_Click()
On Error Resume Next
Shell "calc.exe", vbNormalFocus

End Sub

Private Sub Image2_Click()
On Error Resume Next
Shell "calc.exe", vbNormalFocus
End Sub


Private Sub Text10_Change()
Call computecon
End Sub

Private Sub Text12_Change()
Call computereg
End Sub

Private Sub Text13_Change()
Call computereg
End Sub

Private Sub Text15_Change()
Call computereg
End Sub

Private Sub Text16_Change()
Call computereg
End Sub

Private Sub Text17_Change()
Call computereg
End Sub

Private Sub Text18_Change()
Call computereg
End Sub

Private Sub Text18_LostFocus()
If Val(Text18.Text) > Val(Text21.Text) Then
MsgBox "Cannot accept > Cashadvanced.", vbExclamation, "Engcore Merchandising"
Text18.SetFocus
Call hlfocus(Text18)
Exit Sub
End If
End Sub

Private Sub Text19_Change()
Call computereg
End Sub

Private Sub Text4_Change()
Call computecon
End Sub

Private Sub Text5_Change()
Call computecon
End Sub

Private Sub Text6_Change()
Call computecon
End Sub

Private Sub Text7_Change()
Call computecon
End Sub

Private Sub Text8_Change()
Call computecon
End Sub


Private Sub Text9_Change()
Call computecon
End Sub

Private Sub Text9_LostFocus()
If Val(Text9.Text) > Val(Text11.Text) Then
MsgBox "Cannot accept > Cash Advanced.", vbExclamation, "Engcore Merchandising"
Text9.SetFocus
Call hlfocus(Text9)
Exit Sub
End If
End Sub
