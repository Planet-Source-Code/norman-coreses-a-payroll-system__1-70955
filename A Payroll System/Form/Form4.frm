VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit existing employee."
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5655
      Begin VB.TextBox Text6 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   3480
         Width           =   3135
      End
      Begin VB.TextBox Text5 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   10
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox Text4 
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   9
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox Text3 
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   8
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox Text2 
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
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1560
         Width           =   3375
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
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2160
         TabIndex        =   5
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   3480
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   3000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   65280
         CalendarTrailingForeColor=   16777215
         Format          =   3735553
         CurrentDate     =   38936
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   2520
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarTitleBackColor=   65280
         CalendarTrailingForeColor=   16777215
         Format          =   3735553
         CurrentDate     =   38936
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "(&Last, First, Middle)"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Status:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Designation:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Date Hired:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Date of Birth:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Contact No:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Address:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Full Name:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter Employee ID:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2010
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Employee"
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
      TabIndex        =   25
      Top             =   120
      Width           =   2895
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
      TabIndex        =   1
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Select drop down button before Enter Employee ID."
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
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "Form4.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6060
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Public Sub combo()
Dim cncombo As New ADODB.connection
Dim rscombo As New ADODB.recordset

Call connection(cncombo, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rscombo, cncombo, "SELECT * FROM EmployeeInfo ORDER BY emp_id ASC")

Combo1.clear

    With rscombo
        While Not .EOF
            Combo1.AddItem .Fields!emp_id
            .MoveNext
        Wend
    End With
    
Set cncombo = Nothing
Set rscombo = Nothing

End Sub

Private Sub Combo1_Change()
Call clear
End Sub

Private Sub Combo1_Click()
Dim cnview As New ADODB.connection
Dim rsview As New ADODB.recordset

Call connection(cnview, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsview, cnview, "SELECT * FROM EmployeeInfo WHERE emp_id='" & Combo1.Text & "'")

If rsview.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

With rsview
Text1.Text = .Fields!f_name
Text2.Text = .Fields!address
Text3.Text = .Fields!contact_no
Text4.Text = .Fields!d_birth
Text5.Text = .Fields!d_hired
Text6.Text = .Fields!Designation
Combo2.Text = .Fields!Status
End With

Set cnview = Nothing
Set rsview = Nothing

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Combo1_Click
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Form5.Show 1
End Sub

Private Sub Command2_Click()
Dim cnupdate As New ADODB.connection
Dim rsupdate As New ADODB.recordset

If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text3) = True Then Exit Sub
If snumber(Text3) = True Then Call hlfocus(Text3): Exit Sub
If sempty(Text4) = True Then Exit Sub
If sempty(Text5) = True Then Exit Sub
If sempty(Text6) = True Then Exit Sub
If sempty(Combo2) = True Then Exit Sub

Call connection(cnupdate, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsupdate, cnupdate, "SELECT * FROM EmployeeInfo WHERE emp_id ='" & Combo1.Text & "'")

If rsupdate.RecordCount = 0 Then
MsgBox "No record found.", vbExclamation, "Engcore Merchandisng"
Combo1.SetFocus
Exit Sub
End If

With rsupdate
.Fields!f_name = Text1.Text
.Fields!address = Text2.Text
.Fields!contact_no = Text3.Text
.Fields!d_birth = Text4.Text
.Fields!d_hired = Text5.Text
.Fields!Designation = Text6.Text
.Fields!Status = Combo2.Text
.Update
End With

MsgBox "Employee ID " & Combo1 & " successfully updated.", vbExclamation, "Engcore Merchandising"

Call clear
Combo1.SetFocus

Set cnupdate = Nothing
Set rsupdate = Nothing
End Sub

Private Sub Command3_Click()
Unload Me
Form3.SetFocus
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
Text4.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Change()
Text5.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Click()
Text5.Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
Call combo
Call viewcombo
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Sub viewcombo()
Combo2.AddItem "Contractual"
Combo2.AddItem "Regular"
End Sub


Private Sub Text1_LostFocus()
Text1.Text = StrConv(Text1, vbProperCase)
End Sub

Private Sub Text2_LostFocus()
Text2.Text = StrConv(Text2, vbProperCase)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Form5.Show 1
End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo2.Text = ""
End Sub
