VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Addnew Employee."
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4080
      TabIndex        =   24
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
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
      Left            =   2880
      TabIndex        =   23
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5295
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
         Height          =   345
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1080
         Width           =   3375
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1560
         Width           =   3375
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
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   8
         Top             =   2040
         Width           =   3375
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   7
         Top             =   2520
         Width           =   3135
      End
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   6
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox Text7 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   3480
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   2
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
         Left            =   1560
         TabIndex        =   1
         Top             =   3960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   3000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   65280
         CalendarTrailingForeColor=   -2147483639
         Format          =   3801089
         CurrentDate     =   38935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   2520
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         CalendarTitleBackColor=   65280
         CalendarTrailingForeColor=   -2147483639
         Format          =   3801089
         CurrentDate     =   38935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Employee ID:"
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
         Top             =   360
         Width           =   1410
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
         TabIndex        =   19
         Top             =   1080
         Width           =   1125
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
         TabIndex        =   18
         Top             =   1560
         Width           =   945
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
         TabIndex        =   16
         Top             =   2520
         Width           =   1350
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
         TabIndex        =   15
         Top             =   3000
         Width           =   1200
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
         TabIndex        =   14
         Top             =   3480
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
         TabIndex        =   13
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "(&Last, First, Middle)"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   1320
      End
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Please filled all the fields."
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
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   480
      Width           =   2535
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
      TabIndex        =   21
      Top             =   480
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "Form1.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
Form2.Show 1
End Sub

Private Sub Command2_Click()
Dim cnnew As New ADODB.connection
Dim rsnew As New ADODB.recordset

If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text3) = True Then Exit Sub
If sempty(Text4) = True Then Exit Sub
If snumber(Text4) = True Then Call hlfocus(Text4): Exit Sub
If sempty(Text5) = True Then Exit Sub
If sempty(Text6) = True Then Exit Sub
If sempty(Text7) = True Then Exit Sub
If sempty(Combo1) = True Then Exit Sub

Call connection(cnnew, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsnew, cnnew, "SELECT * FROM EmployeeInfo")

If recexist("EmployeeInfo", "emp_id", Text1, cnnew) = True Then Call hlfocus(Text1): Exit Sub

With rsnew
.AddNew
.Fields!emp_id = Text1.Text
.Fields!f_name = Text2.Text
.Fields!address = Text3.Text
.Fields!contact_no = Text4.Text
.Fields!d_birth = Text5.Text
.Fields!d_hired = Text6.Text
.Fields!Designation = Text7.Text
.Fields!Status = Combo1.Text
.Update
End With

MsgBox "New Employee has been successfully saved.", vbInformation, "Engcore Merchandising"

Call clear
Text1.SetFocus

Set cnnew = Nothing
Set rsnew = Nothing

End Sub

Private Sub Command3_Click()
Unload Me
Form3.SetFocus
End Sub

Private Sub DTPicker1_Change()
Text5.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
Text5.Text = DTPicker1.Value
End Sub

Private Sub DTPicker2_Click()
Text6.Text = DTPicker2.Value
End Sub

Private Sub DTPicker2_Change()
Text6.Text = DTPicker2.Value
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Call combo
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Text = StrConv(Text1, vbProperCase)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text2.Text = StrConv(Text2, vbProperCase)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text3_LostFocus()
Text3.Text = StrConv(Text3, vbProperCase)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

Form2.Show 1

End Sub

Sub combo()
Combo1.AddItem "Contractual"
Combo1.AddItem "Regular"
End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = ""
End Sub
