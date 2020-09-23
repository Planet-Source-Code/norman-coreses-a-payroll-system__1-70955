VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete existing employee."
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
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
      TabIndex        =   14
      Top             =   4680
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
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   5655
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
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Enter Employee ID:"
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
         TabIndex        =   10
         Top             =   240
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   5655
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   600
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   960
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   3375
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2040
         Width           =   3375
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
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
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   945
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Date Hired"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label6 
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
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label7 
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
         Left            =   360
         TabIndex        =   1
         Top             =   2400
         Width           =   720
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
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   120
      Width           =   4575
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
      TabIndex        =   12
      Top             =   480
      Width           =   3975
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
      TabIndex        =   11
      Top             =   480
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   -480
      Picture         =   "Form6.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6780
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Public Sub viewemployee()
Dim cnviewemp As New ADODB.connection
Dim rsviewemp As New ADODB.recordset

Call connection(cnviewemp, App.Path & "\EngcorePay.mdb", "engman")
Call recordset(rsviewemp, cnviewemp, "SELECT * FROM EmployeeInfo ORDER BY emp_id ASC")

If rsviewemp.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Engcore Merchandising"
Combo1.SetFocus
Exit Sub
End If

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


Private Sub Combo1_Change()
Call clear
End Sub

Private Sub Combo1_Click()
Dim cndelemployee As New ADODB.connection
Dim rsdelemployee As New ADODB.recordset

Call connection(cndelemployee, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdelemployee, cndelemployee, "SELECT * FROM EmployeeInfo WHERE emp_id ='" & Combo1.Text & "'")

If rsdelemployee.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Engcore Merchandising"
Combo1.SetFocus
Exit Sub
End If

With rsdelemployee
Text1.Text = .Fields!f_name
Text2.Text = .Fields!address
Text3.Text = .Fields!contact_no
Text4.Text = .Fields!d_birth
Text5.Text = .Fields!d_hired
Text6.Text = .Fields!Designation
Text7.Text = .Fields!Status
End With

Set cndelemployee = Nothing
Set rsdelemployee = Nothing

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Combo1_Click
End If
End Sub

Private Sub Command2_Click()
Dim cndel As New ADODB.connection
Dim rsdel As New ADODB.recordset
Dim reply As String

Call connection(cndel, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdel, cndel, "SELECT * FROM EmployeeInfo WHERE emp_id='" & Combo1.Text & "'")

If rsdel.RecordCount = 0 Then
MsgBox "No record found.", vbExclamation, "Engcore Merchandising"
Combo1.SetFocus
Exit Sub
End If

If existdel("Cashadvanced", "emp_id", Combo1, cndel) = True Then Exit Sub

reply = MsgBox("Are you sure you want to delete this record.", vbExclamation + vbYesNo, "Engcore Merchandising")
If reply = vbNo Then
Exit Sub
End If

With rsdel
.Delete
End With

MsgBox "Employee Id " & Combo1 & " Successfully deleted,", vbInformation, "Engcore Merchandising"

Call viewemployee
Call clear
Combo1.SetFocus

Set cndel = Nothing
Set rsdel = Nothing

End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
Form3.SetFocus

End Sub

Private Sub Form_Load()
Call viewemployee
End Sub


