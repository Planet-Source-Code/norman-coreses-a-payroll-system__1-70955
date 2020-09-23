VERSION 5.00
Begin VB.Form Form20 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log - In "
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   -5880
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&CLOSE"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1005
   End
   Begin VB.Image imgclose 
      Height          =   480
      Left            =   600
      Picture         =   "Form20.frx":57E2
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741

Option Explicit
Dim log As Boolean

Private Sub Command1_Click()
Dim cnlog As New ADODB.connection
Dim rslog As New ADODB.recordset

If Text1.Text = "" Then Text1.SetFocus: Exit Sub
If Text2.Text = "" Then Text2.SetFocus: Exit Sub

Call connection(cnlog, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rslog, cnlog, "SELECT * FROM UserAccount")

If recfound(rslog, "username", Text1.Text) = False Then

    MsgBox "Invalid Username.", vbExclamation, "Engcore Merchandising"
    Call terminate
    Text1.SetFocus
    Call hlfocus(Text1)
    
    Else
    
    If password = Text2.Text Then
    Me.Hide
    MDIForm1.Show
    
    Else
    
    MsgBox "Invalid Password.", vbExclamation, "Engcore Merchandising"
    Call terminate
    Text2.SetFocus
    Call hlfocus(Text2)
    
    End If
    End If

    log = True
    
    Set cnlog = Nothing
    Set rslog = Nothing
    
End Sub

Private Sub Command2_Click()
log = False
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then MsgBox "the system already Open.", vbInformation, "Engcore Merchandising: unload me :exit sub"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Private Sub Text1_LostFocus()
Text1.Text = StrConv(Text1, vbProperCase)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub

Public Sub terminate()
'Static attempt As Integer
'empt"
'                End
'            End If
End Sub

