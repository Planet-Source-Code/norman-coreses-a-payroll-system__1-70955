VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Working days setting."
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Left            =   5640
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4440
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6855
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
         Height          =   285
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   3801089
         TitleBackColor  =   65280
         TrailingForeColor=   16777215
         CurrentDate     =   38937
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
         Height          =   285
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&No. of Working (15-30) Days:"
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
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "&Leaves Available:"
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
         Left            =   2880
         TabIndex        =   2
         Top             =   1200
         Width           =   1740
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Days"
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
      TabIndex        =   10
      Top             =   120
      Width           =   3135
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
      TabIndex        =   7
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Please filled the fields. Enter numeric only."
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
      TabIndex        =   6
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   -360
      Picture         =   "Form8.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7620
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Private Sub Command1_Click()
Dim cnleave As New ADODB.connection
Dim rsleave As New ADODB.recordset

If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If snumber(Text1) = True Then Call hlfocus(Text1): Exit Sub
If snumber(Text2) = True Then Call hlfocus(Text2): Exit Sub

Call connection(cnleave, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsleave, cnleave, "SELECT * FROM Workingdays")

With rsleave
.Fields!no_of_wkd = Text1.Text
.Fields!Avbl_leave = Text2.Text
.Update
End With

MsgBox "Leave setting successfully updated.", vbInformation, "Engcore Merchandising"

Set cnleave = Nothing
Set rsleave = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
MonthView1.Value = Date
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
