VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commission setting"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3510
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
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
      Left            =   840
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3255
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
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Commission:"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "(ex : 1500)"
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Commission"
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
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Regular Slmn, Drvr, Hlpr, Sgegtr, and Rfnr only have commission."
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
      TabIndex        =   7
      Top             =   480
      Width           =   2775
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
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   0
      Picture         =   "Form14.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3900
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741



Option Explicit

Private Sub Command1_Click()
Dim cncom As New ADODB.connection
Dim rscom As New ADODB.recordset

If sempty(Text1) = True Then Exit Sub
If snumber(Text1) = True Then Call hlfocus(Text1): Exit Sub

Call connection(cncom, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rscom, cncom, "SELECT * FROM Commission")

With rscom
.Fields!comm = Text1.Text
.Update
End With

MsgBox "Successfully updated.", vbInformation, "Engcore Merchandising"

Unload Me

Set cncom = Nothing
Set rscom = Nothing

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
