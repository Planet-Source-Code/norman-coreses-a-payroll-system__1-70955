VERSION 5.00
Begin VB.Form Form21 
   BorderStyle     =   0  'None
   Caption         =   "Form21"
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form21"
   ScaleHeight     =   1560
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   1560
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Enter Password:"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Image imgclose 
      Height          =   480
      Left            =   3000
      Picture         =   "Form21.frx":0000
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741

Option Explicit

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1.Text <> password Then
        MsgBox "INVALID PASSWORD. Please type the correct password.", vbCritical, "Engcore Merchandising"
        Call hlfocus(Text1)
    Else
        Unload Me
    End If
End If
    
End Sub

