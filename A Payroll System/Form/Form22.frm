VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Form22.frx":0000
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3240
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   3975
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Payroll System: "
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1680
            TabIndex        =   6
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "vr 2.0.0 "
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1680
            TabIndex        =   5
            Top             =   480
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   900
            Left            =   240
            Picture         =   "Form22.frx":57E2
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "&E-mail : cell_nor@yahoo.com"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1680
            TabIndex        =   3
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Contact#: +639212733741"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1680
            TabIndex        =   2
            Top             =   720
            Width           =   1920
         End
      End
      Begin VB.Image Image2 
         Height          =   900
         Left            =   240
         Picture         =   "Form22.frx":84F3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3945
      End
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741

Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

