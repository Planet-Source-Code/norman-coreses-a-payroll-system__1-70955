VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting."
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "&Designation Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Tax Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Cash Advanced Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Commission Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&User Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Working days Setting"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&View all Setting."
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
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "Form7.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3660
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741

Option Explicit

Private Sub Command1_Click()
Form8.Show 1
End Sub

Private Sub Command2_Click()
Form9.Show 1
End Sub

Private Sub Command3_Click()
Form12.Show 1
End Sub

Private Sub Command4_Click()
Form13.Show 1
End Sub

Private Sub Command5_Click()
Form14.Show 1
End Sub

Private Sub Command6_Click()
If MDIForm1.Label3.Caption <> "Admin" Then
MsgBox "This Function is for Administrator only. Log - In for Administrator", vbExclamation, "Engcore Merchandising": Exit Sub
End If

Form17.Show 1

End Sub

Private Sub Command7_Click()
Unload Me
End Sub
