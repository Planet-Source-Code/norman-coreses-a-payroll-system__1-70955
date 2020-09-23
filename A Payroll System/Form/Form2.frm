VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Designation."
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3600
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
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
      Left            =   960
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
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
      Left            =   2160
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Designation"
            Object.Width           =   5292
         EndProperty
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741

Option Explicit

Private Sub Command1_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

With Form1
.Text7.Text = ListView1.SelectedItem
End With

Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Public Sub fillLv()
Dim cnlv As New ADODB.connection
Dim rslv As New ADODB.recordset
Dim x

Call connection(cnlv, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rslv, cnlv, "SELECT * FROM Designation ORDER BY d_designation ASC")

ListView1.ListItems.clear

With rslv
    While Not .EOF
        Set x = ListView1.ListItems.Add(, , .Fields!d_designation)
                .MoveNext
    Wend
End With

Set cnlv = Nothing
Set rslv = Nothing

End Sub

Private Sub Form_Load()
Call fillLv
End Sub

Private Sub ListView1_DblClick()
Call Command1_Click
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
Call Command1_Click
End Sub
