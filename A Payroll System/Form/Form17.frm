VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form17 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System user setting."
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   Icon            =   "Form17.frx":0000
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
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
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit"
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
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
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
      Left            =   5280
      TabIndex        =   9
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
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
      Left            =   6480
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7695
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   1680
         Width           =   2175
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
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
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
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   2175
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
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   2175
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000018&
         Caption         =   "&View password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "USERNAME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Firstname:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Mi"
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Lastname:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Username:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Password:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Confirm password:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   1560
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&User"
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
      TabIndex        =   21
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Addnew Username and delete existing username. select listview to edit, delete the existing username."
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
      Height          =   135
      Left            =   840
      TabIndex        =   20
      Top             =   480
      Width           =   6975
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
      TabIndex        =   19
      Top             =   480
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   -240
      Picture         =   "Form17.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9060
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Private Sub Check1_Click()

If Check1.Value = 1 Then
    Text5.PasswordChar = "*"
    Text6.PasswordChar = "*"
Else
    Text5.PasswordChar = ""
    Text6.PasswordChar = ""
End If

End Sub

Private Sub Command1_Click()
Form18.Show 1
End Sub

Private Sub Command2_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

If Text4.Text = "" Then
MsgBox "Select Username.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

With Form19
.Text1.Text = Text4.Text
.Text2.Text = Text1.Text
.Text3.Text = Text2.Text
.Text4.Text = Text3.Text
.Text5.Text = Text4.Text
.Text6.Text = Text5.Text
.Text7.Text = Text6.Text
.Show 1
End With


End Sub

Private Sub Command3_Click()
Dim cndeluser As New ADODB.connection
Dim rsdeluser As New ADODB.recordset
Dim reply As String

If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

Call connection(cndeluser, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdeluser, cndeluser, "SELECT * FROM UserAccount WHERE username ='" & ListView1.SelectedItem & "'")

reply = MsgBox("Are you sure you want to delete this username,", vbInformation + vbYesNo, "Engcore Merchandising")
If reply = vbNo Then
Exit Sub
End If

With rsdeluser
.Delete
End With

MsgBox "Username " & ListView1.SelectedItem & " Successfully deleted.", vbInformation, "Engcore Merchandising"

Call viewuser
Call clear

Set cndeluser = Nothing
Set rsdeluser = Nothing

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call viewuser
End Sub

Private Sub Form_Load()
Call viewuser
End Sub

Public Sub viewuser()
Dim cnviewlv As New ADODB.connection
Dim rsviewlv As New ADODB.recordset
Dim x

Call connection(cnviewlv, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsviewlv, cnviewlv, "SELECT * FROM UserAccount ORDER BY username ASC")

If rsviewlv.RecordCount = 0 Then
Exit Sub
End If

ListView1.ListItems.clear

    With rsviewlv
        While Not .EOF
            Set x = ListView1.ListItems.Add(, , .Fields!username)
                x.SubItems(1) = Trim(.Fields!l_name) & ", " & Trim(.Fields!f_name) & " " & Trim(.Fields!mi)
                .MoveNext
        Wend
    End With
    
Set cnviewlv = Nothing
Set rsviewlv = Nothing

End Sub
Private Sub ListView1_Click()
If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

Text4.Text = ListView1.SelectedItem

Call fill

End Sub

Public Sub fill()
Dim cnview As New ADODB.connection
Dim rsview As New ADODB.recordset

Call connection(cnview, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsview, cnview, "SELECT * FROM UserAccount WHERE username='" & Text4.Text & "'")

If rsview.RecordCount = 0 Then
Exit Sub
End If

With rsview
Text1.Text = .Fields!f_name
Text2.Text = .Fields!mi
Text3.Text = .Fields!l_name
Text5.Text = .Fields!password
Text6.Text = .Fields!password
End With

Set cnview = Nothing
Set rsview = Nothing

End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
