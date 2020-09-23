VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Designation setting."
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4920
      TabIndex        =   5
      Top             =   4680
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
      Left            =   3600
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
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
      Left            =   2280
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
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
      Left            =   960
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6135
      Begin MSComctlLib.ListView ListView1 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5953
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Designation"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Contractual Rate"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Regular Rate"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Designation"
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
      TabIndex        =   8
      Top             =   120
      Width           =   3735
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
      Caption         =   "&Addnew designation, edit and delete existing designation."
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
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "Form9.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6660
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741



Option Explicit

Public Sub viewposition()
Dim cnview As New ADODB.connection
Dim rsview As New ADODB.recordset
Dim x

Call connection(cnview, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsview, cnview, "SELECT * FROM Designation ORDER BY d_designation ASC")

ListView1.ListItems.clear

With rsview
    While Not .EOF
        Set x = ListView1.ListItems.Add(, , .Fields!d_designation)
            x.SubItems(1) = .Fields!r_con
            x.SubItems(2) = .Fields!r_reg
            .MoveNext
    Wend
End With

Set cnview = Nothing
Set rsview = Nothing

End Sub

Private Sub Command1_Click()
Form10.Show 1
End Sub

Private Sub Command2_Click()

If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

With Form11
.Text1.Text = ListView1.SelectedItem
.Text2.Text = ListView1.SelectedItem.SubItems(1)
.Text3.Text = ListView1.SelectedItem.SubItems(2)
.Show 1
End With
                                                        
End Sub
                                                                
Private Sub Command3_Click()
Dim cndel As New ADODB.connection
Dim rsdel As New ADODB.recordset
Dim reply As String

If ListView1.ListItems.Count < 1 Then MsgBox "No record found.", vbExclamation, "Engcore Merchandising": Exit Sub

Call connection(cndel, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsdel, cndel, "SELECT * FROM Designation WHERE d_designation ='" & ListView1.SelectedItem & "'")

If recexistlv("EmployeeInfo", "Designation", ListView1.SelectedItem, cndel) = True Then Exit Sub

reply = MsgBox("Are you sure you want to delete " & ListView1.SelectedItem & " ", vbExclamation + vbYesNo, "Engcore Merchandising")
If reply = vbNo Then
Exit Sub
End If

With rsdel
.Delete
End With

MsgBox "" & ListView1.SelectedItem & " Successfully deleted", vbInformation, "Engcore Merchandising"

Call viewposition

Set cndel = Nothing
Set rsdel = Nothing

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call viewposition
End Sub

Private Sub Form_Load()
Call viewposition
End Sub


