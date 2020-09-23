VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash advanced setting."
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7560
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
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
      Left            =   6360
      TabIndex        =   12
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
      Left            =   5160
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8775
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
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   14
         Top             =   2280
         Width           =   2655
      End
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
         TabIndex        =   5
         Top             =   240
         Width           =   2655
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
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2655
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1800
         Width           =   2655
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   735
         Left            =   5040
         TabIndex        =   1
         Top             =   840
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1296
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Amount"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
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
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Amount:"
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
         TabIndex        =   6
         Top             =   2280
         Width           =   840
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Cash "
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
      TabIndex        =   17
      Top             =   120
      Width           =   3615
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
      Height          =   255
      Left            =   840
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   480
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   0
      Picture         =   "Form13.frx":57E2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9180
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Public Sub combo()
Dim cnview As New ADODB.connection
Dim rsview As New ADODB.recordset

Call connection(cnview, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsview, cnview, "SELECT * FROM EmployeeInfo ORDER BY emp_id ASC")

Combo1.clear
    With rsview
        While Not .EOF
            Combo1.AddItem .Fields!emp_id
            .MoveNext
        Wend
    End With
    
    
Set cnview = Nothing
Set rsview = Nothing

End Sub

Private Sub Combo1_Change()
Call clear
End Sub

Private Sub Combo1_Click()
Dim cnemployee As New ADODB.connection
Dim rsemployee As New ADODB.recordset

Call connection(cnemployee, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsemployee, cnemployee, "SELECT * FROM EmployeeInfo WHERE emp_id ='" & Combo1.Text & "'")

If rsemployee.RecordCount = 0 Then
MsgBox "The record you requested could not be found.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

Call viewlv

With rsemployee
Text1.Text = .Fields!f_name
Text2.Text = .Fields!Designation
Text3.Text = .Fields!Status
End With

Set cnemployee = Nothing
Set rsemployee = Nothing

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Combo1_Click
End If
End Sub

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
ListView1.ListItems.clear
End Sub

Public Sub viewlv()
Dim cnlv As New ADODB.connection
Dim rslv As New ADODB.recordset
Dim x

Call connection(cnlv, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rslv, cnlv, "SELECT * FROM Cashadvanced WHERE emp_id='" & Combo1.Text & "'")

ListView1.ListItems.clear
    With rslv
        While Not .EOF
            Set x = ListView1.ListItems.Add(, , .Fields!c_adv)
                x.SubItems(1) = .Fields!Date
                .MoveNext
        Wend
    End With
    
Set cnlv = Nothing
Set rslv = Nothing

End Sub

Private Sub Command1_Click()
Dim cnupdate As New ADODB.connection
Dim rsupdate As New ADODB.recordset
Dim reply As String

If sempty(Combo1) = True Then Exit Sub
If sempty(Text1) = True Then Exit Sub
If sempty(Text2) = True Then Exit Sub
If sempty(Text3) = True Then Exit Sub
If sempty(Text4) = True Then Exit Sub
If snumber(Text4) = True Then Call hlfocus(Text4): Exit Sub

Call connection(cnupdate, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rsupdate, cnupdate, "SELECT * FROM Cashadvanced")

If recexistdel("Cashadvanced", "emp_id", Combo1, cnupdate) = True Then Exit Sub

reply = MsgBox("Are you sure saved this entry.", vbExclamation + vbYesNo, "Engcore Merchandising")
If reply = vbNo Then
Exit Sub
End If

With rsupdate
.AddNew
.Fields!emp_id = Combo1.Text
.Fields!c_adv = Text4.Text
.Fields!Date = Date
.Update
End With

MsgBox "Successfully saved.", vbInformation, "Engcore Merchandising"

Combo1.Text = ""
Call clear
Call viewlv

Set cnupdate = Nothing
Set rsupdate = Nothing

End Sub

Private Sub Command2_Click()
Dim cncash As New ADODB.connection
Dim rscash As New ADODB.recordset

Call connection(cncash, App.Path & "\Engcorepay.mdb", "engman")
Call recordset(rscash, cncash, "SELECT * FROM Cashadvanced WHERE emp_id='" & Combo1.Text & "'")

If rscash.RecordCount = 0 Then
MsgBox "No record found. Enter amount first then click update.", vbExclamation, "Engcore Merchandising"
Exit Sub
End If

Set RPT1.DataSource = rscash

Unload Me

RPT1.Show 1

Set cncash = Nothing
Set rscash = Nothing


End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call combo
End Sub

