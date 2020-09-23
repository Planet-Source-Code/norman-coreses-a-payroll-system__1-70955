VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Engcore Merchandising Payroll System ver 2.0.0"
   ClientHeight    =   3090
   ClientLeft      =   -240
   ClientTop       =   555
   ClientWidth     =   5355
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   2835
      Width           =   5355
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
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
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   5295
      TabIndex        =   1
      Top             =   975
      Width           =   5355
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   12600
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":57E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":84EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":BFF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":11878
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":17BEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":1A8F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":205D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":232E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":25032
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   570
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   1005
         ButtonWidth     =   2699
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Edit"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "View Record"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Setting"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Payslip"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "View payslip"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Lock"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "About"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   5325
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      Begin VB.Image Image2 
         Height          =   900
         Left            =   600
         Picture         =   "MDIForm1.frx":2A8B4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   2040
         Picture         =   "MDIForm1.frx":2D5C5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7305
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' comments or suggestion please email @: cell_nor@yahoo.com
' if you want full code of this system just contact @: 639212733741


Option Explicit

Private Sub Command11_Click()
End
End Sub

Private Sub MDIForm_Load()
Form3.Show

MDIForm1.Label3.Caption = "" & username

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reply As String

reply = MsgBox("Close Application.", vbExclamation + vbYesNo, "Engcore Merchandising")
If reply = vbNo Then
Cancel = 1
Exit Sub
Else
End
End If

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
        Picture2.Width = ScaleWidth
        
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2: Form1.Show 1
        Case 4: Form4.Show 1
        Case 6: Form6.Show 1
        Case 8: Form3.Show
        Case 10: Form7.Show 1
        Case 12: Form15.Show 1
        Case 14: Form16.Show 1
        Case 16: Form21.Show 1
        Case 18: Form22.Show 1
    End Select
End Sub
