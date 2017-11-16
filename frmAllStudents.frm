VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAllStudents 
   Caption         =   "Students"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   Picture         =   "frmAllStudents.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboState 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmAllStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboState_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "select *from student where sState = '" & cboState.Text & "'", con, 3, 3
    grid.Clear
    fillheader
    fillgrid
End Sub

Private Sub Form_Load()
    If rs.State = 1 Then rs.Close
    '----------------------------------
    cboState.AddItem "Active"
    cboState.AddItem "Inactive"
    '----------------------------------
    cboState.Text = "active"
    cboState_Click

End Sub

Private Sub fillheader()
        grid.Rows = 1
        grid.Cols = 6
        grid.ColAlignment(0) = 4
        grid.ColAlignment(4) = 1
        grid.ColAlignment(5) = 1
        grid.ColWidth(1) = 2200
        grid.ColWidth(3) = 1600
        grid.ColWidth(4) = 1600
        grid.ColWidth(5) = 1800
        grid.TextMatrix(0, 0) = "ID"
        grid.TextMatrix(0, 1) = "NAME"
        grid.TextMatrix(0, 2) = "SEX"
        grid.TextMatrix(0, 3) = "COURSE"
        grid.TextMatrix(0, 4) = "SEMESTER"
        grid.TextMatrix(0, 5) = "PHONE"
End Sub

Private Sub fillgrid()
    Dim i As Integer
    i = 1
    While rs.EOF = False
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(i, 0) = rs.Fields!sId
        grid.TextMatrix(i, 1) = rs.Fields!sName
        grid.TextMatrix(i, 2) = rs.Fields!sSex
        grid.TextMatrix(i, 3) = rs.Fields!sCourse
        grid.TextMatrix(i, 4) = rs.Fields!sSemester
        grid.TextMatrix(i, 5) = rs.Fields!sPhone
        rs.MoveNext
        i = i + 1
    Wend
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain.Picture1.Visible = True
End Sub
