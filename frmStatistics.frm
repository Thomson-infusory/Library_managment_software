VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStatistics 
   Caption         =   "Statistics"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   Picture         =   "frmStatistics.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Returned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   2775
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4895
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Issued"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   9855
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2895
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5106
         _Version        =   393216
      End
   End
   Begin VB.ComboBox cboStudentId 
      BackColor       =   &H00C0FFFF&
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
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
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
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub cboStudentId_Click()
    rs.Close
    rs.Open "select * from stati where studentId = '" & cboStudentId.Text & "' and returned = 'no' ", con, 3, 3
    grid1.Clear
    grid1.Rows = 1
    fillIssuedHeader
    fillIssue
    rs.Close
    rs.Open "select * from stati where studentId = '" & cboStudentId.Text & "' and returned = 'yes' ", con, 3, 3
    grid2.Clear
    grid2.Rows = 1
    fillReturnHeader
    fillReturn
End Sub

Private Sub Command1_Click()
    DataReport1.Show
End Sub



Private Sub Form_Load()
    i = 1
    If rs.State = 1 Then rs.Close
    rs.Open "select distinct studentId from stati ", con, 3, 3
    While Not rs.EOF
        cboStudentId.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
    fillIssuedHeader
    fillReturnHeader
End Sub

Private Sub fillIssuedHeader()
    grid1.Rows = 1
    grid1.Cols = 3
    grid1.ColWidth(0) = 3000
    grid1.ColWidth(1) = 3000
    grid1.ColWidth(2) = 3000
    grid1.TextMatrix(0, 0) = "ID"
    grid1.TextMatrix(0, 1) = "NAME"
    grid1.TextMatrix(0, 2) = "ISSUE DATE"
End Sub

Private Sub fillIssue()
    i = 1
    While rs.EOF = False
        grid1.Rows = grid1.Rows + 1
        grid1.TextMatrix(i, 0) = rs.Fields!bookId
        grid1.TextMatrix(i, 1) = rs.Fields!bookName
        grid1.TextMatrix(i, 2) = rs.Fields!issueDate
        rs.MoveNext
        i = i + 1
    Wend
End Sub

Private Sub fillReturnHeader()
    grid2.Rows = 1
    grid2.Cols = 3
    grid2.ColWidth(0) = 3000
    grid2.ColWidth(1) = 3000
    grid2.ColWidth(2) = 3000
    grid2.TextMatrix(0, 0) = "ID"
    grid2.TextMatrix(0, 1) = "NAME"
    grid2.TextMatrix(0, 2) = "RETURN DATE"
End Sub

Private Sub fillReturn()
    Dim j As Integer
    j = 1
    While rs.EOF = False
        grid2.Rows = grid2.Rows + 1
        grid2.TextMatrix(j, 0) = rs.Fields!bookId
        grid2.TextMatrix(j, 1) = rs.Fields!bookName
        grid2.TextMatrix(j, 2) = rs.Fields!issueDate
        rs.MoveNext
        j = j + 1
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmMain.Picture1.Visible = True
End Sub
