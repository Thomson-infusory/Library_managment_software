VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeleteStudent 
   Caption         =   "Delete Student"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   Picture         =   "frmDeleteStudent.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmDeleteStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If rs.State = 1 Then rs.Close
    rs.Open " select * from student where sState ='" & "active" & "'", con, 3, 3
    fillheader
    fillgrid
End Sub

Private Sub fillheader()
        grid.Rows = 1
        grid.Cols = 5
        grid.ColAlignment(0) = 4
        grid.ColAlignment(3) = 1
        grid.ColAlignment(4) = 1
        grid.ColWidth(1) = 2200
        grid.ColWidth(2) = 1600
        grid.ColWidth(3) = 1600
        grid.ColWidth(4) = 1800
        grid.TextMatrix(0, 0) = "ID"
        grid.TextMatrix(0, 1) = "NAME"
        grid.TextMatrix(0, 2) = "COURSE"
        grid.TextMatrix(0, 3) = "SEMESTER"
        grid.TextMatrix(0, 4) = "PHONE"
End Sub

Private Sub fillgrid()
    i = 1
    While rs.EOF = False
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(i, 0) = rs.Fields!sId
        grid.TextMatrix(i, 1) = rs.Fields!sName
        grid.TextMatrix(i, 2) = rs.Fields!sCourse
        grid.TextMatrix(i, 3) = rs.Fields!sSemester
        grid.TextMatrix(i, 4) = rs.Fields!sPhone
        rs.MoveNext
        i = i + 1
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Picture1.Visible = True
End Sub

Private Sub grid_Click()
    If MsgBox("Are you sure to delete " & grid.TextMatrix(grid.RowSel, 1), vbYesNo) = vbYes Then
        If rs.State = 1 Then rs.Close
        rs.Open " select * from student where sState ='" & "active" & "'", con, 3, 3
        While Not (rs.Fields!sName = grid.TextMatrix(grid.RowSel, 1))
            rs.MoveNext
        Wend
        rs.Fields!sState = "inactive"
        rs.Update
        MsgBox ("Student Deleted")
        grid.RemoveItem (grid.Row)
    Else
        MsgBox ("Student Not Deleted")
    End If
End Sub

Private Sub txtSearch_Change()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from student where (sId like '" & txtSearch.Text & "%' Or sName like '" & txtSearch.Text & "%' Or sCourse like '" & txtSearch.Text & "%' Or sSemester like '" & txtSearch.Text & "%') and sState = 'active'"
    fillheader
    fillgrid
End Sub
