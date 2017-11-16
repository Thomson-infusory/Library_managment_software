VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeleteStaff 
   Caption         =   "Delete Staff"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   Picture         =   "frmDeleteStaff.frx":0000
   ScaleHeight     =   6105
   ScaleWidth      =   14010
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
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   11600
      _ExtentX        =   20452
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmDeleteStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If rs.State = 1 Then rs.Close
    rs.Open " select * from staff where stState ='" & "active" & "'", con, 3, 3
    fillheader
    fillgrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Picture1.Visible = True
End Sub

Private Sub fillheader()
        grid.Rows = 1
        grid.Cols = 7
        grid.ColAlignment(0) = 4
        grid.ColAlignment(3) = 1
        grid.ColAlignment(5) = 1
        grid.ColAlignment(6) = 1
        grid.ColWidth(1) = 2300
        grid.ColWidth(4) = 2500
        grid.ColWidth(5) = 1800
        grid.ColWidth(6) = 2000
        grid.TextMatrix(0, 0) = "ID"
        grid.TextMatrix(0, 1) = "NAME"
        grid.TextMatrix(0, 2) = "SEX"
        grid.TextMatrix(0, 3) = "AGE"
        grid.TextMatrix(0, 4) = "PLACE"
        grid.TextMatrix(0, 5) = "PHONE"
        grid.TextMatrix(0, 6) = "DOJ"
End Sub

Private Sub fillgrid()
    i = 1
    While rs.EOF = False
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(i, 0) = rs.Fields!stId
        grid.TextMatrix(i, 1) = rs.Fields!stName
        grid.TextMatrix(i, 2) = rs.Fields!stSex
        grid.TextMatrix(i, 3) = rs.Fields!stAge
        grid.TextMatrix(i, 4) = rs.Fields!stPlace
        grid.TextMatrix(i, 5) = rs.Fields!stPhone
        grid.TextMatrix(i, 6) = rs.Fields!stDoj
        rs.MoveNext
        i = i + 1
    Wend
End Sub

Private Sub grid_Click()
    If MsgBox("Are you sure to delete " & grid.TextMatrix(grid.RowSel, 1), vbYesNo) = vbYes Then
        If rs.State = 1 Then rs.Close
        rs.Open " select * from staff where stState ='" & "active" & "'", con, 3, 3
        While Not (rs.Fields!stName = grid.TextMatrix(grid.RowSel, 1))
            rs.MoveNext
        Wend
        rs.Fields!stState = "inactive"
        rs.Update
        MsgBox ("Staff Deleted")
        grid.RemoveItem (grid.Row)
    Else
        MsgBox ("Staff Not Deleted")
    End If
End Sub

Private Sub txtSearch_Change()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from staff where (stId like '" & txtSearch.Text & "%' Or stName like '" & txtSearch.Text & "%' Or stDoj like '" & txtSearch.Text & "%') and stState ='active'"
    fillheader
    fillgrid
End Sub
