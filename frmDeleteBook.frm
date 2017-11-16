VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeleteBook 
   Caption         =   "Delete Book"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   Picture         =   "frmDeleteBook.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   15480
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   12960
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
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
      TabIndex        =   2
      Top             =   360
      Width           =   9135
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Categories : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11160
      TabIndex        =   3
      Top             =   360
      Width           =   1455
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
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmDeleteBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboType_Click()
    If rs.State = 1 Then rs.Close
    If cboType.Text = "All" Then
        rs.Open "select * from stock", con, 3, 3
    Else
        rs.Open " select *from stock where bookType = '" & cboType.Text & "'", con, 3, 3
    End If
    grid.Clear
    fillheader
    fillgrid
End Sub

Private Sub Form_Load()
    If rs.State = 1 Then rs.Close
    rs.Open " select * from stock ", con, 3, 3
    fillheader
    fillgrid
    
    cboType.AddItem "All"
    If rs.State = 1 Then rs.Close
    rs.Open "select distinct bookType from stock", con, 3, 3
    While Not rs.EOF
        cboType.AddItem rs.Fields!bookType
        rs.MoveNext
    Wend
End Sub

Private Sub fillheader()
        grid.Rows = 1
        grid.Cols = 6
        grid.ColWidth(0) = 2000
        grid.ColWidth(1) = 3660
        grid.ColWidth(2) = 3660
        grid.ColWidth(3) = 3000
        grid.ColWidth(4) = 1500
        grid.ColAlignment(0) = 4
        grid.ColAlignment(4) = 1
        grid.ColAlignment(5) = 1
        grid.TextMatrix(0, 0) = "ID"
        grid.TextMatrix(0, 1) = "NAME"
        grid.TextMatrix(0, 2) = "AUTHOR"
        grid.TextMatrix(0, 3) = "TYPE"
        grid.TextMatrix(0, 4) = "COPIES"
        grid.TextMatrix(0, 5) = "PRICE"
End Sub

Private Sub fillgrid()
    i = 1
    While rs.EOF = False
        grid.Rows = grid.Rows + 1
        grid.TextMatrix(i, 0) = rs.Fields!bookId
        grid.TextMatrix(i, 1) = rs.Fields!bookName
        grid.TextMatrix(i, 2) = rs.Fields!bookAuthor
        grid.TextMatrix(i, 3) = rs.Fields!bookType
        grid.TextMatrix(i, 4) = rs.Fields!bookCopies
        grid.TextMatrix(i, 5) = rs.Fields!bookPrice
        rs.MoveNext
        i = i + 1
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmMain
    If frmadminLogin.lblAdmin.Caption = "true" Then frmadminLogin.adminMethod
    If frmstaffLogin.lblStaff.Caption = "true" Then frmstaffLogin.staffMethod
End Sub

Private Sub grid_Click()
    If MsgBox("Are you sure to delete " & grid.TextMatrix(grid.RowSel, 1), vbYesNo) = vbYes Then
        If rs.State = 1 Then rs.Close
        rs.Open "select * from stock", con, 3, 3
        While Not (rs.Fields!bookName = grid.TextMatrix(grid.RowSel, 1))
            rs.MoveNext
        Wend
        Dim temp As Integer
        temp = rs.Fields!bookCopies
        If temp = 1 Then
            MsgBox ("Last book deleted . ")
            rs.Delete
            rs.Update
            grid.RemoveItem (grid.RowSel)
        Else
            temp = temp - 1
            rs.Fields!bookCopies = temp
            MsgBox ("Book Copies reduced to " & temp)
            rs.Update
            rs.Close
            rs.Open "select * from stock", con, 3, 3
            grid.Clear
            fillheader
            fillgrid
        End If
'----------------------------------------------------------------
        cboType.Clear
        cboType.AddItem "All"
        If rs.State = 1 Then rs.Close
        rs.Open "select distinct bookType from stock", con, 3, 3
        While Not rs.EOF
            cboType.AddItem rs.Fields!bookType
            rs.MoveNext
        Wend
'----------------------------------------------------------------
        
    Else
        MsgBox ("Book Not Deleted")
    End If
End Sub

Private Sub txtSearch_Change()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from stock where bookName like '" & txtSearch.Text & "%' Or bookAuthor like '" & txtSearch.Text & "%' Or bookId like '" & txtSearch.Text & "%' Or bookType like '" & txtSearch.Text & "%'"
    fillheader
    fillgrid
End Sub
