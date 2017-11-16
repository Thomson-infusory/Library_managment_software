VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAddBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Book"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAddBook.frx":0000
   ScaleHeight     =   6810
   ScaleWidth      =   15810
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   14280
      Picture         =   "frmAddBook.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   12240
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtCopies 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtAuthor 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtId 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3375
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   5953
      _Version        =   393216
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   12360
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copies"
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
      Left            =   10440
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   8280
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
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
      Left            =   5400
      TabIndex        =   10
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Byte
Dim temp As Integer

Private Sub cmdAdd_Click()
    If txtId.Text = "" Or txtName.Text = "" Or txtAuthor.Text = "" Or txtPrice.Text = "" Or cboType.Text = "" Or txtCopies.Text = "" Then
        i = 1
        GoTo p1
    End If
    rs.AddNew
        rs.Fields!bookId = UCase(txtId.Text)
        rs.Fields!bookName = UCase(txtName.Text)
        rs.Fields!bookAuthor = UCase(txtAuthor.Text)
        rs.Fields!bookType = cboType.Text
        rs.Fields!bookCopies = UCase(txtCopies.Text)
        rs.Fields!bookPrice = UCase(txtPrice.Text)
    rs.Update
    MsgBox ("Book Added")
    grid.Clear
    If rs.State = 1 Then rs.Close
    rs.Open "select * from stock", con, 3, 3
    fillheader
    fillgrid
    clearfields
    If i = 1 Then
p1:
        MsgBox ("Fill all fields")
    End If
End Sub



Private Sub clearfields()
    txtId.Text = ""
    txtName.Text = ""
    txtAuthor.Text = ""
    txtCopies.Text = ""
    txtPrice.Text = ""
'---------------------------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select bookId from stock", con, 3, 3
    rs.MoveLast
    temp = rs.Fields(0)
    txtId.Text = temp + 1
'---------------------------------------------------
    txtName.SetFocus
End Sub

Private Sub Form_Load()
    i = 0
'-----------------------------------------------------
    txtId.Enabled = False
    If rs.State = 1 Then rs.Close
    rs.Open "select bookId from stock", con, 3, 3
    If rs.BOF = True And rs.EOF = True Then
        txtId.Text = "1300"
    Else
    rs.MoveLast
    temp = rs.Fields(0)
    txtId.Text = temp + 1
    End If
'-----------------------------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open " select * from stock ", con, 3, 3
    cboType.AddItem "AUTOBIOGRAPHY"
    cboType.AddItem "STUDY"
    cboType.AddItem "NOVEL"
    cboType.AddItem "ELECTRONICS"
    cboType.AddItem "COMPUTER"
    fillheader
    fillgrid
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

Private Sub txtAuthor_Change()
    txtAuthor.MaxLength = 25
End Sub

Private Sub txtCopies_Change()
    txtCopies.MaxLength = 3
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtId_Change()
    txtId.MaxLength = 5
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_Change()
    txtName.MaxLength = 35
End Sub

Private Sub txtPrice_Change()
    txtPrice.MaxLength = 4
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
