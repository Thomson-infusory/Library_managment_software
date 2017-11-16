VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReturnBook 
   Caption         =   "Return Book"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   Picture         =   "frmReturnBook.frx":0000
   ScaleHeight     =   9030
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   3480
      TabIndex        =   21
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   7667713
      CurrentDate     =   43052
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   7667713
      CurrentDate     =   43052
   End
   Begin VB.CommandButton cmdReturnBook 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Picture         =   "frmReturnBook.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   5415
      Begin VB.ComboBox cboBookId 
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label12 
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
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label13 
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
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   1128
         Width           =   1215
      End
      Begin VB.Label Label14 
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
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1776
         Width           =   1215
      End
      Begin VB.Label Label15 
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
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2424
         Width           =   1215
      End
      Begin VB.Label Label16 
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
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   3072
         Width           =   1215
      End
      Begin VB.Label Label17 
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
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblBookName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblbookAuthor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   1710
         Width           =   3015
      End
      Begin VB.Label lblbookType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   2340
         Width           =   3015
      End
      Begin VB.Label lblbookCopies 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   5
         Top             =   2970
         Width           =   3015
      End
      Begin VB.Label lblbookPrice 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   210
         Y2              =   4180
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
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1080
      TabIndex        =   18
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
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
      Height          =   495
      Left            =   960
      TabIndex        =   17
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblNoOfBooks 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOOKS TAKEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT ID"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmReturnBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Byte
Dim i As Byte
Private Sub cboBookId_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from stock where bookId = '" & cboBookId.Text & "'", con, 3, 3
    fillBookDetails
End Sub

Private Sub cboStudentId_Click()
    cboBookId.Clear
    temp = 0
    rs.Close
    rs.Open "select bookId from issueBook where studentId = '" & cboStudentId.Text & "'", con, 3, 3
    While Not rs.EOF
        temp = temp + 1
        cboBookId.AddItem rs.Fields!bookId
        rs.MoveNext
    Wend
    lblNoOfBooks.Caption = temp
End Sub


Private Sub cmdReturnBook_Click()
    rs.Close
    rs.Open "select * from returnBook", con, 3, 3
    rs.AddNew
    rs.Fields!studentId = cboStudentId.Text
    rs.Fields!bookId = cboBookId.Text
    rs.Fields!returnDate = DTPicker2.Value
    rs.Update
    rs.Close
    rs.Open "select * from issueBook where (bookId = '" & cboBookId.Text & "') and (studentId = '" & cboStudentId.Text & "')", con, 3, 3
    rs.Delete
    rs.Update
    
    rs.Close
    rs.Open "select * from stati where (bookId = '" & cboBookId.Text & "') and (studentId = '" & cboStudentId.Text & "')", con, 3, 3
    rs.Fields!returned = "yes"
    rs.Fields!returnDate = DTPicker2.Value
    rs.Update
    '-----------------------------------------------------------------------------------------
    rs.Close
    rs.Open "select bookCopies from stock where bookId = '" & cboBookId.Text & "'", con, 3, 3
    temp = rs.Fields!bookCopies
    temp = temp + 1
    rs.Fields!bookCopies = temp
    rs.Update
    lblbookCopies.Caption = temp
    '------------------------------------------------------------------------------------------
    MsgBox ("Book Returned .")
    clearallFields
    Form_Load
End Sub

Private Sub Form_Load()
    x = 0
    temp = 0
    If rs.State = 1 Then rs.Close
    rs.Open "select distinct studentId from issueBook", con, 3, 3
    If rs.BOF = True And rs.EOF = True Then
        MsgBox ("No book is issued.")
    Else
        While Not rs.EOF
            cboStudentId.AddItem rs.Fields!studentId
            rs.MoveNext
        Wend
    End If
End Sub

Private Sub fillBookDetails()
    lblBookName.Caption = rs.Fields(1)
    lblbookAuthor.Caption = rs.Fields(2)
    lblbookType.Caption = rs.Fields(3)
    lblbookCopies.Caption = rs.Fields(4)
    lblbookPrice.Caption = rs.Fields(5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmMain
    If frmadminLogin.lblAdmin.Caption = "true" Then frmadminLogin.adminMethod
    If frmstaffLogin.lblStaff.Caption = "true" Then frmstaffLogin.staffMethod
End Sub

Private Sub clearallFields()
    cboStudentId.Clear
    cboBookId.Clear
    lblBookName.Caption = ""
    lblbookAuthor.Caption = ""
    lblbookType.Caption = ""
    lblbookCopies.Caption = ""
    lblbookPrice.Caption = ""
    lblNoOfBooks.Caption = ""
End Sub
