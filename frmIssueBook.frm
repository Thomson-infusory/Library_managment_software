VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIssueBook 
   Caption         =   "Issue Book"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   Picture         =   "frmIssueBook.frx":0000
   ScaleHeight     =   6840
   ScaleWidth      =   11955
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   6000
      TabIndex        =   28
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   7667713
      CurrentDate     =   43052
   End
   Begin VB.CommandButton cmdIssueBook 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Picture         =   "frmIssueBook.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Student"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   360
      Width           =   5415
      Begin VB.ComboBox cboStudentId 
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
         TabIndex        =   14
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label11 
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
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         TabIndex        =   24
         Top             =   1128
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
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
         TabIndex        =   23
         Top             =   1776
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
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
         TabIndex        =   22
         Top             =   2424
         Width           =   1215
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Semester"
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
         TabIndex        =   21
         Top             =   3072
         Width           =   1215
      End
      Begin VB.Label label67 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
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
         TabIndex        =   20
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblStudentName 
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
         TabIndex        =   19
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblStudentSex 
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
         TabIndex        =   18
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lblStudentCourse 
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
         TabIndex        =   17
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label lblStudentSemester 
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
         TabIndex        =   16
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label lblStudentPhone 
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
         TabIndex        =   15
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   210
         Y2              =   4180
      End
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
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
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   210
         Y2              =   4180
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
         TabIndex        =   12
         Top             =   3600
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
         TabIndex        =   11
         Top             =   2970
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
         TabIndex        =   10
         Top             =   2340
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
         TabIndex        =   9
         Top             =   1710
         Width           =   3015
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
         TabIndex        =   6
         Top             =   3720
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
         TabIndex        =   5
         Top             =   3072
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
         TabIndex        =   4
         Top             =   2424
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
         TabIndex        =   3
         Top             =   1776
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
         TabIndex        =   2
         Top             =   1128
         Width           =   1215
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
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
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
      Left            =   3840
      TabIndex        =   26
      Top             =   5040
      Width           =   1815
   End
End
Attribute VB_Name = "frmIssueBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Byte
Dim temp As Integer


Private Sub cboBookId_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from stock where bookId = '" & cboBookId.Text & "'", con, 3, 3
    fillBookDetails
    If lblbookCopies.Caption = "0" Then
        MsgBox ("No Copies are available. Select a different book . ")
        clearBookDetails
    End If
End Sub

Private Sub cboStudentId_Click()
        If rs.State = 1 Then rs.Close
        rs.Open "select * from student where sId = '" & cboStudentId.Text & "'", con, 3, 3
        fillStudentDetails
End Sub

Private Sub cmdIssueBook_Click()
    temp = 0
    x = 0
    If lblBookName.Caption = "" Or lblStudentName.Caption = "" Then
        MsgBox ("Select both Book ID and Student ID .")
        cboBookId.SetFocus
    Else
        If rs.State = 1 Then rs.Close
        rs.Open "select * from issueBook ", con, 3, 3
        While Not rs.EOF
            If cboBookId.Text = rs.Fields!bookId And cboStudentId.Text = rs.Fields!studentId Then
                temp = 1
            End If
            rs.MoveNext
        Wend
        
        If temp = 1 Then
            MsgBox ("Same book has been issued to the student .")
        Else
            If rs.State = 1 Then rs.Close
            rs.Open "select * from issueBook where studentId = '" & cboStudentId.Text & "'", con, 3, 3
            While Not rs.EOF
                x = x + 1
                rs.MoveNext
            Wend
            If x >= 5 Then
                MsgBox ("Maximum number of book has been issued to the student . ")
            Else
                rs.Close
                rs.Open "select * from issueBook", con, 3, 3
                rs.AddNew
                rs.Fields!bookId = cboBookId.Text
                rs.Fields!studentId = cboStudentId.Text
                rs.Fields!issueDate = DTPicker1.Value
                rs.Update
                rs.Close
                rs.Open "select bookCopies from stock where bookId = '" & cboBookId.Text & "'", con, 3, 3
                temp = rs.Fields!bookCopies
                temp = temp - 1
                rs.Fields!bookCopies = temp
                rs.Update
                lblbookCopies.Caption = temp
                MsgBox ("Book Issued.")
                
                rs.Close
                rs.Open "select * from stati", con, 3, 3
                rs.AddNew
                rs.Fields!bookId = cboBookId.Text
                rs.Fields!bookName = lblBookName.Caption
                rs.Fields!studentId = cboStudentId.Text
                rs.Fields!issueDate = DTPicker1.Value
                rs.Fields!returned = "no"
                rs.Update
            End If
        End If
      End If
End Sub

Private Sub Form_Load()
    temp = 0
    x = 0
    If rs.State = 1 Then rs.Close
    rs.Open "select bookId from stock", con, 3, 3
    While Not rs.EOF
        cboBookId.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
    rs.Close
    rs.Open "select SId from student where sState = 'active' ", con, 3, 3
    While Not rs.EOF
        cboStudentId.AddItem rs.Fields(0)
        rs.MoveNext
    Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmMain
    If frmadminLogin.lblAdmin.Caption = "true" Then frmadminLogin.adminMethod
    If frmstaffLogin.lblStaff.Caption = "true" Then frmstaffLogin.staffMethod
End Sub

Private Sub fillBookDetails()
    lblBookName.Caption = rs.Fields(1)
    lblbookAuthor.Caption = rs.Fields(2)
    lblbookType.Caption = rs.Fields(3)
    lblbookCopies.Caption = rs.Fields(4)
    lblbookPrice.Caption = rs.Fields(5)
End Sub

Private Sub fillStudentDetails()
    lblStudentName.Caption = rs.Fields(1)
    lblStudentSex.Caption = rs.Fields(2)
    lblStudentCourse.Caption = rs.Fields(3)
    lblStudentSemester.Caption = rs.Fields(4)
    lblStudentPhone.Caption = rs.Fields(5)
End Sub



Private Sub clearBookDetails()
    lblBookName.Caption = ""
    lblbookAuthor.Caption = ""
    lblbookType.Caption = ""
    lblbookCopies.Caption = ""
    lblbookPrice.Caption = ""
End Sub
