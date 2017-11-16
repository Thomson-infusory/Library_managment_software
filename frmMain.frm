VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Library Management System"
   ClientHeight    =   13440
   ClientLeft      =   165
   ClientTop       =   -1755
   ClientWidth     =   20955
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   20955
      TabIndex        =   3
      Top             =   14415
      Width           =   20955
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   14415
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   14355
      ScaleWidth      =   20895
      TabIndex        =   4
      Top             =   0
      Width           =   20955
      Begin VB.CommandButton cmdStaff 
         Height          =   375
         Left            =   19080
         Picture         =   "frmMain.frx":98C1E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdmin 
         BackColor       =   &H00C0E0FF&
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
         Left            =   17520
         MaskColor       =   &H00000080&
         Picture         =   "frmMain.frx":9F952
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   18480
         TabIndex        =   12
         Top             =   11400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   7667713
         CurrentDate     =   43052
      End
      Begin VB.CommandButton cmdLogout 
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
         Left            =   19080
         Picture         =   "frmMain.frx":A6763
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox cboType 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   18120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4560
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
         Left            =   7440
         TabIndex        =   0
         Top             =   4440
         Width           =   8175
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   4455
         Left            =   5640
         TabIndex        =   2
         Top             =   6000
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   7858
         _Version        =   393216
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   11760
         TabIndex        =   10
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
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
         Height          =   375
         Left            =   16200
         TabIndex        =   9
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lblNoOfBooks 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   6840
         TabIndex        =   8
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Library Management System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   10680
         TabIndex        =   7
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Books :"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
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
         Left            =   5520
         TabIndex        =   5
         Top             =   4440
         Width           =   1095
      End
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnAdd 
      Caption         =   "&Add"
      Begin VB.Menu mnAddBook 
         Caption         =   "Add Book"
      End
      Begin VB.Menu mnAddStudent 
         Caption         =   "Add Student"
      End
      Begin VB.Menu mnAddStaff 
         Caption         =   "Add Staff"
      End
   End
   Begin VB.Menu mnEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnEditBook 
         Caption         =   "Edit Book"
      End
      Begin VB.Menu mnEditStudent 
         Caption         =   "Edit Student"
      End
      Begin VB.Menu mnEditStaff 
         Caption         =   "Edit Staff"
      End
   End
   Begin VB.Menu mnDelete 
      Caption         =   "&Delete"
      Begin VB.Menu mnDeleteBook 
         Caption         =   "Delete Book "
      End
      Begin VB.Menu mnDeleteStudent 
         Caption         =   "Delete Student "
      End
      Begin VB.Menu mnDeleteStaff 
         Caption         =   "Delete Staff"
      End
   End
   Begin VB.Menu mnIssue 
      Caption         =   "&Issue"
      Begin VB.Menu mnIssueBook 
         Caption         =   "Issue Book"
      End
      Begin VB.Menu mnReturnBook 
         Caption         =   "Return Book"
      End
      Begin VB.Menu mnStatistics 
         Caption         =   "Statistics"
      End
   End
   Begin VB.Menu mnView 
      Caption         =   "&View"
      Begin VB.Menu mnStudents 
         Caption         =   "Students"
      End
      Begin VB.Menu mnStaff 
         Caption         =   "Staff"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

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
    lblNoOfBooks.Caption = (grid.Rows - 1)
End Sub

Private Sub cmdAdmin_Click()
    frmadminLogin.Show
    Picture1.Visible = False
End Sub

Private Sub cmdLogout_Click()
    frmadminLogin.lblAdmin.Caption = "false"
    frmstaffLogin.lblStaff.Caption = "false"
    Unload Me
    frmMain.Show
End Sub

Private Sub cmdStaff_Click()
    Picture1.Visible = False
    frmstaffLogin.Show
End Sub


Private Sub MDIForm_Load()
    If rs.State = 1 Then rs.Close
    rs.Open "stock", con, 3, 3
    
    '-----------------------------------------
    'Disabling some items in the menu bar .
        mnAddBook.Enabled = False
        mnAddStaff.Enabled = False
        mnAddStudent.Enabled = False
        mnDeleteBook.Enabled = False
        mnDeleteStaff.Enabled = False
        mnDeleteStudent.Enabled = False
        mnIssueBook.Enabled = False
        mnReturnBook.Enabled = False
        mnStatistics.Enabled = False
        mnStaff.Enabled = False
        mnStudents.Enabled = False
        mnEditBook.Enabled = False
        mnEditStaff.Enabled = False
        mnEditStudent.Enabled = False
    '-----------------------------------------
    
    If frmadminLogin.lblAdmin.Caption = "true" Then
        '-----------------------------------------
        'Enabling all items in the menu bar.
        frmMain.mnAddBook.Enabled = True
        frmMain.mnAddStaff.Enabled = True
        frmMain.mnAddStudent.Enabled = True
        frmMain.mnDeleteBook.Enabled = True
        frmMain.mnDeleteStaff.Enabled = True
        frmMain.mnDeleteStudent.Enabled = True
        frmMain.mnIssueBook.Enabled = True
        frmMain.mnReturnBook.Enabled = True
        frmMain.mnStatistics.Enabled = True
        frmMain.mnStaff.Enabled = True
        frmMain.mnStudents.Enabled = True
        mnEditBook.Enabled = True
        mnEditStaff.Enabled = True
        mnEditStudent.Enabled = True
        '-------------------------------------------
   End If
        
        If frmstaffLogin.lblStaff.Caption = "true" Then
            '-----------------------------------------
            'Enabling some items in the menu bar.
            frmMain.mnAddBook.Enabled = True
            frmMain.mnAddStudent.Enabled = True
            frmMain.mnDeleteBook.Enabled = True
            frmMain.mnDeleteStudent.Enabled = True
            frmMain.mnIssueBook.Enabled = True
            frmMain.mnReturnBook.Enabled = True
            frmMain.mnStatistics.Enabled = True
            frmMain.mnStaff.Enabled = True
            frmMain.mnStudents.Enabled = True
            '-------------------------------------------
        
        End If
    
    Label5.Visible = False
    cmdLogout.Visible = False
    
    fillheader
    fillgrid
    lblNoOfBooks.Caption = (grid.Rows - 1)
    cboType.AddItem "All"
    rs.Close
    rs.Open "select distinct bookType from stock", con, 3, 3
    While Not rs.EOF
        cboType.AddItem rs.Fields!bookType
        rs.MoveNext
    Wend
    DTPicker1 = Date
    DTPicker1.Enabled = False
    rs.Close
    rs.Open "select * from stock", con, 3, 3
End Sub

'Private Sub MDIForm_Unload(Cancel As Integer)
'    If frmAddBook.Visible = True Then Unload frmAddBook
'    If frmAddStaff.Visible = True Then Unload frmAddStaff
'    If frmAddStudent.Visible = True Then Unload frmAddStudent
'    If frmadminLogin.Visible = True Then Unload frmadminLogin
'    If frmAllStaff.Visible = True Then Unload frmAllStaff
'    If frmAllStudents.Visible = True Then Unload frmAllStudents
'    If frmDeleteBook.Visible = True Then Unload frmDeleteBook
'    If frmDeleteStaff.Visible = True Then Unload frmDeleteStaff
'    If frmDeleteStudent.Visible = True Then Unload frmDeleteStudent
'    If frmEditBook.Visible = True Then Unload frmEditBook
'    If frmEditStaff.Visible = True Then Unload frmEditBook
'    If frmEditStudent.Visible = True Then Unload frmEditStudent
'    If frmForgotAdminPass.Visible = True Then Unload frmForgotAdminPass
'    If frmForgotStaffPass.Visible = True Then Unload frmForgotStaffPass
'    If frmIssueBook.Visible = True Then Unload frmIssueBook
'    If frmReturnBook.Visible = True Then Unload frmReturnBook
'    If frmstaffLogin.Visible = True Then Unload frmstaffLogin
'    If frmStatistics.Visible = True Then Unload frmStatistics
'End Sub

Private Sub mnAddBook_Click()
    Picture1.Visible = False
    frmAddBook.Show
End Sub

Private Sub mnAddStaff_Click()
    Picture1.Visible = False
    frmAddStaff.Show
End Sub

Private Sub mnAddStudent_Click()
    Picture1.Visible = False
    frmAddStudent.Show
End Sub

Private Sub mnDeleteBook_Click()
    Picture1.Visible = False
    frmDeleteBook.Show
End Sub

Private Sub mnDeleteStaff_Click()
    Picture1.Visible = False
    frmDeleteStaff.Show
End Sub

Private Sub mnDeleteStudent_Click()
    Picture1.Visible = False
    frmDeleteStudent.Show
End Sub

Private Sub mnEditBook_Click()
    Picture1.Visible = False
    frmEditBook.Show
End Sub

Private Sub mnEditStaff_Click()
    Picture1.Visible = False
    frmEditStaff.Show
End Sub

Private Sub mnEditStudent_Click()
    Picture1.Visible = False
    frmEditStudent.Show
End Sub

Private Sub mnExit_Click()
    Unload frmMain
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

Private Sub mnIssueBook_Click()
    Picture1.Visible = False
    frmIssueBook.Show
End Sub

Private Sub mnReturnBook_Click()
    Picture1.Visible = False
    frmReturnBook.Show
End Sub

Private Sub mnStaff_Click()
    Picture1.Visible = False
    frmAllStaff.Show
End Sub

Private Sub mnStatistics_Click()
    Picture1.Visible = False
    frmStatistics.Show
End Sub


Private Sub mnStudents_Click()
    Picture1.Visible = False
    frmAllStudents.Show
End Sub

Private Sub txtSearch_Change()
    If rs.State = 1 Then rs.Close
    rs.Open "select * from stock where bookName like '" & txtSearch.Text & "%' Or bookAuthor like '" & txtSearch.Text & "%' Or bookId like '" & txtSearch.Text & "%' Or bookType like '" & txtSearch.Text & "%'"
    fillheader
    fillgrid
    lblNoOfBooks.Caption = (grid.Rows - 1)
End Sub
