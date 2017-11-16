VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAddStudent 
   Caption         =   "Add Student"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   Picture         =   "frmAddStudent.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optMale 
      BackColor       =   &H00C0FFFF&
      Caption         =   "M"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4320
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00C0FFFF&
      Caption         =   "F"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4800
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   11880
      Picture         =   "frmAddStudent.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3615
      Left            =   1800
      TabIndex        =   13
      Top             =   1800
      Width           =   9235
      _ExtentX        =   16298
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.ComboBox cboSemester 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox cboCourse 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtPhone 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtName 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2190
      TabIndex        =   1
      Top             =   840
      Width           =   1575
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10080
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   240
      Width           =   1695
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
      Left            =   2310
      TabIndex        =   9
      Top             =   240
      Width           =   1335
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
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim temp As Integer
Private Sub cmdAdd_Click()
    If txtId.Text = "" Or txtPhone.Text = "" Or txtName.Text = "" Or cboCourse.Text = "" Or cboSemester.Text = "" Then
        i = 1
        GoTo p1
    End If
    rs.AddNew
        rs.Fields!sId = UCase(txtId.Text)
        rs.Fields!sName = UCase(txtName.Text)
        If optMale.Value = True Then
            rs.Fields!sSex = "M"
        Else
            rs.Fields!sSex = "F"
        End If
        rs.Fields!sCourse = cboCourse.Text
        rs.Fields!sSemester = cboSemester.Text
        rs.Fields!sPhone = UCase(txtPhone.Text)
        rs.Fields!sState = "active"
    rs.Update
    MsgBox ("Student Added")
    grid.Clear
    If rs.State = 1 Then rs.Close
    rs.Open "select * from student where sState = '" & "active" & "'", con, 3, 3
    fillheader
    fillgrid
    clearfields
    txtName.SetFocus
    If i = 1 Then
p1:
        MsgBox ("Fill all fields")
    End If
End Sub

Private Sub clearfields()
    txtId.Text = ""
    txtName.Text = ""
    txtPhone.Text = ""
    optFemale.Value = False
    optMale.Value = False
'---------------------------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from student", con, 3, 3
    rs.MoveLast
    temp = rs.Fields(0)
    txtId.Text = temp + 1
'---------------------------------------------------
    rs.Close
    rs.Open "select * from student where sState = '" & "active" & "'", con, 3, 3
End Sub

Private Sub Form_Load()
'-----------------------------------------------------
    txtId.Enabled = False
    If rs.State = 1 Then rs.Close
    rs.Open "select * from student", con, 3, 3
    If rs.BOF = True And rs.EOF = True Then
        txtId.Text = "25000"
    Else
    rs.MoveLast
    temp = rs.Fields(0)
    txtId.Text = temp + 1
    End If
'-----------------------------------------------------
'----------------------------------------------------------------
'--------------------------------------
    cboCourse.AddItem "BCA"
    cboCourse.AddItem "BBA"
    cboCourse.AddItem "BCOM"
    cboCourse.AddItem "BA ENGLISH"
    cboCourse.AddItem "BA ELECTRONICS"
    cboCourse.AddItem "MA ELECTRONICS"
    cboCourse.AddItem "MCOM"
    cboCourse.AddItem "MSW"
'--------------------------------------
'--------------------------------------
    cboSemester.AddItem "1"
    cboSemester.AddItem "2"
    cboSemester.AddItem "3"
    cboSemester.AddItem "4"
    cboSemester.AddItem "5"
    cboSemester.AddItem "6"
'--------------------------------------
'----------------------------------------------------------------
    rs.Close
    rs.Open "select * from student where sState = '" & "active" & "'", con, 3, 3
    fillheader
    fillgrid
End Sub


Private Sub fillheader()
        grid.Rows = 1
        grid.Cols = 6
        grid.ColAlignment(0) = 4
        grid.ColAlignment(3) = 1
        grid.ColAlignment(4) = 1
        grid.ColAlignment(5) = 1
        grid.ColWidth(1) = 2200
        grid.ColWidth(3) = 1600
        grid.ColWidth(4) = 1600
        grid.ColWidth(5) = 1800
        grid.TextMatrix(0, 0) = "ID"
        grid.TextMatrix(0, 1) = "FULL NAME"
        grid.TextMatrix(0, 2) = "SEX"
        grid.TextMatrix(0, 3) = "COURSE"
        grid.TextMatrix(0, 4) = "SEMESTER"
        grid.TextMatrix(0, 5) = "PHONE"
End Sub

Private Sub fillgrid()
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


Private Sub txtName_Change()
    txtName.MaxLength = 25
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPhone_Change()
    txtPhone.MaxLength = 10
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
