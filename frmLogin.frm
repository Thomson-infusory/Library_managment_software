VERSION 5.00
Begin VB.Form frmadminLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2370
   ClientLeft      =   7860
   ClientTop       =   7440
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2370
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdForgot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Picture         =   "frmLogin.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Picture         =   "frmLogin.frx":A00FF
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Picture         =   "frmLogin.frx":A7447
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblAdmin 
      Caption         =   "Label3"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmadminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Byte
Dim flag As Byte
Public admin As Boolean
Private Sub cmdCancel_Click()
    Unload Me
    frmMain.Show
End Sub

Private Sub cmdForgot_Click()
    Me.Hide
    frmForgotAdminPass.Show
End Sub

Private Sub cmdOk_Click()

    If rs.State = 1 Then rs.Close
    rs.Open "adminLogin", con, 3, 3
    
    While Not rs.EOF
        If txtUsername.Text = rs.Fields!UserName And txtPassword.Text = rs.Fields!Password Then
            Dim nameholder As String
            nameholder = rs.Fields!Name
            Me.Hide
            Unload frmMain
            frmMain.Show
            flag = 1
            
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
                frmMain.mnEditBook.Enabled = True
                frmMain.mnEditStudent.Enabled = True
                frmMain.mnEditStaff.Enabled = True
            '-------------------------------------------
            
            '-----------------------------------------------------------
            'Disabling and enabling some labels and command buttons
                frmMain.cmdAdmin.Visible = False
                frmMain.cmdStaff.Visible = False
                frmMain.Label5.Visible = True
                frmMain.cmdLogout.Visible = True
                lblAdmin.Caption = "true"
                txtUsername.Text = ""
                txtPassword.Text = ""
                frmMain.Label5.Caption = "Welcome " & nameholder
            '-----------------------------------------------------------
            GoTo p1
                
         Else
            rs.MoveNext
            flag = 0
        End If
    Wend
    
p1:
    
    If flag = 0 Then
        counter = counter + 1
        cmdCancel.Visible = False
        cmdForgot.Visible = True
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.SetFocus
        If counter < 3 Then
            MsgBox ("Username or password is incorrect. Try again!")
        Else
            MsgBox ("Too many failed attempts.Program is terminating")
            Unload frmadminLogin
            frmMain.Show
        End If
    End If

End Sub



Private Sub Form_Load()
    lblAdmin.Visible = False
    lblAdmin.Caption = "false"
    cmdForgot.Visible = False
    counter = 0
    flag = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain.Picture1.Visible = True
End Sub

Private Sub txtPassword_Change()
    txtPassword.MaxLength = 8
End Sub

Private Sub txtUsername_Change()
    txtUsername.MaxLength = 10
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 64 Or KeyAscii < 123 And KeyAscii = 8) Then
        KeyAscii = 0
    End If
        
End Sub

Public Sub adminMethod()
            frmMain.Show
            flag = 1
            
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
                frmMain.mnEditBook.Enabled = True
                frmMain.mnEditStudent.Enabled = True
                frmMain.mnEditStaff.Enabled = True
            '-------------------------------------------
            
            '-----------------------------------------------------------
            'Disabling and enabling some labels and command buttons
                frmMain.cmdAdmin.Visible = False
                frmMain.cmdStaff.Visible = False
                frmMain.Label5.Visible = True
                frmMain.cmdLogout.Visible = True
                lblAdmin.Caption = "true"
                txtUsername.Text = ""
                txtPassword.Text = ""
            '-----------------------------------------------------------
End Sub
