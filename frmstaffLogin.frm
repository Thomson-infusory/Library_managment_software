VERSION 5.00
Begin VB.Form frmstaffLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmstaffLogin.frx":0000
   ScaleHeight     =   2310
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2175
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
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2175
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
      Left            =   960
      Picture         =   "frmstaffLogin.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   2520
      Picture         =   "frmstaffLogin.frx":9FBF7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdForgot 
      Caption         =   "Forgot?"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblStaff 
      Caption         =   "Label3"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
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
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   1455
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
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmstaffLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Byte
Dim counter As Byte
Private Sub cmdCancel_Click()
    Unload Me
    frmMain.Show
End Sub

Private Sub cmdForgot_Click()
    Me.Hide
    frmForgotStaffPass.Show
End Sub

Private Sub cmdOk_Click()

    If rs.State = 1 Then rs.Close
    rs.Open "staffLogin", con, 3, 3

    While Not rs.EOF
        If txtUsername.Text = rs.Fields!UserName And txtPassword.Text = rs.Fields!Password Then
            Dim nameholder As String
            nameholder = rs.Fields!Name
            Me.Hide
            Unload frmMain
            frmMain.Show
            Dim test As Integer
            flag = 1
            
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
            
            '-----------------------------------------------------------
            'Disabling and enabling some labels and command buttons
                frmMain.cmdAdmin.Visible = False
                frmMain.cmdStaff.Visible = False
                frmMain.Label5.Visible = True
                frmMain.cmdLogout.Visible = True
                frmstaffLogin.lblStaff.Caption = "true"
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
        If counter < 3 Then
            MsgBox ("Username or password is incorrect. Try again!")
        Else
            MsgBox ("Too many failed attempts.Program is terminating")
            Unload frmstaffLogin
            frmMain.Show
        End If
    End If
    
End Sub

Private Sub Form_Load()
    lblStaff.Visible = False
    lblStaff.Caption = "false"
    cmdForgot.Visible = False
    flag = 0
    counter = 0
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

Public Sub staffMethod()
frmMain.Show
            Dim test As Integer
            flag = 1
            
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
            
            '-----------------------------------------------------------
            'Disabling and enabling some labels and command buttons
                frmMain.cmdAdmin.Visible = False
                frmMain.cmdStaff.Visible = False
                frmMain.Label5.Visible = True
                frmMain.cmdLogout.Visible = True
                frmstaffLogin.lblStaff.Caption = "true"
                txtUsername.Text = ""
                txtPassword.Text = ""
            '-----------------------------------------------------------
End Sub
