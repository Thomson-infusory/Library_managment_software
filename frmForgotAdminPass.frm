VERSION 5.00
Begin VB.Form frmForgotAdminPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forgot Password?"
   ClientHeight    =   5205
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmForgotAdminPass.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Height          =   495
      Left            =   7200
      Picture         =   "frmForgotAdminPass.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtAnswer 
      Alignment       =   2  'Center
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
      Left            =   4680
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdFind 
      Height          =   495
      Left            =   4920
      Picture         =   "frmForgotAdminPass.frx":9FBF7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
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
      Left            =   7320
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
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
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      Left            =   2880
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   9015
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
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   720
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
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmForgotAdminPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Byte
Private Sub cmdFind_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "adminLogin", con, 3, 3
    While Not rs.EOF
        If txtName = rs.Fields!Name And txtUsername = rs.Fields!UserName Then
            lblQuestion.Caption = rs.Fields!question
            txtAnswer.Enabled = True
            flag = 1
            GoTo p1
        Else
        rs.MoveNext
        End If
    Wend
    
p1:
    If flag <> 1 Then
         MsgBox ("Invalid Name or Username")
         txtUsername.Text = ""
         txtName.Text = ""
         txtName.SetFocus
    End If
End Sub

Private Sub cmdOk_Click()
    If txtAnswer.Text = rs.Fields!answer Then
        lblPassword.Caption = rs.Fields!Password
        flag = 2
    End If
    If flag <> 2 Then
         MsgBox ("Invalid Name or Username")
         txtAnswer.Text = ""
         txtAnswer.SetFocus
    End If
        
    
End Sub


Private Sub Form_Load()
    flag = 0
    txtAnswer.Enabled = False
    cmdOk.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Picture1.Visible = True
End Sub

Private Sub mnExit_Click()
    Unload Me
    frmMain.Show
End Sub

Private Sub txtAnswer_Change()
    If Len(txtAnswer.Text) = 0 Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub
