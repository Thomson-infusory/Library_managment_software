VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddStaff 
   Caption         =   "Add Staff"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   Picture         =   "frmAddStaff.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00C0FFFF&
      Caption         =   "F"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4800
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00C0FFFF&
      Caption         =   "M"
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4320
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   11760
      TabIndex        =   14
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   7667713
      CurrentDate     =   43052
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   615
      Left            =   13560
      Picture         =   "frmAddStaff.frx":98C1E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   855
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
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtPlace 
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
      Left            =   7680
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtAge 
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
      Left            =   6000
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
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
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
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
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3615
      Left            =   1440
      TabIndex        =   13
      Top             =   2040
      Width           =   11575
      _ExtentX        =   20426
      _ExtentY        =   6376
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Join"
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
      Left            =   11640
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label6 
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
      Left            =   9720
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Place"
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
      Left            =   7755
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   5940
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   4125
      TabIndex        =   2
      Top             =   480
      Width           =   1335
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
      Left            =   2055
      TabIndex        =   1
      Top             =   480
      Width           =   1575
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmAddStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim temp As Integer
Private Sub cmdAdd_Click()
    If txtId.Text = "" Or txtName.Text = "" Or txtAge.Text = "" Or txtPlace.Text = "" Or txtPhone.Text = "" = True Then
        i = 1
        GoTo p1
    End If
    rs.AddNew
        rs.Fields!stId = UCase(txtId.Text)
        rs.Fields!stName = UCase(txtName.Text)
        If optMale.Value = True Then
            rs.Fields!stSex = "M"
        Else
            rs.Fields!stSex = "F"
        End If
        rs.Fields!stAge = txtAge.Text
        rs.Fields!stPlace = UCase(txtPlace.Text)
        rs.Fields!stPhone = txtPhone.Text
        rs.Fields!stDoj = DTPicker1.Value
        rs.Fields!stState = "active"
    rs.Update
    MsgBox ("Staff Added")
    grid.Clear
    If rs.State = 1 Then rs.Close
    rs.Open "select * from staff where stState = '" & "active" & "'", con, 3, 3
    fillheader
    fillgrid
    clearfields
    txtName.SetFocus
    If i = 1 Then
p1:
        MsgBox ("Fill all fields")
    End If
End Sub

Private Sub Form_Load()
'-----------------------------------------------------
    txtId.Enabled = False
    If rs.State = 1 Then rs.Close
    rs.Open "select stId from staff", con, 3, 3
    If rs.BOF = True And rs.EOF = True Then
        txtId.Text = "1"
    Else
    rs.MoveLast
    temp = rs.Fields(0)
    txtId.Text = temp + 1
    End If
'-----------------------------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from staff where stState = '" & "active" & "'", con, 3, 3
    fillheader
    fillgrid
End Sub

Private Sub fillheader()
        grid.Rows = 1
        grid.Cols = 7
        grid.ColAlignment(0) = 4
        grid.ColAlignment(1) = 1
        grid.ColAlignment(3) = 1
        grid.ColAlignment(4) = 1
        grid.ColAlignment(5) = 1
        grid.ColAlignment(6) = 1
        grid.ColWidth(1) = 2200
        grid.ColWidth(2) = 1600
        grid.ColWidth(3) = 1600
        grid.ColWidth(4) = 1800
        grid.ColWidth(5) = 1800
        grid.ColWidth(6) = 1500
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


Private Sub clearfields()
    txtId.Text = ""
    txtName.Text = ""
    txtPhone.Text = ""
    optFemale.Value = False
    optMale.Value = False
    txtAge.Text = ""
    txtPlace.Text = ""
    DTPicker1.Value = Date
'---------------------------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select * from staff", con, 3, 3
    rs.MoveLast
    temp = rs.Fields(0)
    txtId.Text = temp + 1
'---------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Picture1.Visible = True
End Sub



Private Sub txtAge_Change()
    txtAge.MaxLength = 2
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_Change()
    txtName.MaxLength = 30
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

Private Sub txtPlace_Change()
    txtPlace.MaxLength = 35
End Sub

Private Sub txtPlace_KeyPress(KeyAscii As Integer)
    If (KeyAscii > 47 And KeyAscii < 58) Then
        KeyAscii = 0
    End If
End Sub
