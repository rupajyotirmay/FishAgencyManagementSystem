VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3960
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2339.698
   ScaleMode       =   0  'User
   ScaleWidth      =   5999.865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3285
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   870
      TabIndex        =   4
      Top             =   3240
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   3285
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2940
      Left            =   4200
      Picture         =   "frmLogin.frx":0000
      Top             =   840
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Fish Agency Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1425
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
Dim tp As String
Dim rs3 As New ADODB.Recordset
rs3.Open "select * from login where user='" & txtUserName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs3.EOF = False Then
 tp = rs3.Fields(1)
    If txtPassword.Text = tp Then
        LoginSucceeded = True
        Me.Hide
        MDIForm1.Show
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
'        SendKeys "{Home}+{End}"
    End If
    End If
End Sub


'
'Private Sub txtUserName_Change()
'Dim tp As String
'Dim rs3 As New ADODB.Recordset
'rs3.Open "select * from login where user='" & txtUserName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
'If rs3.EOF = False Then
' tp = rs3.Fields(1)
'
'End If
'End Sub
'
'
'
'Private Sub txtUserName_Click()
'txtUserName_Change
'End Sub
Private Sub Image1_Click()

End Sub
