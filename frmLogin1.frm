VERSION 5.00
Begin VB.Form frmLogin1 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   5550
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3279.123
   ScaleMode       =   0  'User
   ScaleWidth      =   6028.032
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Select Login Type  "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      Begin VB.CommandButton cmdProceed 
         BackColor       =   &H0080FFFF&
         Caption         =   "Proceed ->"
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
         Left            =   3953
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdmodify 
         BackColor       =   &H0080FFFF&
         Caption         =   "Modify"
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
         Left            =   2513
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3720
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   2565
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
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
         Left            =   3975
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FFFFFF&
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
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2760
         Width           =   2565
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Login"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3720
         Width           =   1140
      End
      Begin VB.OptionButton OptEmp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OptAdm 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add"
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
         Left            =   1088
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3720
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   3180
         Left            =   3360
         Picture         =   "frmLogin1.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2490
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   " User Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   2055
      End
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
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdAdd_Click()
Dim rs As New ADODB.Recordset
rs.Open "login", con, adOpenDynamic, adLockOptimistic, adCmdTable
rs.AddNew
rs.Fields(0) = txtUserName.Text
rs.Fields(1) = txtPassword.Text
rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Unload Me
    End
End Sub

Private Sub cmdmodify_Click()
Dim srs As New ADODB.Recordset
If txtUserName.Text = "" Then Exit Sub
If txtPassword.Text = "" Then
 MsgBox " Enter Last Password "
End If
srs.Open "Select * from login where user='" & txtUserName.Text & "' and pass='" & txtPassword.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If srs.EOF = False Then

srs.Fields(1) = txtPassword.Text
       srs.update
        MsgBox "Record are Updated"
Else
    MsgBox "Record NOt Found to be edited"
End If

End Sub

Private Sub cmdOK_Click()
Dim tp, id As String
Dim rs3 As New ADODB.Recordset
rs3.Open "select * from login where user='" & txtUserName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs3.EOF = False Then
 tp = rs3.Fields(1)
 id = rs3.Fields(0)
 
     If txtPassword.Text = tp Then
        LoginSucceeded = True
     If (id = "Admin") Then
     frmLogin1.cmdOK.Visible = False
     frmLogin1.cmdCancel.Visible = False
     frmLogin1.cmdadd.Visible = True
     frmLogin1.cmdmodify.Visible = True
     frmLogin1.cmdProceed.Visible = True
     End If
     If Left(id, 2) = "EP" Then
      MDIForm1.mnuMaster.Enabled = False
      MDIForm1.mnuBussiness.Enabled = False
      MDIForm1.mnuQry.Enabled = False
      MDIForm1.mnuReport.Enabled = False
      Unload Me
        MDIForm1.Show
        End If
    Else
    MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
    End If
    End If
End Sub

Private Sub cmdProceed_Click()
     Unload Me
     MDIForm1.Show
End Sub


Private Sub OptAdm_Click()
If OptAdm.Value = True Then
txtUserName.Text = "Admin"
    txtPassword.SetFocus
End If

End Sub

Private Sub OptEmp_Click()
If OptEmp.Value = True Then
txtUserName.Text = "EP"
    txtUserName.SetFocus
End If
End Sub


Private Sub txtUserName_LostFocus()
txtUserName.Text = ""
End Sub
