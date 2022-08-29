VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4395
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7545
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -120
         Top             =   3960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   " System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   660
         Left            =   5160
         TabIndex        =   7
         Top             =   960
         Width           =   2130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "BRM College ( 2016-2019 )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   2760
         Width           =   2970
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "PARN Group ( BCA Forth Semester )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   2400
         Width           =   4140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Developed by: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5520
         TabIndex        =   4
         Top             =   2040
         Width           =   1725
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Fish Agency Management"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   660
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   7035
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 01.00.18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5760
         TabIndex        =   2
         Top             =   3360
         Width           =   1530
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright Fish Agency Management System. All right reserved."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   3600
         Width           =   5535
      End
      Begin VB.Image imgLogo 
         Height          =   3465
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer

Private Sub Timer1_Timer()
i = i + 1
If i = 3 Then
Me.Hide
    frmLogin1.Show
End If
End Sub
Private Sub Form_Load()
    lblVersion.Caption = "18.10.1"
Dim i As Long
End Sub
