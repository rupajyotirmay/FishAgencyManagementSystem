VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Company 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Com.info"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Company Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5415
      Left            =   135
      TabIndex        =   0
      Top             =   960
      Width           =   11160
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9720
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Attach Photo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   2055
      End
      Begin VB.ComboBox cmbState 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form2.frx":0000
         Left            =   3960
         List            =   "Form2.frx":005B
         TabIndex        =   5
         Top             =   3360
         Width           =   4455
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4448
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5888
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtownerName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   2
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox txtComName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   1
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtComPhone 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2760
         Width           =   4455
      End
      Begin VB.TextBox txtComAddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3960
         TabIndex        =   3
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   " Owner Name"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   " Company Name"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   " Address"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   " Phone"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   10
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   " State"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   720
         TabIndex        =   9
         Top             =   3360
         Width           =   2295
      End
   End
   Begin VB.Label shopname 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Fish Agency"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim picfile As String
Dim rs As New ADODB.Recordset

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
CommonDialog1.ShowOpen
picfile = CommonDialog1.FileName
Image1.Picture = LoadPicture(picfile)
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
rs.Open "Company", con, adOpenDynamic, adLockOptimistic, adCmdTable
Dim rs2 As New ADODB.Recordset
rs2.Open "select  * from company", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtComName.Text = rs2.Fields(0)
   txtownerName.Text = rs2.Fields(1)
    txtComAddress.Text = rs2.Fields(2)
    txtComPhone.Text = rs2.Fields(3)
    cmbState.Text = rs2.Fields(4)
    Image1.Picture = LoadPicture(rs2.Fields(5))
    End If
    rs2.Close
End Sub

Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = txtComName.Text
rs.Fields(1) = txtownerName.Text
rs.Fields(2) = txtComAddress.Text
rs.Fields(3) = txtComPhone.Text
rs.Fields(4) = cmbState.Text
rs.Fields(5) = picfile
rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub

