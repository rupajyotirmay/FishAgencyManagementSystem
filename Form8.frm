VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Recipt 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recipt"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7500
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCurTotal 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   5
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Recipt Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7215
      Begin VB.TextBox txtDAmt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   3
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtCAmt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   4
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox cmbPayMode 
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
         ItemData        =   "Form8.frx":0000
         Left            =   3960
         List            =   "Form8.frx":0010
         TabIndex        =   6
         Text            =   " Select Mode"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.TextBox txtPayAmt 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         TabIndex        =   7
         Top             =   4080
         Width           =   2415
      End
      Begin VB.ComboBox cmbpayCoust 
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
         ItemData        =   "Form8.frx":0036
         Left            =   3960
         List            =   "Form8.frx":0038
         TabIndex        =   1
         Text            =   "  Select Name"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4800
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   3960
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111804417
         CurrentDate     =   43371
      End
      Begin VB.Label CurTotal 
         BackColor       =   &H0080C0FF&
         Caption         =   " CurrentTotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   " Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   " Total Credit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   " Total Debit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   " Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   " Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   " Coustomer Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Label lbshopname 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shop Name"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Fish Agency"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Recipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmbpayCoust_Change()
Dim rs2 As New ADODB.Recordset
rs2.Open "select * from recipt where CoustName='" & cmbpayCoust.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
    txtDAmt.Text = rs2.Fields(3)
End If
rs2.Close
rs2.Open "select sum(Gramdtotal) as mytotal  from Sale where CoustName='" & cmbpayCoust.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtCAmt.Text = rs2.Fields(0)
End If
txtCurTotal.Text = txtCAmt.Text - txtDAmt.Text
End Sub

Private Sub cmbpayCoust_Click()
cmbpayCoust_Change
End Sub

Private Sub cmdAdd_Click()
cmbpayCoust.Text = " "
DTPicker1.Value = Now
cmbPayMode.Text = " "
txtPayAmt.Text = " "
txtCAmt.Text = " "
txtDAmt.Text = " "
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = cmbpayCoust.Text
rs.Fields(1) = DTPicker1.Value
rs.Fields(2) = cmbPayMode.Text
rs.Fields(3) = txtPayAmt.Text
rs.Fields(4) = txtCAmt.Text
rs.Fields(5) = txtDAmt.Text
rs.update
rs.MoveNext
MsgBox "Record are saved"
End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
Dim rs2 As New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "recipt", con, adOpenDynamic, adLockOptimistic, adCmdTable
rs2.Open "select distinct CustomerName from coustomer", con, adOpenDynamic, adLockOptimistic, adCmdText
Do While rs2.EOF = False
    cmbpayCoust.AddItem rs2.Fields(0)
    rs2.MoveNext
Loop
rs2.Close
End Sub

Private Sub txtPayAmt_LostFocus()
Dim rs2 As New ADODB.Recordset
rs2.Open "select * from sale where CoustName =' " & cmbpayCoust.Text & " ' ", con, adOpenDynamic, adLockOptimistic, adCmdText
Dim camt As Double
camt = IIf(IsNull(rs.Fields(4)), 0, rs.Fields(4)) - IIf(IsNull(rs.Fields(5)), 0, rs.Fields(5))
txtCAmt.Text = IIf(IsNull(rs.Fields(5)), 0, rs.Fields(5))
txtDAmt.Text = camt
End Sub


