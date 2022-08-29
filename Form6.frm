VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Payment 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Payment"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6735
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
      Left            =   3480
      TabIndex        =   5
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Payment Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6495
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   3360
         TabIndex        =   2
         Top             =   1320
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
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
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
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5040
         Width           =   975
      End
      Begin VB.ComboBox cmbPayDealer 
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
         ItemData        =   "Form6.frx":0000
         Left            =   3360
         List            =   "Form6.frx":0002
         TabIndex        =   1
         Text            =   "  Select Name"
         Top             =   720
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
         Left            =   3360
         TabIndex        =   7
         Top             =   4320
         Width           =   2415
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
         ItemData        =   "Form6.frx":0004
         Left            =   3360
         List            =   "Form6.frx":0014
         TabIndex        =   6
         Text            =   " Select Mode"
         Top             =   3720
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
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
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5040
         Width           =   1095
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
         Left            =   3360
         TabIndex        =   4
         Top             =   2520
         Width           =   2415
      End
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
         Left            =   3360
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   " Current Total"
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
         Left            =   600
         TabIndex        =   19
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   " Deler Name"
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
         Left            =   600
         TabIndex        =   16
         Top             =   720
         Width           =   1815
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
         Left            =   600
         TabIndex        =   15
         Top             =   4320
         Width           =   1815
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
         Left            =   600
         TabIndex        =   14
         Top             =   3720
         Width           =   1815
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
         Left            =   600
         TabIndex        =   13
         Top             =   1920
         Width           =   1815
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
         Left            =   600
         TabIndex        =   12
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   " Date"
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
         Left            =   600
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
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
      Width           =   6495
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
      Width           =   6495
   End
End
Attribute VB_Name = "Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmbPayDealer_Change()
Dim rs2 As New ADODB.Recordset
rs2.Open "select * from payment where DealerName='" & cmbPayDealer.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtCAmt.Text = rs2.Fields(3)
End If
rs2.Close
rs2.Open "select sum(GrandTotal) as mytotal  from Purchase where DealerName='" & cmbPayDealer.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtDAmt.Text = IIf(IsNull(rs2.Fields(0)), 0, rs2.Fields(0))
End If
txtCurTotal.Text = txtDAmt.Text - txtCAmt.Text
End Sub

Private Sub cmbPayDealer_Click()
cmbPayDealer_Change
End Sub

Private Sub cmdAdd_Click()
cmbPayDealer.Text = " "
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
rs.Fields(0) = cmbPayDealer.Text
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
rs.Open "payment", con, adOpenDynamic, adLockOptimistic, adCmdTable
rs2.Open "select distinct DeaterName from Dealer", con, adOpenDynamic, adLockOptimistic, adCmdText
Do While rs2.EOF = False
    cmbPayDealer.AddItem rs2.Fields(0)
    rs2.MoveNext
Loop
rs2.Close
End Sub



