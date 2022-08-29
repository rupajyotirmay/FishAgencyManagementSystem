VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form payQry 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Payment View"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Payment View"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7815
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H0080C0FF&
         Caption         =   "Exit"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdview 
         BackColor       =   &H0080C0FF&
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         CurrentDate     =   43101
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4680
         TabIndex        =   2
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         CurrentDate     =   43101
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "To"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Select Date"
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
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1575
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
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   7815
   End
   Begin VB.Label Label6 
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
      TabIndex        =   8
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "payQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdview_Click()
Dim rowctr As Integer
Grid1.Rows = 1
Grid1.Cols = 6
Grid1.TextMatrix(0, 0) = "Bill No"
Grid1.TextMatrix(0, 1) = "Date"
Grid1.TextMatrix(0, 2) = "Dealer"
Grid1.TextMatrix(0, 3) = "TotalAmount"
Grid1.TextMatrix(0, 4) = "PaidAmount"
Grid1.TextMatrix(0, 5) = "DueAmount"
If rs.State = 1 Then rs.Close
rs.Open "select * from  payment where paydate>=#" & DTPicker1.Value & "# and paydate <=#" & DTPicker2.Value & "#", con, adOpenDynamic, adLockOptimistic, adCmdText
rowctr = 1
Do While rs.EOF = False
    Grid1.Rows = Grid1.Rows + 1
    Grid1.TextMatrix(rowctr, 0) = rs.Fields(0)
    Grid1.TextMatrix(rowctr, 1) = rs.Fields(1)
    Grid1.TextMatrix(rowctr, 2) = rs.Fields(2)
    Grid1.TextMatrix(rowctr, 3) = rs.Fields(3)
    Grid1.TextMatrix(rowctr, 4) = IIf(IsNull(rs.Fields(4)), 0, rs.Fields(4))
    Grid1.TextMatrix(rowctr, 5) = IIf(IsNull(rs.Fields(5)), 0, rs.Fields(5))
    rowctr = rowctr + 1
    rs.MoveNext
Loop
End Sub

Private Sub Form_Load()
DTPicker2.Value = Now
Grid1.ColWidth(2) = 1440
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
End Sub

