VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EmpAtte 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EmpAtte"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5925
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5925
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee Attends"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5655
      Begin VB.ComboBox cmbstatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "EmpAtt.frx":0000
         Left            =   2760
         List            =   "EmpAtt.frx":0010
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H0080C0FF&
         Caption         =   "SAVE"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox cmbid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "EmpAtt.frx":0020
         Left            =   2760
         List            =   "EmpAtt.frx":0022
         TabIndex        =   2
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H0080C0FF&
         Caption         =   "EXIT"
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
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   2760
         TabIndex        =   1
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   115081217
         CurrentDate     =   43384
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   "Employee Id"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
      End
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
      TabIndex        =   10
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label lbshopname 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shop Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5655
   End
End
Attribute VB_Name = "EmpAtte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset

Dim rs2 As New ADODB.Recordset

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
If rs.State = 1 Then rs.Close
rs.Open "Attends", con, adOpenDynamic, adLockOptimistic, adCmdTable
rs2.Open "select EmployeeID from Employee", con, adOpenDynamic, adLockOptimistic, adCmdText
Do While rs2.EOF = False
    cmbid.AddItem rs2.Fields(0)
    rs2.MoveNext
Loop
rs2.Close
End Sub

Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = DTPicker1.Value
rs.Fields(1) = cmbid.Text
rs.Fields(2) = cmbstatus.Text
rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub


