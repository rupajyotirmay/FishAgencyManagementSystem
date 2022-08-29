VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SellReport 
   BackColor       =   &H8000000E&
   Caption         =   "con.ConnectionString = ""Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\fishAgency\MYDB.mdb;Persist Security Info=False"""
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   10065
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sale Id"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Fish"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CustoName"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Price"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label shopname 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shop Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quantity"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount"
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
      Left            =   8040
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "SellReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\fishAgency\MYDB.mdb;Persist Security Info=False"
con.Open
rs.Open "SaleDetails", con, adOpenDynamic, adLockOptimistic, adCmdTable
End Sub
