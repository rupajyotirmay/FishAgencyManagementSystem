VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Purchase Report"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9960
   Begin VB.Frame Frame1 
      Caption         =   "Purchase Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      Begin VB.ListBox List7 
         Height          =   2205
         Left            =   6720
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.ListBox List6 
         Height          =   2205
         Left            =   5400
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Height          =   2205
         Left            =   7920
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   2205
         Left            =   4440
         TabIndex        =   5
         Top             =   2160
         Width           =   735
      End
      Begin VB.ListBox List3 
         Height          =   2205
         Left            =   3240
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   1800
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   480
         TabIndex        =   2
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   5040
         Width           =   855
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
         Left            =   6720
         TabIndex        =   16
         Top             =   1560
         Width           =   975
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
         Left            =   5400
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bill No"
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
         Top             =   1560
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
         Left            =   3360
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DealerName"
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
         Left            =   1680
         TabIndex        =   10
         Top             =   1560
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
         Left            =   4440
         TabIndex        =   9
         Top             =   1560
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
         Left            =   7920
         TabIndex        =   8
         Top             =   1560
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
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

End Sub
