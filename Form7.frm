VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   5430
   Begin VB.Frame Frame1 
      Caption         =   "Stock Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.ListBox List4 
         Height          =   2205
         Left            =   3240
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.ListBox List3 
         Height          =   2205
         Left            =   960
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   4560
         Width           =   855
      End
      Begin VB.ListBox List6 
         Height          =   2205
         Left            =   2040
         TabIndex        =   1
         Top             =   1920
         Width           =   1095
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
         Left            =   960
         TabIndex        =   8
         Top             =   1560
         Width           =   975
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
         Left            =   3240
         TabIndex        =   7
         Top             =   1560
         Width           =   855
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
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   3975
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
         Left            =   2040
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
