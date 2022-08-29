VERSION 5.00
Begin VB.Form Dealer 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dealer "
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dealer Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   8175
      Begin VB.Frame frModify 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Modify  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1215
         Left            =   3720
         TabIndex        =   17
         Top             =   3480
         Width           =   3615
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
            Left            =   2505
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmddelete 
            BackColor       =   &H0080C0FF&
            Caption         =   "Delete"
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
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdupdate 
            BackColor       =   &H0080C0FF&
            Caption         =   "Update"
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame fradd 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Add  "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1215
         Left            =   600
         TabIndex        =   16
         Top             =   3480
         Width           =   2535
         Begin VB.CommandButton cmdsave 
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
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdadd 
            BackColor       =   &H0080C0FF&
            Caption         =   "New"
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
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtDealerAdd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox txtDealerPh 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox txtDealerName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox cmbDealerState 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Form3.frx":0000
         Left            =   2760
         List            =   "Form3.frx":005B
         TabIndex        =   4
         Top             =   2760
         Width           =   4815
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   " Dealer Name"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   2055
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
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   " Mobile No"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
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
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   2055
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
      TabIndex        =   15
      Top             =   120
      Width           =   8175
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
      TabIndex        =   14
      Top             =   960
      Width           =   8175
   End
End
Attribute VB_Name = "Dealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset

Private Sub cmdAdd_Click()
txtDealerName.Text = ""
txtDealerAdd.Text = ""
txtDealerPh.Text = ""
cmbDealerState.Text = ""

txtDealerName.SetFocus
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub delete_Click()

End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
rs.Open "Dealer", con, adOpenDynamic, adLockOptimistic, adCmdTable
End Sub



Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = txtDealerName.Text
rs.Fields(1) = txtDealerAdd.Text
rs.Fields(2) = txtDealerPh.Text
rs.Fields(3) = cmbDealerState.Text

rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub


Private Sub cmdupdate_Click()
Dim srs As New ADODB.Recordset
If txtDealerName.Text = "" Then Exit Sub
srs.Open "Select * from Dealer where DeaterName='" & txtDealerName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If srs.EOF = False Then

srs.Fields(1) = txtDealerAdd.Text
srs.Fields(2) = txtDealerPh.Text
srs.Fields(3) = cmbDealerState.Text

       srs.update
        MsgBox "Record are Updated"
Else
    MsgBox "Record NOt Found to be edited"
End If
End Sub

Private Sub cmddelete_Click()
If MsgBox("Do you really want to delete selected record", vbYesNo + vbQuestion) = vbYes Then
    con.Execute "delete from Dealer where DeaterName='" & txtDealerName.Text & "'"
     MsgBox "Record deleted"
End If
End Sub


Private Sub txtDealerName_LostFocus()
  Dim rs2 As New ADODB.Recordset
 If txtDealerName.Text = "" Then Exit Sub
rs2.Open "select  * from Dealer where DeaterName='" & txtDealerName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtDealerName.Text = rs2.Fields(0)
   txtDealerAdd.Text = rs2.Fields(1)
    txtDealerPh.Text = rs2.Fields(2)
    cmbDealerState.Text = rs2.Fields(3)

Else
    txtDealerAdd.Text = ""
    txtDealerPh.Text = ""
    cmbDealerState.Text = ""
   
End If
rs2.Close
End Sub
