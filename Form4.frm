VERSION 5.00
Begin VB.Form Coustomer 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Coustomer"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9120
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
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
      Left            =   4200
      TabIndex        =   16
      Top             =   4920
      Width           =   3615
      Begin VB.CommandButton exit 
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
      Begin VB.CommandButton delete 
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton update 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Coustomer Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21
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
      Width           =   8880
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
         Left            =   840
         TabIndex        =   17
         Top             =   3360
         Width           =   2535
         Begin VB.CommandButton add 
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
         Begin VB.CommandButton save 
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
      End
      Begin VB.ComboBox cmbCoustState 
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
         ItemData        =   "Form4.frx":0000
         Left            =   2640
         List            =   "Form4.frx":005B
         TabIndex        =   4
         Top             =   2520
         Width           =   5415
      End
      Begin VB.TextBox txtCoustPh 
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
         Left            =   2640
         TabIndex        =   3
         Top             =   1920
         Width           =   5415
      End
      Begin VB.TextBox txtCoustAdd 
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
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   5415
      End
      Begin VB.TextBox txtCoustName 
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
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   720
         Width           =   1695
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
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
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
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
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
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   2520
         Width           =   1695
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
      Width           =   8880
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
      Width           =   8880
   End
End
Attribute VB_Name = "Coustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rs As New ADODB.Recordset

Private Sub add_Click()
txtCoustName.Text = ""
txtCoustAdd.Text = ""
txtCoustPh.Text = ""
cmbCoustState.Text = ""

txtCoustName.SetFocus
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
If rs.State = 1 Then rs.Close
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
rs.Open "coustomer", con, adOpenDynamic, adLockOptimistic, adCmdTable
End Sub

Private Sub save_Click()
rs.AddNew
rs.Fields(0) = txtCoustName.Text
rs.Fields(1) = txtCoustAdd.Text
rs.Fields(2) = txtCoustPh.Text
rs.Fields(3) = cmbCoustState.Text

rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub


Private Sub txtCoustName_LostFocus()
  Dim rs2 As New ADODB.Recordset
 If txtCoustName.Text = "" Then Exit Sub
rs2.Open "select  * from coustomer where CustomerName='" & txtCoustName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtCoustName.Text = rs2.Fields(0)
   txtCoustAdd.Text = rs2.Fields(1)
    txtCoustPh.Text = rs2.Fields(2)
    cmbCoustState.Text = rs2.Fields(3)

Else
    txtCoustAdd.Text = ""
    txtCoustPh.Text = ""
    cmbCoustState.Text = ""
   
End If
rs2.Close
End Sub

Private Sub update_Click()
Dim srs As New ADODB.Recordset
If txtCoustName.Text = "" Then Exit Sub
srs.Open "Select * from coustomer where CustomerName='" & txtCoustName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If srs.EOF = False Then

srs.Fields(1) = txtCoustAdd.Text
srs.Fields(2) = txtCoustPh.Text
srs.Fields(3) = cmbCoustState.Text

       srs.update
        MsgBox "Record are Updated"
Else
    MsgBox "Record NOt Found to be edited"
End If
End Sub

Private Sub delete_Click()
If MsgBox("Do you really want to delete selected record", vbYesNo + vbQuestion) = vbYes Then
    con.Execute "delete from coustomer where CustomerName='" & txtCoustName.Text & "'"
     MsgBox "Record deleted"
End If
End Sub

