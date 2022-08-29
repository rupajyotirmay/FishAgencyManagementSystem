VERSION 5.00
Begin VB.Form EditFishDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Fish Details"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7485
   ForeColor       =   &H00400040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame EditFishDetailsEditFishDetails 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Fish Details"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   7215
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
         Left            =   622
         TabIndex        =   13
         Top             =   2520
         Width           =   2415
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
            TabIndex        =   15
            Top             =   480
            Width           =   855
         End
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
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   480
            Width           =   855
         End
      End
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
         Left            =   3247
         TabIndex        =   9
         Top             =   2520
         Width           =   3615
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
            TabIndex        =   12
            Top             =   480
            Width           =   975
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
            TabIndex        =   11
            Top             =   480
            Width           =   855
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
            Left            =   2505
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox txtFishName 
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
         Left            =   3720
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtqtyType 
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
         Left            =   3720
         TabIndex        =   3
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtfishtype 
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
         Left            =   3720
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   " Fish Name"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   " Fish Type"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   " Qty Type"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
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
      TabIndex        =   8
      Top             =   960
      Width           =   7215
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
      TabIndex        =   7
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "EditFishDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset


Private Sub cmdadd_Click()
txtFishName.Text = ""
txtfishtype.Text = ""
txtqtyType.Text = ""
Text1.SetFocus
End Sub


Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = txtFishName.Text
rs.Fields(1) = txtfishtype.Text
rs.Fields(2) = txtqtyType.Text
rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub

Private Sub cmdupdate_Click()
Dim srs As New ADODB.Recordset
If txtFishName.Text = "" Then Exit Sub
srs.Open "Select * from Fish where FishName='" & txtFishName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If srs.EOF = False Then

 srs.Fields(1) = txtfishtype.Text
srs.Fields(2) = txtqtyType.Text
       srs.update
        MsgBox "Record are Updated"
Else
    MsgBox "Record NOt Found to be edited"
End If
End Sub

Private Sub Form_Load()
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
If rs.State = 1 Then rs.Close
rs.Open "Fish", con, adOpenDynamic, adLockOptimistic, adCmdTable
End Sub


Private Sub cmddelete_Click()
If MsgBox("Do you really want to delete selected record", vbYesNo + vbQuestion) = vbYes Then
    con.Execute "delete from Fish where FishName='" & txtFishName.Text & "'"
     MsgBox "Record deleted"
End If
End Sub


Private Sub txtFishName_LostFocus()
  Dim rs2 As New ADODB.Recordset
 If txtFishName.Text = "" Then Exit Sub
rs2.Open "select  * from Fish where FishName='" & txtFishName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtFishName.Text = rs2.Fields(0)
   txtfishtype.Text = rs2.Fields(1)
   txtqtyType.Text = rs2.Fields(2)
Else
   txtfishtype.Text = ""
   txtqtyType = ""
End If
rs2.Close
End Sub
