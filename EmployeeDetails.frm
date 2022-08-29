VERSION 5.00
Begin VB.Form EmployeeDetails 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Details"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame EmployeeDetails 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   8655
      Begin VB.CommandButton cmdadd 
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3840
         Width           =   855
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H0080C0FF&
         Caption         =   "Delete"
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton cmdupdate 
         BackColor       =   &H0080C0FF&
         Caption         =   "Update"
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
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtEmpNm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3600
         TabIndex        =   1
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtEmpId 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3600
         TabIndex        =   0
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtEmpPhone 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox txtEmpAddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3600
         TabIndex        =   3
         Top             =   1920
         Width           =   4455
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000C0&
         Caption         =   " Add"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   " Modify"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   4200
         TabIndex        =   16
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   960
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Shape Shape2 
         Height          =   975
         Left            =   4080
         Top             =   3600
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   " Employee Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   " Employee Id"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080C0FF&
         Caption         =   " Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080C0FF&
         Caption         =   " Phone"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   2520
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
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
      Width           =   8640
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
      Width           =   8640
   End
End
Attribute VB_Name = "EmployeeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset

Private Sub cmdadd_Click()
txtEmpId.Text = ""
txtEmpNm.Text = ""
txtEmpAddress.Text = ""
txtEmpPhone.Text = ""
txtEmpId.SetFocus
End Sub

Private Sub cmddelete_Click()
If MsgBox("Do you really want to delete selected record", vbYesNo + vbQuestion) = vbYes Then
    con.Execute "delete from Employee where EmployeeId='" & txtEmpId.Text & "'"
End If

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub


Private Sub add_Click()

End Sub

Private Sub delete_Click()

End Sub

Private Sub Form_Load()
Dim rs1 As New ADODB.Recordset
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
   lbshopname.Caption = rs1.Fields(0)
End If
rs1.Close
If rs.State = 1 Then rs.Close
rs.Open "Employee", con, adOpenDynamic, adLockOptimistic, adCmdTable

End Sub

Private Sub cmdsave_Click()
rs.AddNew
rs.Fields(0) = txtEmpId.Text
rs.Fields(1) = txtEmpNm.Text
rs.Fields(2) = txtEmpAddress.Text
rs.Fields(3) = txtEmpPhone.Text
rs.update
rs.MoveNext
 MsgBox "Record are saved"
End Sub


Private Sub cmdupdate_Click()
Dim srs As New ADODB.Recordset
If txtEmpId.Text = "" Then Exit Sub
srs.Open "Select * from Employee where EmployeeId='" & txtEmpId.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If srs.EOF = False Then

    srs.Fields(1) = txtEmpNm.Text
    srs.Fields(2) = txtEmpAddress.Text
    srs.Fields(3) = txtEmpPhone.Text
       srs.update
        MsgBox "Record are Updated"
Else
    MsgBox "Record NOt Found to be edited"
End If

End Sub

Private Sub txtEmpId_LostFocus()
 Dim rs2 As New ADODB.Recordset
 If txtEmpId.Text = "" Then Exit Sub
rs2.Open "select  * from Employee where EmployeeId='" & txtEmpId.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs2.EOF = False Then
   txtEmpId.Text = rs2.Fields(0)
   txtEmpNm.Text = rs2.Fields(1)
    txtEmpAddress.Text = rs2.Fields(2)
    txtEmpPhone.Text = rs2.Fields(3)
    
Else
   txtEmpName.Text = ""
    txtEmpcmdAddress.Text = ""
    txtEmpPhone.Text = ""
 End If
rs2.Close
End Sub
