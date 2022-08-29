VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SaleDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sale Details"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10350
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Sales Deatails"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10095
      Begin VB.CommandButton cmdnew 
         BackColor       =   &H0080C0FF&
         Caption         =   "New"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7320
         Width           =   855
      End
      Begin VB.TextBox txtQty 
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
         Left            =   4920
         TabIndex        =   6
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "ADD"
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
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtPrice 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtPdamt 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   10
         Top             =   6600
         Width           =   2175
      End
      Begin VB.TextBox txtTamt 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   19
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox txtDis 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   9
         Top             =   5640
         Width           =   2175
      End
      Begin VB.TextBox txtBillNo 
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
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cmbName 
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
         Left            =   3000
         TabIndex        =   3
         Top             =   1200
         Width           =   6375
      End
      Begin VB.ComboBox cmbFishName 
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
         ItemData        =   "SaleDetails.frx":0000
         Left            =   720
         List            =   "SaleDetails.frx":0040
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cmbFishQty 
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
         ItemData        =   "SaleDetails.frx":00EC
         Left            =   3360
         List            =   "SaleDetails.frx":00FC
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtIGST 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   17
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox txtCGST 
         Alignment       =   1  'Right Justify
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
         Left            =   7320
         TabIndex        =   16
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox txtSGST 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   15
         Top             =   5640
         Width           =   1935
      End
      Begin VB.TextBox txtgndtotal 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         TabIndex        =   13
         Top             =   6120
         Width           =   2175
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
         Left            =   4365
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7320
         Width           =   855
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   5565
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7320
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MMM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   2
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111804419
         CurrentDate     =   43370
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   2295
         Left            =   600
         TabIndex        =   18
         Top             =   2640
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "PRICE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   32
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   31
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "FISH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   30
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Paid Amount"
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
         TabIndex        =   29
         Top             =   6600
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         Caption         =   "Total Amount"
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
         TabIndex        =   28
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Discount"
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
         TabIndex        =   27
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H0080C0FF&
         Caption         =   "Coustomer Name"
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
         TabIndex        =   26
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         Caption         =   "Date "
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
         Left            =   6000
         TabIndex        =   25
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
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
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label LbSGST 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "SGST"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   23
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label LbCGST 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "CGST"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   22
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label LbIGST 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "IGST"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   21
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         Caption         =   "Grand Total"
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
         TabIndex        =   20
         Top             =   6120
         Width           =   1695
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
      TabIndex        =   34
      Top             =   960
      Width           =   10095
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
      TabIndex        =   33
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "SaleDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prs As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub cmbName_LostFocus()
Dim rs1 As New ADODB.Recordset
rs1.Open "select * from coustomer where CustomerName='" & cmbName.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs1.Fields(3) = "Bihar" Then
    LbIGST.Visible = False
    txtIGST.Visible = False
Else
LbCGST.Visible = False
txtCGST.Visible = False

LbSGST.Visible = False
txtSGST.Visible = False
End If
rs1.Close
End Sub

Private Sub cmdAdd_Click()
Dim tamt As Double
If IsNumeric(txtQty.Text) = False Then
    MsgBox "Please check Qty can not be zero", vbCritical
    Exit Sub
End If
If Val(txtQty.Text) = 0 Then
    MsgBox "Please check Qty can not be zero", vbCritical
    Exit Sub
End If
If Val(txtPrice.Text) = 0 Then
    MsgBox "Please check Price can not be zero", vbCritical
    Exit Sub
End If
Grid1.Rows = Grid1.Rows + 1
Grid1.TextMatrix(Grid1.Rows - 1, 0) = Grid1.Rows - 1
Grid1.TextMatrix(Grid1.Rows - 1, 1) = cmbFishName.Text
Grid1.TextMatrix(Grid1.Rows - 1, 2) = cmbFishQty.Text
Grid1.TextMatrix(Grid1.Rows - 1, 3) = txtQty.Text
Grid1.TextMatrix(Grid1.Rows - 1, 4) = txtPrice.Text
Grid1.TextMatrix(Grid1.Rows - 1, 5) = Val(txtQty.Text) * Val(txtPrice.Text)
 For i = 1 To Grid1.Rows - 1
    tamt = tamt + Val(Grid1.TextMatrix(i, 5))
Next i
txtTamt.Text = tamt
txtIGST.Text = txtTamt.Text * 5 / 100
txtCGST.Text = txtTamt.Text * 2.5 / 100
txtSGST.Text = txtTamt.Text * 2.5 / 100
cmbFishName.ListIndex = -1
cmbFishQty.ListIndex = -1
txtQty.Text = ""
txtPrice.Text = ""
txtgndtotal.Text = Val(txtTamt.Text) + Val(txtCGST.Text) + Val(txtSGST.Text) + Val(txtIGST.Text) - (Val(txtTamt.Text) + Val(txtCGST.Text) + Val(txtSGST.Text) + Val(txtIGST.Text))
cmbFishName.SetFocus
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Grid1.Rows = 1
txtBillNo.Text = ""
DTPicker1.Value = Now
cmbName.Text = ""
txtIGST.Text = ""
 txtSGST.Text = ""
 txtCGST.Text = ""
 txtDis.Text = ""
 txtTamt.Text = ""
 txtgndtotal.Text = ""
 txtPdamt.Text = ""
End Sub

Private Sub cmdsave_Click()
  Dim bill As String
  Dim trs As New Recordset
  trs.Open "select Max(Right(BillNo,4))from sale", con, adOpenDynamic, adLockOptimistic, adCmdText
  bill = IIf(IsNull(trs.Fields(0)), 1, trs.Fields(0) + 1)
  bill = "FA" & Format(bill, "0000")
Dim rs3 As New ADODB.Recordset
rs.AddNew
rs.Fields(0) = bill
rs.Fields(1) = DTPicker1.Value
rs.Fields(2) = cmbName.Text
rs.Fields(3) = txtIGST
rs.Fields(4) = txtSGST
rs.Fields(5) = txtCGST
rs.Fields(6) = txtDis.Text
rs.Fields(7) = txtTamt.Text
rs.Fields(8) = txtgndtotal.Text
rs.Fields(9) = Val(txtPdamt.Text)
rs.update
rs.MoveNext
 For i = 1 To Grid1.Rows - 1
    prs.AddNew
    prs.Fields(0) = txtBillNo.Text
    prs.Fields(1) = Grid1.TextMatrix(i, 1)
    prs.Fields(2) = Grid1.TextMatrix(i, 2)
    prs.Fields(3) = Grid1.TextMatrix(i, 3)
    prs.Fields(4) = Grid1.TextMatrix(i, 4)
    prs.update
    If rs3.State = 1 Then rs3.Close
    rs3.Open "select * from Fish where FishName='" & Grid1.TextMatrix(i, 1) & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
    If rs3.EOF = False Then
        rs3.Fields(3) = rs3.Fields(3) - Val(Grid1.TextMatrix(i, 3))
        rs3.update
    End If
Next i
    If rs3.State = 1 Then rs3.Close
    rs3.Open "Recipt", con, adOpenDynamic, adLockOptimistic, adCmdTable
    rs3.AddNew
    rs3.Fields(0) = cmbName.Text
    rs3.Fields(1) = DTPicker1.Value
    rs3.Fields(2) = "Cash"
    rs3.Fields(3) = Val(txtPdamt.Text)
    rs3.update
    rs3.Close
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
If rs.State = 1 Then rs.Close
rs.Open "sale", con, adOpenDynamic, adLockOptimistic, adCmdTable
rs2.Open "select distinct CustomerName from coustomer", con, adOpenDynamic, adLockOptimistic, adCmdText
Do While rs2.EOF = False
    cmbName.AddItem rs2.Fields(0)
    rs2.MoveNext
Loop
rs2.Close
rs2.Open "select distinct * from Fish", con, adOpenDynamic, adLockOptimistic, adCmdText
Do While rs2.EOF = False
    cmbFishName.AddItem rs2.Fields(0)
    cmbFishQty.AddItem rs2.Fields(2)
    rs2.MoveNext
Loop
rs2.Close
If prs.State = 1 Then prs.Close
prs.Open "saledetails", con, adOpenDynamic, adLockOptimistic, adCmdTable
Grid1.Cols = 6
Grid1.TextMatrix(0, 0) = "sr."
Grid1.TextMatrix(0, 1) = "Item"
Grid1.TextMatrix(0, 2) = "Qty Type"
Grid1.TextMatrix(0, 3) = "Qty"
Grid1.TextMatrix(0, 4) = "Rate"
Grid1.TextMatrix(0, 5) = "Amount"
End Sub

Private Sub txtDis_Change()
txtgndtotal.Text = Val(txtTamt.Text) + Val(txtCGST.Text) + Val(txtSGST.Text) + Val(txtIGST.Text) - (Val(txtTamt.Text) + Val(txtCGST.Text) + Val(txtSGST.Text) + Val(txtIGST.Text)) * Val(txtDis.Text) / 100
End Sub

Private Sub txtQty_LostFocus()
Dim rsf As New ADODB.Recordset
Dim pv As Integer
rsf.Open "select * from Fish where FishName='" & cmbFishName.Text & "' ", con, adOpenDynamic, adLockOptimistic, adCmdText
If rsf.EOF = True Then
    MsgBox "iTEM NOT FOUND IN STOCK"
Else
    If rsf.Fields(3) <= txtQty Then
        MsgBox "stock is only " & rsf.Fields(3) & "."
    Else
        If IsNull(rsf.Fields(3)) Or rsf.Fields(3) = 0 Then
            MsgBox "sorry zero qty available"
        End If
    End If
End If
End Sub
