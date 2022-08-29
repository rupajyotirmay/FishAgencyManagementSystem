VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Fish Agency"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnucominfo 
         Caption         =   "&Company Info"
      End
      Begin VB.Menu mnuFish 
         Caption         =   "Fish"
      End
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employee"
      End
      Begin VB.Menu mnuDealer 
         Caption         =   "&Dealer"
      End
      Begin VB.Menu mnuCoustomer 
         Caption         =   "Coustomer"
      End
   End
   Begin VB.Menu mnuBussiness 
      Caption         =   "&Bussiness"
      Begin VB.Menu mnuPurchase 
         Caption         =   "&Purchase"
      End
      Begin VB.Menu mnusale 
         Caption         =   "&Sale"
      End
      Begin VB.Menu mnuPay 
         Caption         =   "&Payment"
      End
      Begin VB.Menu mnurecipt 
         Caption         =   "&Recipt"
      End
   End
   Begin VB.Menu mnuEmployeeStatus 
      Caption         =   "&Employee Status"
      Begin VB.Menu mnuempatt 
         Caption         =   "&Employee Attends"
      End
      Begin VB.Menu mnuEmpSal 
         Caption         =   "Employee salary "
      End
   End
   Begin VB.Menu mnuQry 
      Caption         =   "&Query"
      Begin VB.Menu mnuPurQry 
         Caption         =   "Purchase"
      End
      Begin VB.Menu mnuSalQry 
         Caption         =   "Sale"
      End
      Begin VB.Menu mnuPayQry 
         Caption         =   "Payment"
      End
      Begin VB.Menu mnuRecQry 
         Caption         =   "Recipt"
      End
      Begin VB.Menu mnuAstock 
         Caption         =   "All stock"
      End
      Begin VB.Menu mnuStock 
         Caption         =   "Stock"
      End
      Begin VB.Menu mnuEmSlQry 
         Caption         =   "Employee Salary"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuStockRpt 
         Caption         =   "Stock Report"
      End
      Begin VB.Menu mnuPurchaseRpt 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu mnuSaleRpt 
         Caption         =   "Sale Report"
      End
      Begin VB.Menu mnuRecRep 
         Caption         =   "Recipt Report"
      End
      Begin VB.Menu mnuPayRep 
         Caption         =   "Payment report"
      End
      Begin VB.Menu mnuEmpRep 
         Caption         =   "Employee Report"
      End
      Begin VB.Menu EmpSalRep 
         Caption         =   "Employee Salary Report"
      End
      Begin VB.Menu mnuDealRep 
         Caption         =   "Dealer Report"
      End
      Begin VB.Menu mnucoustRep 
         Caption         =   "CoustomerReport"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnulogout 
      Caption         =   "&Logout"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub FishDetails_Click()
EditFishDetails.Show
End Sub

Private Sub mnuAstock_Click()
StockQry.Show
End Sub

Private Sub mnuCoustomer_Click()
Coustomer.Show
End Sub

Private Sub mnuCoutQry_Click()
StockQry.Show
End Sub
Private Sub EmpSalRep_Click()
Dim sl As Integer
Dim sdate As Date
Dim edate As Date
sdate = CDate(InputBox("Type From Date", "Please Enter", Date))
edate = CDate(InputBox("Type Upto Date", "Please Enter", Date))
sl = Val(InputBox("Type for a month Salary ", "Please Enter "))
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
    cname = rs1.Fields(0)
    ctel = rs1.Fields(3)
    picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
End If
rs1.Close

rs.Open "Select EmId,count(Status) as pDays,count(Status)*" & sl & " as Sal from Attends where Status='P' and Date>=#" & sdate & " # and date<=#" & edate & "# group by EmId", con, adOpenDynamic, adLockOptimistic, adCmdText


With ESRPt

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
    
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        
    End With
End With


ESRPt.Show
End Sub

Private Sub mnuDealRep_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
    cname = rs1.Fields(0)
    ctel = rs1.Fields(3)
    picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
End If
rs1.Close
rs.Open "Dealer", con, adOpenDynamic, adLockOptimistic, adCmdTable

With CRpt

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        .Item(4).DataField = rs.Fields(3).Name
       
    End With
End With
CRpt.Show
End Sub

Private Sub mnudealer_Click()
Dealer.Show
End Sub

Private Sub mnucominfo_Click()
Company.Show
End Sub

Private Sub mnucoustRep_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
    cname = rs1.Fields(0)
    ctel = rs1.Fields(3)
    picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
End If
rs1.Close
rs.Open "coustomer", con, adOpenDynamic, adLockOptimistic, adCmdTable

With DRpt

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        .Item(4).DataField = rs.Fields(3).Name
       
    End With
End With
DRpt.Show
End Sub

Private Sub mnuDLQry_Click()
frmLogin.Show
End Sub

Private Sub mnuempatt_Click()
EmpAtte.Show
End Sub

Private Sub mnuEmployee_Click()
EmployeeDetails.Show
End Sub

Private Sub mnuEmpRep_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
    cname = rs1.Fields(0)
    ctel = rs1.Fields(3)
    picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
End If
rs1.Close
rs.Open "Employee", con, adOpenDynamic, adLockOptimistic, adCmdTable

With DRpt

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        .Item(4).DataField = rs.Fields(3).Name
       
    End With
End With
DRpt.Show
End Sub

Private Sub mnuEmpSal_Click()
EmpSalary.Show
End Sub

Private Sub mnuEmSlQry_Click()
ESalQry.Show
End Sub

Private Sub mnuexit_Click()

End
End Sub

Private Sub mnuFish_Click()
EditFishDetails.Show
End Sub

Private Sub mnulogout_Click()
Unload Me
 frmLogin1.Show
 
End Sub

Private Sub mnupay_Click()
Payment.Show
End Sub

Private Sub mnuPayQry_Click()
payQry.Show
End Sub

Private Sub mnuPayRep_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
    If rs1.EOF = False Then
        cname = rs1.Fields(0)
        ctel = rs1.Fields(3)
        picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
    End If
    rs1.Close
    rs.Open "payment", con, adOpenDynamic, adLockOptimistic, adCmdTable

With PyRpt

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        .Item(4).DataField = rs.Fields(3).Name
        .Item(5).DataField = rs.Fields(4).Name
        .Item(6).DataField = rs.Fields(5).Name
    End With
End With
PyRpt.Show

End Sub

Private Sub mnuPurchase_Click()
PurchaseDetails.Show
End Sub

Private Sub mnuPurchaseRpt_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
If rs1.EOF = False Then
    cname = rs1.Fields(0)
    ctel = rs1.Fields(3)
    picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
End If
rs1.Close
rs.Open "Purchase", con, adOpenDynamic, adLockOptimistic, adCmdTable

With PREP

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        .Item(4).DataField = rs.Fields(4).Name
        .Item(5).DataField = rs.Fields(6).Name
    End With
End With
PREP.Show

End Sub

Private Sub mnuPurQry_Click()
PurQry.Show
End Sub

Private Sub mnurecipt_Click()
Recipt.Show
End Sub

Private Sub mnuRecQry_Click()
RecQry.Show
End Sub

Private Sub mnuRecRep_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
    If rs1.EOF = False Then
        cname = rs1.Fields(0)
        ctel = rs1.Fields(3)
        picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
    End If
    rs1.Close
    rs.Open "recipt", con, adOpenDynamic, adLockOptimistic, adCmdTable

With RCRpt

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(1).Name
        .Item(3).DataField = rs.Fields(2).Name
        .Item(4).DataField = rs.Fields(3).Name
        .Item(5).DataField = rs.Fields(4).Name
        .Item(6).DataField = rs.Fields(5).Name
    End With
End With
RCRpt.Show

End Sub


Private Sub mnuSalQry_Click()
saleQry.Show
End Sub

Private Sub mnuStock_Click()
Stock.Show
End Sub



Private Sub mnusale_Click()
SaleDetails.Show
End Sub

Private Sub mnuSaleRpt_Click()
Dim rs As New ADODB.Recordset
rs.Open "sale", con, adOpenDynamic, adLockOptimistic, adCmdTable

With SREP
Set .DataSource = Nothing
.DataMember = ""
Set .DataSource = rs.DataSource
With .Sections("Section1").Controls
.Item(1).DataField = rs.Fields(0).Name
.Item(2).DataField = rs.Fields(1).Name
.Item(3).DataField = rs.Fields(2).Name
.Item(4).DataField = rs.Fields(7).Name
.Item(5).DataField = rs.Fields(8).Name
.Item(6).DataField = rs.Fields(9).Name
End With
End With
SREP.Show
End Sub




Private Sub mnuStockRpt_Click()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cname, ctel, picfile As String
rs1.Open "company", con, adOpenDynamic, adLockOptimistic, adCmdTable
    If rs1.EOF = False Then
        cname = rs1.Fields(0)
        ctel = rs1.Fields(3)
        picfile = IIf(IsNull(rs1.Fields(5)), "", rs1.Fields(5))
    End If
    rs1.Close
    rs.Open "Fish", con, adOpenDynamic, adLockOptimistic, adCmdTable

With StkRep

    Set .DataSource = Nothing
    .DataMember = ""
    Set .DataSource = rs.DataSource
    With .Sections("Section4").Controls
        .Item(2).Caption = cname
        .Item(3).Caption = ctel
       Set .Item(5).Picture = LoadPicture(picfile)
    End With
    With .Sections("Section1").Controls
        .Item(1).DataField = rs.Fields(0).Name
        .Item(2).DataField = rs.Fields(2).Name
        .Item(3).DataField = rs.Fields(3).Name
    End With
End With
StkRep.Show

End Sub

