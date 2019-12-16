VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInitBalance 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmInitBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10410
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   6645
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   11721
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.TextBox txtManual 
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3000
         Width           =   10095
      End
      Begin prjFarmManagement.uctlDate uctlFromDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Top             =   1530
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   465
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   820
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjFarmManagement.uctlTextBox txtPercent 
         Height          =   465
         Left            =   1860
         TabIndex        =   2
         Top             =   1860
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
      End
      Begin prjFarmManagement.uctlTextBox txtProcess 
         Height          =   465
         Left            =   1860
         TabIndex        =   13
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   820
      End
      Begin VB.Label lblProcess 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   660
         Width           =   1755
      End
      Begin Threed.SSCommand cmdStart 
         Height          =   525
         Left            =   1860
         TabIndex        =   3
         Top             =   2340
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInitBalance.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   3600
         TabIndex        =   11
         Top             =   1980
         Width           =   1275
      End
      Begin VB.Label lblProgress 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   10
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   9
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   210
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   8535
         TabIndex        =   5
         Top             =   2340
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6885
         TabIndex        =   4
         Top             =   2340
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmInitBalance.frx":2ABC
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmInitBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Private m_HasActivate As Boolean
Private m_HasModify As Boolean

Public id As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdStart_Click()
On Error GoTo ErrorHandler

Dim IsOK As Boolean
Dim iCount As Long
Dim RecordCount As Long
Dim Percent As Double
Dim I As Long
Dim HasBegin As Boolean
Dim Result As Boolean

Dim ErrorObj As clsErrorLog
Dim BalanceAmount As Collection
Dim DoItemCollection As Collection
Dim CnDnBuyCollection As Collection
   Call glbDaily.StartTransaction
   
   Set ErrorObj = New clsErrorLog
   Set BalanceAmount = New Collection
   Set DoItemCollection = New Collection
   HasBegin = False
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   
   I = 0
   
   Set BalanceAmount = New Collection
   Set DoItemCollection = New Collection
    Set CnDnBuyCollection = New Collection
    
   txtProcess.Text = "สร้างข้อมูลยอดยกมาวัตถุดิบ"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "LOADING"
   txtPercent.Refresh
   
   'Call LoadInventoryBalanceEx(Nothing, BalanceAmount, uctlFromDate.ShowDate)
   'Copy ข้อมูลยอดยกมาของ MonthlyAccum ไปไว้ที่ BalanceAccum
    Dim YYYYMM   As String
    Dim MonthlyAccums  As Collection
    Dim InventoryBals  As Collection
   YYYYMM = Format(Year(DateAdd("D", -1, uctlFromDate.ShowDate)), "0000") & "-" & Format(Month(DateAdd("D", -1, uctlFromDate.ShowDate)), "00")
   Set MonthlyAccums = New Collection
   Set InventoryBals = New Collection
   Call LoadMonthlyBalance(Nothing, MonthlyAccums, YYYYMM)
   Call glbDaily.CopyMonthlyAccum(MonthlyAccums, InventoryBals)
   
   
   
   'Delete All Data
   txtProcess.Text = "ลบข้อมูลเดิมในระบบ"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "PROCESSING"
   txtPercent.Refresh
   
   Call DeleteBalance(DateAdd("D", -1, uctlFromDate.ShowDate))
   Dim Ma As CMonthlyAccum
   Set Ma = New CMonthlyAccum
   Ma.TO_YYYYMM = Format(Year(DateAdd("M", -2, uctlFromDate.ShowDate)), "0000") & "-" & Format(Month(DateAdd("M", -2, uctlFromDate.ShowDate)), "00")
   Call Ma.ClearData
   Set Ma = Nothing
   
   Call InsertBalanceAccum(InventoryBals, DateAdd("D", -1, uctlFromDate.ShowDate))   ' บันทึกจาก MonthlyAccum ไปยัง BalanceAccum
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim Ivd As CInventoryDoc
   txtProcess.Text = "สร้างข้อมูลยอดยกมาสินค้า/วัตถุดิบ"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "PROCESSING"
   txtPercent.Refresh
   
   Set Ivd = New CInventoryDoc
   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
   Ivd.DOCUMENT_NO = "ตั้งยอดใหม่สินค้า/วัตถุดิบ"
   Ivd.DOCUMENT_TYPE = 1
   Ivd.COMMIT_FLAG = "N"
   Ivd.EXCEPTION_FLAG = "N"
   
   Dim II As CLotItem
   I = 0
   iCount = InventoryBals.Count
   For Each II In InventoryBals
   
      I = I + 1
      Percent = MyDiffEx(I, iCount) * 100
      prgProgress.Value = Percent
      txtPercent.Text = FormatNumber(Percent)
      txtPercent.Refresh
         
      '''Debug.Print (II.CURRENT_AMOUNT)
      If II.PART_ITEM_ID > 0 And Len(II.PART_NO) > 0 Then
      
         II.Flag = "A"
         II.TX_TYPE = "I"
         II.CALCULATE_FLAG = "Y"
         Call Ivd.ImportExports.add(II)
      Else
         ''Debug.Print II.PART_ITEM_ID
         If II.PART_ITEM_ID = 1105 Then
            'Debug.Print
         End If
      End If
   Next II
   
   txtProcess.Text = "บันทึกข้อมูลสินค้า/วัตถุดิบ"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "LOADING"
   txtPercent.Refresh
   
   Call glbDaily.AddEditInventoryDoc(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   Dim m_PaidAmounts As Collection
   Dim m_DnAmounts As Collection
   Dim m_CnAmounts  As Collection
   Dim m_RtAmounts As Collection
   Dim m_DiscountAmounts As Collection
   Dim m_CashBefore As Collection
   Dim m_CashAfter As Collection
   Dim m_AccountTxs As Collection
   Dim ct As CCashTran
   
   Set m_PaidAmounts = New Collection
   Set m_DnAmounts = New Collection
   Set m_CnAmounts = New Collection
   Set m_RtAmounts = New Collection
   Set m_DiscountAmounts = New Collection
   Set m_CashBefore = New Collection
   Set m_CashAfter = New Collection
   Set m_AccountTxs = New Collection
   
   txtProcess.Text = "LOAD ข้อมูลยอดยกมาลูกหนี้"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "LOADING"
   txtPercent.Refresh
   
   Call LoadPaidAmountByBill(Nothing, m_PaidAmounts, , DateAdd("D", -1, uctlFromDate.ShowDate))
   Call LoadDnCnAmountByBill(Nothing, m_DnAmounts, , DateAdd("D", -1, uctlFromDate.ShowDate), 4, 2)
   Call LoadDnCnAmountByBill(Nothing, m_CnAmounts, , DateAdd("D", -1, uctlFromDate.ShowDate), 3, 2)
   Call LoadDnCnAmountByBill(Nothing, m_RtAmounts, , DateAdd("D", -1, uctlFromDate.ShowDate), 18, 2)
   Call LoadBillingDiscountByBill(Nothing, m_DiscountAmounts, , DateAdd("D", -1, uctlFromDate.ShowDate))
   
   txtProcess.Text = "LOAD ข้อมูลการเงิน"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "LOADING"
   txtPercent.Refresh
   
   Set ct = New CCashTran
   Call ct.SetFieldValue("FROM_DATE", -1)
   Call ct.SetFieldValue("TO_DATE", DateAdd("D", -1, uctlFromDate.ShowDate))
   Call LoadSumCashTrnAmount(ct, Nothing, m_CashBefore)
   Set ct = Nothing
      
   Set ct = New CCashTran
   Call LoadBankAccountInCashTrn(ct, Nothing, m_AccountTxs)
   Set ct = Nothing
   
   Dim BD As CBillingDoc
   Dim Rs As ADODB.Recordset
   Dim Ri1_0 As CReceiptItem
   Dim Ri1_1 As CReceiptItem
   Dim Ri1_2 As CReceiptItem
   Dim Ri1_3 As CReceiptItem
   Dim Bdc As CBillingDiscount

   Set BD = New CBillingDoc
   Set Rs = New ADODB.Recordset
   
   txtProcess.Text = "LOAD ข้อมูลยอดยกมาลูกหนี้"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "LOADING"
   txtPercent.Refresh
   
   BD.BILLING_DOC_ID = -1
   BD.TO_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
   BD.OrderBy = 3
   BD.DOCUMENT_TYPE = 1
   BD.ItemSumFlag = True
   Call glbDaily.QueryBillingDoc(BD, Rs, iCount, IsOK, glbErrorLog)
   
   txtProcess.Text = "สร้างข้อมูลยอดยกมาลูกหนี้"
   txtProcess.Refresh
   prgProgress.Value = 0
   txtPercent.Text = "PROCESSING"
   txtPercent.Refresh
   
   I = 0
   
   While Not Rs.EOF
         I = I + 1
         Percent = MyDiffEx(I, iCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         txtPercent.Refresh
         Call BD.PopulateFromRS(1, Rs)
         
         Set Ri1_0 = GetReceiptItem(m_PaidAmounts, BD.BILLING_DOC_ID) 'รับชำระ
         Set Ri1_1 = GetReceiptItem(m_DnAmounts, BD.BILLING_DOC_ID) 'เพิ่มหนี้
         Set Ri1_2 = GetReceiptItem(m_CnAmounts, BD.BILLING_DOC_ID) 'ลดหนี้
         Set Ri1_3 = GetReceiptItem(m_RtAmounts, BD.BILLING_DOC_ID) 'รับคืน
         Set Bdc = GetBillingDiscount(m_DiscountAmounts, BD.BILLING_DOC_ID) 'ส่วนลด
         
         BD.PAID_AMOUNT = Ri1_0.PAID_AMOUNT
         BD.CASH_DISCOUNT = Ri1_0.CASH_DISCOUNT
         BD.DEBIT_AMOUNT = Ri1_1.DEBIT_CREDIT_AMOUNT
         BD.CREDIT_AMOUNT = Ri1_2.DEBIT_CREDIT_AMOUNT
         BD.RETURN_AMOUNT = Ri1_3.DEBIT_CREDIT_AMOUNT
         BD.DISCOUNT_AMOUNT = Bdc.DISCOUNT_AMOUNT
                  
         'ทำให้เป็น 2 ตำแหน่งก่อนแล้วค่อยรวม
         BD.PAID_AMOUNT = Val(Format(BD.PAID_AMOUNT, "0.00"))
         BD.CASH_DISCOUNT = Val(Format(BD.CASH_DISCOUNT, "0.00"))
         BD.DEBIT_AMOUNT = Val(Format(BD.DEBIT_AMOUNT, "0.00"))
         BD.CREDIT_AMOUNT = Val(Format(BD.CREDIT_AMOUNT, "0.00"))
         BD.RETURN_AMOUNT = Val(Format(BD.RETURN_AMOUNT, "0.00"))
         BD.DISCOUNT_AMOUNT = Val(Format(BD.DISCOUNT_AMOUNT, "0.00"))
         
         BD.DO_TOTAL_PRICE = Val(Format(BD.DO_TOTAL_PRICE, "0.00"))
         BD.REVENUE_TOTAL_PRICE = Val(Format(BD.REVENUE_TOTAL_PRICE, "0.00"))
         BD.PAID_AMOUNT = Val(Format(BD.PAID_AMOUNT, "0.00"))
         
         '''Debug.Print (Bd.CUSTOMER_CODE)
         If ROUND(BD.DO_TOTAL_PRICE + BD.REVENUE_TOTAL_PRICE - BD.DISCOUNT_AMOUNT + (BD.DEBIT_AMOUNT - BD.CREDIT_AMOUNT - BD.RETURN_AMOUNT) - BD.PAID_AMOUNT - BD.CASH_DISCOUNT, 2) = 0 Then
            
            Ri1_0.DO_ID = BD.BILLING_DOC_ID
            Call Ri1_0.DeleteDataFromDoID
         
            Call BD.DeleteData
            
         End If
         Rs.MoveNext
      Wend
      
      txtProcess.Text = "จัดเรียงข้อมูลยอดยกมาลูกหนี้"
      txtProcess.Refresh
      prgProgress.Value = 0
      txtPercent.Text = "PROCESSING"
      txtPercent.Refresh
   
      Dim TempDate  As String
      Dim SQL1 As String
      TempDate = DateToStringIntHi(Trim(DateAdd("D", -1, uctlFromDate.ShowDate)))
      '-----------------------------ขาย------------------------------------------------------
      SQL1 = "DELETE FROM DO_ITEM UG WHERE UG.DO_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 1 AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
      
      SQL1 = "DELETE FROM SALE_ORDER UG WHERE UG.DO_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 19  AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบ SO
      
      SQL1 = "DELETE FROM GL_DETAIL UG WHERE UG.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 1 AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'GL
      
      SQL1 = "DELETE FROM GL_DETAIL UG WHERE UG.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
      
      SQL1 = "DELETE FROM CASH_TRAN CT WHERE CT.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 1 AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 1 AND DOCUMENT_DATE <= '" & TempDate & "' "
      m_Conn.Execute (SQL1)               'ใบเสร็จขายสด
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 19 AND DOCUMENT_DATE <= '" & TempDate & "' "
      m_Conn.Execute (SQL1)               'ใบ SO
      
      SQL1 = "DELETE FROM CASH_TRAN CT WHERE CT.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 3  AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0)) "
      m_Conn.Execute (SQL1)               'ใบเสร็จขายเชื่อที่ไม่มี Item ลูกแล้ว
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 4  AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0) "
      m_Conn.Execute (SQL1)               'ใบเพิ่มหนี้ที่ไม่มี Item ลูกแล้ว
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 3  AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0) "
      m_Conn.Execute (SQL1)               'ใบลดหนี้ที่ไม่มี Item ลูกแล้ว
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE DOCUMENT_TYPE = 2 AND BD.RECEIPT_TYPE = 3 AND DOCUMENT_DATE <= '" & TempDate & "' AND ((SELECT COUNT(*) FROM RECEIPT_ITEM RCI WHERE RCI.BILLING_DOC_ID = BD.BILLING_DOC_ID)<=0) "
      m_Conn.Execute (SQL1)               'ใบเสร็จขายเชื่อที่ไม่มี Item ลูกแล้ว
      
      SQL1 = "DELETE FROM CASH_TRAN CT WHERE CT.CASH_DOC_ID IN (SELECT BD.CASH_DOC_ID FROM  CASH_DOC BD WHERE DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'CASH_TRAN จาก TABLE CASH_DOC
      
      SQL1 = "DELETE FROM CASH_DOC_POST CT WHERE CT.CASH_DOC_ID IN (SELECT BD.CASH_DOC_ID FROM  CASH_DOC BD WHERE DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'CASH_DOC_POST จาก TABLE CASH_DOC
      
      SQL1 = "DELETE FROM CASH_DOC BD WHERE DOCUMENT_DATE <= '" & TempDate & "'"
      m_Conn.Execute (SQL1)               'TABLE CASH_DOC
      '-------------------------------ขาย-----------------------------------
      '-------------------------------ซื้อ------------------------------------
      SQL1 = "DELETE FROM RECEIPT_ITEM RI WHERE RI.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 18 AND BD.RECEIPT_TYPE = 3  AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบ เสร็จซื้อ
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE BD.DOCUMENT_TYPE = 8 AND BD.RECEIPT_TYPE = 3 AND DOCUMENT_DATE <= '" & TempDate & "' "
      m_Conn.Execute (SQL1)               'ใบ เสร็จซื้อ
      
      SQL1 = "DELETE FROM RECEIPT_ITEM RI WHERE RI.BILLING_DOC_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE (BD.DOCUMENT_TYPE = 10 OR BD.DOCUMENT_TYPE = 9 OR BD.DOCUMENT_TYPE = 110 ) AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบ เพิ่มหนี้ ลดหนี้ รับคืน
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE (BD.DOCUMENT_TYPE = 10 OR BD.DOCUMENT_TYPE = 9 OR BD.DOCUMENT_TYPE = 110 ) AND DOCUMENT_DATE <= '" & TempDate & "' "
      m_Conn.Execute (SQL1)               'ใบ เพิ่มหนี้ ลดหนี้ รับคืน
      
      Call LoadDistinctDoIDFromReceipt(Nothing, CnDnBuyCollection, uctlFromDate.ShowDate)
      
      Set BD = New CBillingDoc
      BD.BILLING_DOC_ID = -1
      BD.TO_DATE = DateAdd("D", -1, uctlFromDate.ShowDate)
      BD.OrderBy = 3
      BD.DOCUMENT_TYPE_SET = "('100','101','102','103')"
      BD.ItemSumFlag = False
      Call glbDaily.QueryBillingDoc(BD, Rs, iCount, IsOK, glbErrorLog)
   
      txtProcess.Text = "กำลังลบข้อมูลเจ้าหนี้"
      txtProcess.Refresh
      prgProgress.Value = 0
      txtPercent.Text = "PROCESSING"
      txtPercent.Refresh
   
      I = 0
      
      Dim TempRpt As CReceiptItem
      While Not Rs.EOF
         I = I + 1
                     
         Percent = MyDiffEx(I, iCount) * 100
         prgProgress.Value = Percent
         txtPercent.Text = FormatNumber(Percent)
         txtPercent.Refresh
         Call BD.PopulateFromRS(1, Rs)
         
         Set TempRpt = GetObject("CReceiptItem", CnDnBuyCollection, Trim(str(BD.BILLING_DOC_ID)), False)
         If TempRpt Is Nothing Then
            SQL1 = "DELETE FROM SUP_ITEM WHERE DO_ID = " & BD.BILLING_DOC_ID
            m_Conn.Execute (SQL1)               'ใบ รับเข้า
                  
            SQL1 = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID = " & BD.BILLING_DOC_ID
            m_Conn.Execute (SQL1)               'ใบ รับเข้า
         End If
         
         Rs.MoveNext
      Wend
      
      SQL1 = "DELETE FROM SUP_ITEM UG WHERE UG.DO_ID IN (SELECT BD.BILLING_DOC_ID FROM  BILLING_DOC BD WHERE (BD.DOCUMENT_TYPE = 1000 OR BD.DOCUMENT_TYPE = 1001 OR BD.DOCUMENT_TYPE = 1002 OR BD.DOCUMENT_TYPE = 1003 ) AND DOCUMENT_DATE <= '" & TempDate & "') "
      m_Conn.Execute (SQL1)               'ใบ PO
      
      SQL1 = "DELETE FROM BILLING_DOC BD WHERE (BD.DOCUMENT_TYPE = 1000 OR BD.DOCUMENT_TYPE = 1001 OR BD.DOCUMENT_TYPE = 1002 OR BD.DOCUMENT_TYPE = 1003 ) AND DOCUMENT_DATE <= '" & TempDate & "' "
      m_Conn.Execute (SQL1)               'ใบ PO
      '-------------------------------ซื้อ------------------------------------
      
      
      SQL1 = "DELETE FROM LOGIN_TRACKING"
      m_Conn.Execute (SQL1)         'LOGIN
   
      txtProcess.Text = "สร้าง ข้อมูลยอดเงินสด"
      txtProcess.Refresh
      prgProgress.Value = 0
      txtPercent.Text = "PROCESSING"
      txtPercent.Refresh
      
      Set ct = New CCashTran
      Call ct.SetFieldValue("FROM_DATE", -1)
      Call ct.SetFieldValue("TO_DATE", DateAdd("D", -1, uctlFromDate.ShowDate))
      Call LoadSumCashTrnAmount(ct, Nothing, m_CashAfter)
      Set ct = Nothing
   
      Dim CashDoc As CCashDoc
      Dim Ct1 As CCashTran
      Dim Ct2 As CCashTran
      Dim Ct3 As CCashTran
      Dim Ct4 As CCashTran
      Dim GetBalanceAmount As Double
      Dim TempCt As CCashTran
      
      For Each ct In m_AccountTxs
         Set Ct1 = GetCashTran(m_CashBefore, ct.GetFieldValue("BANK_ACCOUNT") & "-" & "I")
         Set Ct2 = GetCashTran(m_CashBefore, ct.GetFieldValue("BANK_ACCOUNT") & "-" & "E")
         Set Ct3 = GetCashTran(m_CashAfter, ct.GetFieldValue("BANK_ACCOUNT") & "-" & "I")
         Set Ct4 = GetCashTran(m_CashAfter, ct.GetFieldValue("BANK_ACCOUNT") & "-" & "E")
         
         GetBalanceAmount = Ct1.GetFieldValue("NET_AMOUNT") - Ct2.GetFieldValue("NET_AMOUNT") - Ct3.GetFieldValue("NET_AMOUNT") + Ct4.GetFieldValue("NET_AMOUNT")
         
         If ROUND(GetBalanceAmount, 2) <> 0 Then
            Set CashDoc = New CCashDoc
            CashDoc.ShowMode = SHOW_ADD
            Call CashDoc.SetFieldValue("DOCUMENT_DATE", DateAdd("D", -1, uctlFromDate.ShowDate))
            Call CashDoc.SetFieldValue("DOCUMENT_NO", "ตั้งยอดยอดเงิน-" & DateAdd("D", -1, uctlFromDate.ShowDate))
            Call CashDoc.SetFieldValue("DOCUMENT_TYPE", CASH_DEPOSIT)
            Call CashDoc.SetFieldValue("BANK_ID", ct.GetFieldValue("BANK_ID"))
            Call CashDoc.SetFieldValue("BANK_BRANCH", ct.GetFieldValue("BANK_BRANCH"))
            Call CashDoc.SetFieldValue("BANK_ACCOUNT", ct.GetFieldValue("BANK_ACCOUNT"))
            
            'นำฝากเงินสดในมือ
            
            Set TempCt = New CCashTran
            TempCt.Flag = "A"
            Call TempCt.SetFieldValue("PAYMENT_TYPE", 1) 'ออกเป็นเงินสด
            Call TempCt.SetFieldValue("BANK_ID", ct.GetFieldValue("BANK_ID"))
            Call TempCt.SetFieldValue("BANK_BRANCH", ct.GetFieldValue("BANK_BRANCH"))
            Call TempCt.SetFieldValue("BANK_ACCOUNT", ct.GetFieldValue("BANK_ACCOUNT"))
            If ROUND(GetBalanceAmount, 2) > 0 Then
               Call TempCt.SetFieldValue("AMOUNT", GetBalanceAmount)
               Call TempCt.SetFieldValue("TX_TYPE", "I")
               Call TempCt.SetFieldValue("NET_AMOUNT", GetBalanceAmount)
            ElseIf ROUND(GetBalanceAmount, 2) < 0 Then
               Call TempCt.SetFieldValue("AMOUNT", -GetBalanceAmount)
               Call TempCt.SetFieldValue("TX_TYPE", "E")
               Call TempCt.SetFieldValue("NET_AMOUNT", -GetBalanceAmount)
            End If
            
            Call CashDoc.CashTranItems.add(TempCt)
            
            Call glbDaily.AddEditCashDoc(CashDoc, IsOK, False, glbErrorLog)
         End If
      Next ct
   
   Call glbDaily.CommitTransaction
   
   glbErrorLog.LocalErrorMsg = "การปรับยอดประจำปีเสร็จสิ้น"
   glbErrorLog.ShowUserError
   
   OKClick = True
   Unload Me
   Set BD = Nothing
   Exit Sub
   
ErrorHandler:
   Call glbDaily.RollbackTransaction
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   Call EnableForm(Me, True)
End Sub
Private Sub InsertBalanceAccum(MaColls As Collection, DateInsert As Date)
'On Error Resume Next
Dim Ba As CBalanceAccum
Dim II As CLotItem
Dim iCount As Long
Dim TempMa As CMonthlyAccum
   
   For Each II In MaColls
      Set Ba = New CBalanceAccum
      
      Ba.PART_ITEM_ID = II.PART_ITEM_ID
      Ba.FROM_DATE = DateInsert
      Ba.TO_DATE = DateInsert
      Ba.LOCATION_ID = II.LOCATION_ID
      
      Ba.AddEditMode = SHOW_ADD
      
      Ba.DOCUMENT_DATE = DateInsert
      Ba.IMPORT_AMOUNT = II.ALL_IMPORT_AMT
      Ba.EXPORT_AMOUNT = II.ALL_EXPORT_AMT
      Ba.BALANCE_AMOUNT = II.BALANCE_AMOUNT
      Ba.TOTAL_INCLUDE_PRICE = II.TOTAL_INCLUDE_PRICE
      Ba.AVG_PRICE = II.INCLUDE_UNIT_PRICE
      Call Ba.AddEditData
      
      Set Ba = Nothing
   Next II
      
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.NAME
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub
Private Sub ResetStatus()
   prgProgress.MAX = 100
   prgProgress.MIN = 0
   prgProgress.Value = 0
   txtPercent.Text = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = MapText("ประมวลผลประจำปี")
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.NAME = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   Call InitNormalLabel(lblFileName, "วันที่")
   Call InitNormalLabel(lblProgress, "ความคืบหน้า")
   Call InitNormalLabel(lblPercent, "เปอร์เซนต์")
   Call InitNormalLabel(Label1, "%")
   Call InitNormalLabel(lblProcess, "โปรเซส")
   
   Call txtProcess.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   txtProcess.Enabled = False
   Call txtPercent.SetTextLenType(TEXT_FLOAT, glbSetting.MONEY_TYPE)
   txtPercent.Enabled = False
   
   txtManual.FontSize = 12
   txtManual.Text = "ระบบประมวลผลประจำปี " & vbCrLf & _
                                  "มีวิธีการใช้คือ" & vbCrLf & _
                                  "1.ใช้เพื่อลดขนาดข้อมูลให้เล็กลง โดยกรอกวันที่เป็นวันต้นเดือนของปีที่ต้องงการจะทำ เช่น 1 มกราคม 255X " & vbCrLf & _
                                  "2.ก่อนการทำจะต้องทำการ COPY ข้อมูลเดิมพร้อมเปลี่ยนชื่อเป็นชื่อ ปีของฐานข้อมูล เช่น WINPRICING400_2551.GDB " & vbCrLf & _
                                  "3.หลังจากการประมวลผลแล้วขนาดไฟล์จะยังไม่เล็กลง ซึ่งจะต้องทำการ BACKUP ให้เป็น .GBK ก่อนแล้วค่อย RESTORE กลับเป็น .GDB ทับไฟล์เดิม "
                                  
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdStart.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdStart, MapText("เริ่ม"))
   
   Call ResetStatus
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub DeleteBalance(ToDate As Date)
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
Dim WhereStr2 As String
   
   WhereStr = ""
   WhereStr2 = ""
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr = "" Then
         WhereStr = " WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM BALANCE_ACCUM " & WhereStr
   m_Conn.Execute (SQL1)
   
   WhereStr = ""
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr = "" Then
         WhereStr = " WHERE (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If ToDate > -1 Then
      TempDate = DateToStringIntHi(Trim(ToDate))
      If WhereStr2 = "" Then
         WhereStr2 = " WHERE (J.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr2 = WhereStr2 & " AND (J.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM JOB_INOUT II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB_VERIFY II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB_RESOURCE II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB_PARAMETER II WHERE II.JOB_ID IN "
   SQL1 = SQL1 & "(SELECT J.JOB_ID FROM JOB J " & WhereStr2 & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM JOB J " & WhereStr2
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM LOT_ITEM II WHERE II.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
      
   SQL1 = "UPDATE BILLING_DOC BD SET BD.COMMIT_FLAG = 'Y',BD.INVENTORY_DOC_ID = NULL WHERE BD.INVENTORY_DOC_ID IN "
   SQL1 = SQL1 & "(SELECT IVD.INVENTORY_DOC_ID FROM INVENTORY_DOC IVD " & WhereStr & ")"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM INVENTORY_DOC IVD " & WhereStr
   m_Conn.Execute (SQL1)
   
End Sub
