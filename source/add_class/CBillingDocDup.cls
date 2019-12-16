VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillingDocDup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
Public ExtraFlag As Long
' End of generic data member

Public CheckTotalAmount  As Boolean

Public BILLING_DOC_ID As Long
Public ACCOUNT_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_TYPE As Long
Public DUE_DATE As Date
Public NOTE As String
Public ENTERPRISE_ID As Long
Public BILLING_ADDRESS_ID As Long
Public ENTERPRISE_ADDRESS_ID As Long
Public TOTAL_AMOUNT As Double
Public VAT_PERCENT As Double
Public VAT_AMOUNT As Double
Public WH_PERCENT As Double
Public WH_AMOUNT As Double
Public DISCOUNT_AMOUNT As Double
Public PAYMENT_TERM As Long
Public ACCEPT_BY As Long
Public RECEIVE_BY As Long
Public EXCEPTION_FLAG As String
Public PAYEE_NAME As String
Public INVENTORY_DOC_ID As Long
Public COMMIT_FLAG As String
Public SUPPLIER_ID As Long
Public RECEIPT_TYPE As Long
Public DEPOSIT_AMOUNT As Double
Public TOTAL_PRICE As Double
Public ESTIMATE_DATE As Date
Public APPROVE_DATE As Date
Public RESOURCE_ID As Long
Public PAYMENT_TYPE As Long
Public BANK_ID As Long
Public BBRANCH_ID As Long
Public BANK_NAME As String
Public BBRANCH_NAME As String
Public CHECK_NO As String
Public CHECK_DATE As Date
Public AREA_CODE As String
Public CD_PERCENT As Double
Public CD_AMOUNT As Double
Public PACKAGE_TYPE As Long
Public TEMP_DO_NO As String
Public PAID_AMOUNT As Double
Public DEBIT_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public DO_TOTAL_PRICE As Double
Public REVENUE_TOTAL_PRICE As Double
Public PAYMENT_ID As Long
Public BANK_BRANCH_ID As Long
Public BANK_NOTE As String
Public BATCH_ID As Long
Public SIMULATE_FLAG As String
Public ENTRY_DATE As Date
Public EXIT_DATE As Date
Public EFFECTIVE_FLAG As String
Public PAY_FLAG As String

Public RETURN_AMOUNT As Double

Public SUM_DISCOUNT_AMOUNT As Double
Public SUM_PAID_AMOUNT As Double
Public TOTAL_DEBIT_AMOUNT As Double
Public TOTAL_CREDIT_AMOUNT As Double
Public SUM_CASH_DISCOUNT As Double
Public OLD_COMMIT_FLAG As String
Public ACCOUNT_NO As String
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public CUSTOMER_ID As Long
Public DEBIT_CREDIT_AMOUNT As Double
Public RECEIPT_ITEM_ID As Long
Public RECEIPT_REF_COUNT As Long
Public DBN_REF_COUNT As Long
Public CDN_REF_COUNT As Long
Public BLS_REF_COUNT As Long
Public RECEIPT_REF_COUNT2 As Long
Public DBN_REF_COUNT2 As Long
Public CDN_REF_COUNT2 As Long
Public DOCUMENT_TYPE_SET As String
Public TAX_FLAG As String
Public WH_FLAG As String
Public EMP_CODE As String
Public EMP_NAME As String
Public EMP_LNAME As String
Public RESOURCE_NO As String
Public PART_NO As String
Public FEATURE_CODE As String
Public VADILITY As Date
Public DELIVERY As Date
Public PAYMENT As Date
Public PAYMENT_DESC As String
Public SHIPMENT As Date
Public TEMP_PAID_AMOUNT As Double
Public TEMP_CHECK_DATE As String
Public TEMP_CHECK_NO As String
Public PAID_TYPE As Long
Public RECEIPT_PAID_AMOUNT As Double
Public REASON_ID As Long
Public CASH_DISCOUNT As Double
Public DO_RECEIPT_FLAG As String
Public CUSTOMER_TYPE As Long
Public CUSTOMER_GRADE As Long
Public SUPPLIER_TYPE As Long
Public SUPPLIER_GRADE As Long

Public REFER_INV As String
Public PACKING_OF As String
Public SHIP_NAME As String
Public SHIPPING_MARKS As String
Public SHIPON_DATE As Date
Public REF As String
Public Credit As Long
Public CREDIT_LIMIT As Double
Public CUSTOMER_CREDIT As Long
Public NEW_ACCOUNT_ID As Long

Public TOTAL_RCP As Double
Public AGREEMENT_DATA As String
Public AGREEMENT_FINANCE As String

Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As CONFIG_DOC_TYPE

Public OLD_CREDIT_AMOUNT As Double

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

'------------------------------------------------------------------------------------------------------------------------------------------------>
Public DO_NO As String
Public TRUCK_NO As String
Public DELIVERY_FEE As Double
Public SENDER_NAME As String
Public RECEIVE_NAME As String
Public DEPARTMENT_ID As Long
Public QUE_NO As String
Public PR_NO As String
Public PV_NO As String
'------------------------------------------------------------------------------------------------------------------------------------------------>

Public DoItems As Collection
Public SupItems As Collection
Public SaleOrders As Collection
Public RoItems As Collection
Public InventoryDocs As Collection
Public ReceiptItems As Collection
Public Pkglsts As Collection
Public MemoBank As Collection
Public BillingDiscounts As Collection
Public BulkHoles As Collection
Public Payments As Collection
Public GlDetails As Collection

Public ItemSumFlag As Boolean
Public ItemSumFlagSup As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set DoItems = New Collection
   Set SaleOrders = New Collection
   Set RoItems = New Collection
   Set InventoryDocs = New Collection
   FROM_DATE = -1
   TO_DATE = -1
   FROM_DUE_DATE = -1
   TO_DUE_DATE = -1
   ENTRY_DATE = -1
   EXIT_DATE = -1
   EFFECTIVE_FLAG = ""
   PAY_FLAG = ""
   
   Set ReceiptItems = New Collection
   Set Pkglsts = New Collection
   Set MemoBank = New Collection
   Set BillingDiscounts = New Collection
   Set BulkHoles = New Collection
   Set Payments = New Collection
   Set SupItems = New Collection
   Set GlDetails = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String

   If Ind = 1 Then
      SelectStr1 = "SELECT RS.RESOURCE_NO, ED.EMP_CODE, ENM.LONG_NAME EMP_NAME, ENM.LAST_NAME EMP_LNAME, BD.*, SNM.LONG_NAME SUPPLIER_NAME, SP.SUPPLIER_CODE, CM.CUSTOMER_CODE, CM.CUSTOMER_ID, NM.LONG_NAME CUSTOMER_NAME, AC.ACCOUNT_NO, NULL RECEIPT_ITEM_ID,BA.BANK_NAME,BB.BBRANCH_NAME,CM.CREDIT_LIMIT,CM.CREDIT CUSTOMER_CREDIT "
      If ItemSumFlag Then
         SelectStr1 = SelectStr1 & ", 0 REVENUE_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_PRICE) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_WEIGHT) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_WEIGHT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.ITEM_AMOUNT) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT COUNT(*) FROM RECEIPT_ITEM RI1, BILLING_DOC BD1 WHERE (RI1.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 2) AND (RI1.DO_ID = BD.BILLING_DOC_ID)) RECEIPT_REF_COUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(PAID_AMOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.DO_ID = BD.BILLING_DOC_ID) RECEIPT_PAID_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(PAID_AMOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.BILLING_DOC_ID = BD.BILLING_DOC_ID) SUM_PAID_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(CASH_DISCOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.BILLING_DOC_ID = BD.BILLING_DOC_ID) SUM_CASH_DISCOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(RI2.CASH_DISCOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.DO_ID = BD.BILLING_DOC_ID) CASH_DISCOUNT_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 3) AND (RI3.DO_ID = BD.BILLING_DOC_ID)) SUM_CREDIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 4) AND (RI3.DO_ID = BD.BILLING_DOC_ID)) SUM_DEBIT_AMOUNT  "
         
         SelectStr1 = SelectStr1 & ", (SELECT SUM(BDC.DISCOUNT_AMOUNT) FROM BILLING_DISCOUNT BDC WHERE BDC.BILLING_DOC_ID = BD.BILLING_DOC_ID) SUM_DISCOUNT_AMOUNT "
         
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 3) AND (RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID)) TOTAL_CREDIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 4) AND (RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID)) TOTAL_DEBIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 18) AND (RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID)) RETURN_AMOUNT  "
      ElseIf ItemSumFlagSup Then
         SelectStr1 = SelectStr1 & ", 0 REVENUE_TOTAL_PRICE "
         'SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_INCLUDE_PRICE) FROM SUP_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_INCLUDE_PRICE) FROM SUP_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) ) DO_TOTAL_PRICE "
         'SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_WEIGHT) FROM SUP_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_WEIGHT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_WEIGHT) FROM SUP_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID)) DO_TOTAL_WEIGHT "
         'SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TX_AMOUNT) FROM SUP_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TX_AMOUNT) FROM SUP_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID)) DO_TOTAL_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT COUNT(*) FROM RECEIPT_ITEM RI1, BILLING_DOC BD1 WHERE (RI1.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 8) AND (RI1.DO_ID = BD.BILLING_DOC_ID)) RECEIPT_REF_COUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(PAID_AMOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.DO_ID = BD.BILLING_DOC_ID) RECEIPT_PAID_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(PAID_AMOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.BILLING_DOC_ID = BD.BILLING_DOC_ID) SUM_PAID_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(CASH_DISCOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.BILLING_DOC_ID = BD.BILLING_DOC_ID) SUM_CASH_DISCOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(RI2.CASH_DISCOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.DO_ID = BD.BILLING_DOC_ID) CASH_DISCOUNT_AMOUNT "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 9) AND (RI3.DO_ID = BD.BILLING_DOC_ID)) SUM_CREDIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 10) AND (RI3.DO_ID = BD.BILLING_DOC_ID)) SUM_DEBIT_AMOUNT  "
         
         SelectStr1 = SelectStr1 & ", (SELECT SUM(BDC.DISCOUNT_AMOUNT) FROM BILLING_DISCOUNT BDC WHERE BDC.BILLING_DOC_ID = BD.BILLING_DOC_ID) SUM_DISCOUNT_AMOUNT "
         
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 9) AND (RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID)) TOTAL_CREDIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 10) AND (RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID)) TOTAL_DEBIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3, BILLING_DOC BD2 WHERE (RI3.BILLING_DOC_ID = BD2.BILLING_DOC_ID) AND (BD2.DOCUMENT_TYPE = 110) AND (RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID)) RETURN_AMOUNT  "
      Else
         SelectStr1 = SelectStr1 & ", 0 REVENUE_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_WEIGHT "
         SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_AMOUNT "
         SelectStr1 = SelectStr1 & ", 0 RECEIPT_REF_COUNT  "
         SelectStr1 = SelectStr1 & ", 0 RECEIPT_PAID_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 SUM_CASH_DISCOUNT  "
         SelectStr1 = SelectStr1 & ", 0 SUM_PAID_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 CASH_DISCOUNT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 SUM_CREDIT_AMOUNT "
         SelectStr1 = SelectStr1 & ", 0 SUM_DEBIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 SUM_DISCOUNT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 TOTAL_CREDIT_AMOUNT "
         SelectStr1 = SelectStr1 & ", 0 TOTAL_DEBIT_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 RETURN_AMOUNT  "
         
      End If
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC_DUP BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (BD.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SP.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME SNM ON (SN.NAME_ID = SNM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (BD.ACCEPT_BY = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RESOURCE RS ON (BD.RESOURCE_ID = RS.RESOURCE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK BA ON (BD.BANK_ID = BA.BANK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_BRANCH BB ON (BD.BBRANCH_ID = BB.BBRANCH_ID) "
      WhereStr = ""
      If CheckTotalAmount Then               'frmAddEditPoItem
         WhereStr = "WHERE ((BD.TOTAL_AMOUNT > (SELECT SUM(SUP.TX_AMOUNT) FROM SUP_ITEM SUP WHERE SUP.PO_ID = BD.BILLING_DOC_ID)) "
         WhereStr = WhereStr & " OR ((SELECT COUNT(*) FROM SUP_ITEM SUP WHERE SUP.PO_ID = BD.BILLING_DOC_ID) = 0 )) "
      End If
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT (BD.ACCOUNT_ID) "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      
      WhereStr = "WHERE (BD.ACCOUNT_ID IS NOT NULL) AND (BD.ACCOUNT_ID > 0) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT BD.BILLING_DOC_ID, CM.CUSTOMER_CODE, CM.CUSTOMER_ID, CM.CREDIT_LIMIT, NM.LONG_NAME CUSTOMER_NAME , ENM.LONG_NAME SALE_NAME, ENM.LAST_NAME SALE_LNAME "
      SelectStr1 = SelectStr1 & ",BD.TOTAL_PRICE,BD.DISCOUNT_AMOUNT,BD.CD_AMOUNT,BD.VAT_AMOUNT,BD.TOTAL_RCP "
      SelectStr1 = SelectStr1 & ",BD.DOCUMENT_DATE,BD.DOCUMENT_NO, BD.DUE_DATE, ED.EMP_CODE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (BD.ACCEPT_BY = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT EMP.EMP_ID, EMP.EMP_CODE, NM.LONG_NAME EMP_NAME, NM.LAST_NAME EMP_LNAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EMP ON (EMP.EMP_ID = BD.ACCEPT_BY) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EMP.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = EN.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT CM.CUSTOMER_CODE, CM.CUSTOMER_ID, NM.LONG_NAME CUSTOMER_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT  BD.BILLING_DOC_ID,BD.DOCUMENT_NO, BD.DOCUMENT_DATE  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      
      WhereStr = "WHERE ((SELECT COUNT(*) FROM DO_ITEM DOC  LEFT OUTER JOIN BILLING_DOC BD1 ON (DOC.DO_ID = BD1.BILLING_DOC_ID) "
      WhereStr = WhereStr & "WHERE BD1.DOCUMENT_TYPE = 1 AND DOC.PO_ID = BD.BILLING_DOC_ID) < 1) "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT DISTINCT BD.BILLING_DOC_ID,BD.DOCUMENT_NO,BD.INVENTORY_DOC_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      WhereStr = ""
   ElseIf Ind = 100 Then
      SelectStr1 = "SELECT BD.BILLING_DOC_ID, SP.SUPPLIER_CODE, SP.SUPPLIER_ID, SP.CREDIT, NM.LONG_NAME SUPPLIER_NAME  "
      SelectStr1 = SelectStr1 & ",BD.TOTAL_PRICE,BD.DISCOUNT_AMOUNT,BD.CD_AMOUNT,BD.VAT_AMOUNT,BD.TOTAL_RCP "
      SelectStr1 = SelectStr1 & ",BD.DOCUMENT_DATE, BD.DOCUMENT_NO, BD.DUE_DATE, BD.CREDIT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = BD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 101 Then
      SelectStr1 = "SELECT DISTINCT BD.SUPPLIER_ID, SP.SUPPLIER_CODE, NM.LONG_NAME SUPPLIER_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = BD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = "WHERE (BD.SUPPLIER_ID IS NOT NULL) "
   End If
      
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
      
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BD.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
   If Len(EFFECTIVE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD.EFFECTIVE_FLAG = '" & EXCEPTION_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (BD.EFFECTIVE_FLAG = '" & EXCEPTION_FLAG & "') "
      End If
   End If
   
   If Len(EXCEPTION_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(PAY_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAY_FLAG = '" & PAY_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PAY_FLAG = '" & PAY_FLAG & "') "
      End If
   End If
   
   If Val(ACCOUNT_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If

   If Len(Trim(CUSTOMER_CODE)) > 0 Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(FROM_CUSTOMER_CODE)) > 0 Then
      If WildCard(FROM_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(TO_CUSTOMER_CODE)) > 0 Then
      If WildCard(TO_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If

   If Len(Trim(CUSTOMER_NAME)) > 0 Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If
      
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
    If FROM_DUE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DUE_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
   If RECEIPT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RECEIPT_TYPE = " & RECEIPT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (RECEIPT_TYPE = " & RECEIPT_TYPE & ") "
      End If
   End If

   If Len(TAX_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (VAT_PERCENT > 0) "
      Else
         WhereStr = WhereStr & "AND (VAT_PERCENT > 0) "
      End If
   End If
   
   If Len(WH_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (WH_PERCENT > 0) "
      Else
         WhereStr = WhereStr & "AND (WH_PERCENT > 0) "
      End If
   End If
   
   If Len(DOCUMENT_TYPE_SET) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      End If
   End If

   If DO_RECEIPT_FLAG = "Y" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((DOCUMENT_TYPE = 1) OR ((DOCUMENT_TYPE = 2) AND (RECEIPT_TYPE = 1))) "
      Else
         WhereStr = WhereStr & "AND ((DOCUMENT_TYPE = 1) OR ((DOCUMENT_TYPE = 2) AND (RECEIPT_TYPE = 1))) "
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If Len(FEATURE_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(FEATURE_CODE, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN FEATURE FT ON (DI.FEATURE_ID = FT.FEATURE_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (SUBSTR(FT.FEATURE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN FEATURE FT ON (DI.FEATURE_ID = FT.FEATURE_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (FT.FEATURE_CODE  = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         End If
      Else
         If WildCard(FEATURE_CODE, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN FEATURE FT ON (DI.FEATURE_ID = FT.FEATURE_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (SUBSTR(FT.FEATURE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM DO_ITEM DI LEFT OUTER JOIN FEATURE FT ON (DI.FEATURE_ID = FT.FEATURE_ID) WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (FT.FEATURE_CODE  = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)"
         End If
      End If
   End If
   
   
   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If CUSTOMER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      End If
   End If

   If Len(Trim(SUPPLIER_CODE)) > 0 Then
      If WildCard(SUPPLIER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         End If
      End If
   End If

   If SUPPLIER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SP.SUPPLIER_TYPE = " & SUPPLIER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (SP.SUPPLIER_TYPE = " & SUPPLIER_TYPE & ") "
      End If
   End If

   If SUPPLIER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SP.SUPPLIER_GRADE = " & SUPPLIER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (SP.SUPPLIER_GRADE = " & SUPPLIER_GRADE & ") "
      End If
   End If

   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY CUSTOMER_CODE,SUPPLIER_CODE "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY NM.LONG_NAME "
      ElseIf OrderBy = 6 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY DOCUMENT_DATE ASC, DOCUMENT_TYPE ASC, DOCUMENT_NO "
         Else
            OrderByStr = "ORDER BY DOCUMENT_DATE ASC, DOCUMENT_TYPE ASC, DOCUMENT_NO "
         End If
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE ASC, DOCUMENT_NO ASC, DUE_DATE "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, DOCUMENT_NO ASC, DUE_DATE "
      Else
         OrderByStr = "ORDER BY BD.DOCUMENT_DATE DESC, DOCUMENT_NO "
      End If
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, DOCUMENT_DATE ASC, BILLING_DOC_ID "
      Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, DOCUMENT_DATE DESC, BILLING_DOC_ID "
      End If
   ElseIf Ind = 3 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, BD.DOCUMENT_DATE  "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, BD.DOCUMENT_DATE  ASC, BD.DOCUMENT_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY ENM.LONG_NAME ASC, ENM.LAST_NAME ASC, BD.DOCUMENT_DATE  ASC, BD.DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, ENM.LONG_NAME ASC, ENM.LAST_NAME ASC, BD.DOCUMENT_DATE  ASC, BD.DOCUMENT_NO "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY ENM.LONG_NAME ASC, ENM.LAST_NAME ASC, CM.CUSTOMER_CODE ASC, BD.DOCUMENT_DATE  ASC, BD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, BD.DOCUMENT_NO  "
      End If
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY EMP.EMP_CODE "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY CM.CUSTOMER_CODE "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY BD.DOCUMENT_DATE "
   ElseIf Ind = 100 Then
      OrderByStr = "ORDER BY SP.SUPPLIER_CODE ASC, BD.DOCUMENT_DATE  "
   ElseIf Ind = 101 Then
      OrderByStr = "ORDER BY SP.SUPPLIER_CODE ASC, NM.LONG_NAME "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   If Ind = 7 Then
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
'''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      BILLING_ADDRESS_ID = NVLI(Rs("BILLING_ADDRESS_ID"), -1)
      ENTERPRISE_ADDRESS_ID = NVLI(Rs("ENTERPRISE_ADDRESS_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
      WH_PERCENT = NVLD(Rs("WH_PERCENT"), 0)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("DO_TOTAL_AMOUNT"), 0)
      ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      RECEIVE_BY = NVLI(Rs("RECEIVE_BY"), -1)
      PAYMENT_TERM = NVLI(Rs("PAYMENT_TERM"), -1)
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "N")
      OLD_COMMIT_FLAG = COMMIT_FLAG
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), -1)
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT_AMOUNT"), 0)
      ESTIMATE_DATE = InternalDateToDate(NVLS(Rs("ESTIMATE_DATE"), ""))
      APPROVE_DATE = InternalDateToDate(NVLS(Rs("APPROVE_DATE"), ""))
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      RESOURCE_ID = NVLI(Rs("RESOURCE_ID"), -1)
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), -1)
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BBRANCH_ID = NVLI(Rs("BBRANCH_ID"), -1)
      AREA_CODE = NVLS(Rs("AREA_CODE"), "")
      CD_AMOUNT = NVLD(Rs("CD_AMOUNT"), 0)
      CD_PERCENT = NVLD(Rs("CD_PERCENT"), 0)
      PACKAGE_TYPE = NVLI(Rs("PACKAGE_TYPE"), -1)
      TEMP_DO_NO = NVLS(Rs("TEMP_DO_NO"), "")
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      RECEIPT_PAID_AMOUNT = NVLD(Rs("RECEIPT_PAID_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("SUM_CREDIT_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("SUM_DEBIT_AMOUNT"), 0)
      DO_TOTAL_PRICE = NVLD(Rs("DO_TOTAL_PRICE"), 0)
      REVENUE_TOTAL_PRICE = NVLD(Rs("REVENUE_TOTAL_PRICE"), 0)
      PAYMENT_ID = NVLI(Rs("PAYMENT_ID"), -1)
      BANK_BRANCH_ID = NVLI(Rs("BANK_BRANCH_ID"), -1)
      BANK_NOTE = NVLS(Rs("BANK_NOTE"), "")
'      BATCH_ID = NVLI(Rs("BATCH_ID"), -1)
'      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
      EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
      EFFECTIVE_FLAG = NVLS(Rs("EFFECTIVE_FLAG"), "")
      
      TOTAL_DEBIT_AMOUNT = NVLD(Rs("TOTAL_DEBIT_AMOUNT"), 0)
      TOTAL_CREDIT_AMOUNT = NVLD(Rs("TOTAL_CREDIT_AMOUNT"), 0)
      RETURN_AMOUNT = NVLD(Rs("RETURN_AMOUNT"), 0)
      SUM_DISCOUNT_AMOUNT = NVLD(Rs("SUM_DISCOUNT_AMOUNT"), 0)
      SUM_PAID_AMOUNT = NVLD(Rs("SUM_PAID_AMOUNT"), 0)
      SUM_CASH_DISCOUNT = NVLD(Rs("SUM_CASH_DISCOUNT"), 0)
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      BBRANCH_NAME = NVLS(Rs("BBRANCH_NAME"), "")
      
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      CHECK_NO = NVLS(Rs("CHECK_NO"), "")
      CHECK_DATE = InternalDateToDate(NVLS(Rs("CHECK_DATE"), ""))
   
      ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      RECEIPT_REF_COUNT = NVLI(Rs("RECEIPT_REF_COUNT"), -1)
   '   DBN_REF_COUNT = NVLI(Rs("DBN_REF_COUNT"), -1)
   '   CDN_REF_COUNT = NVLI(Rs("CDN_REF_COUNT"), -1)
   '   BLS_REF_COUNT = NVLI(Rs("BLS_REF_COUNT"), -1)
   '   RECEIPT_REF_COUNT2 = NVLI(Rs("RECEIPT_REF_COUNT2"), -1)
   '   DBN_REF_COUNT2 = NVLI(Rs("DBN_REF_COUNT2"), -1)
   '   CDN_REF_COUNT2 = NVLI(Rs("CDN_REF_COUNT2"), -1)
      RESOURCE_NO = NVLS(Rs("RESOURCE_NO"), "")
      
      RECEIPT_ITEM_ID = NVLI(Rs("RECEIPT_ITEM_ID"), -1)
      
         VADILITY = InternalDateToDate(NVLS(Rs("VADILITY"), ""))
         DELIVERY = InternalDateToDate(NVLS(Rs("DELIVERY"), ""))
         PAYMENT = InternalDateToDate(NVLS(Rs("PAYMENT"), ""))
         PAYMENT_DESC = NVLS(Rs("PAYMENT_DESC"), "")
         SHIPMENT = InternalDateToDate(NVLS(Rs("SHIPMENT"), ""))
         
         REFER_INV = NVLS(Rs("REFER_INV"), "")
         PACKING_OF = NVLS(Rs("PACKING_OF"), "")
         SHIP_NAME = NVLS(Rs("SHIP_NAME"), "")
         SHIPPING_MARKS = NVLS(Rs("SHIPPING_MARKS"), "")
        SHIPON_DATE = InternalDateToDate(NVLS(Rs("SHIPON_DATE"), ""))
         REF = NVLS(Rs("REF"), "")
         Credit = NVLI(Rs("CREDIT"), 0)
         
         AGREEMENT_DATA = NVLS(Rs("AGREEMENT_DATA"), "")
         AGREEMENT_FINANCE = NVLS(Rs("AGREEMENT_FINANCE"), "")
         
         TOTAL_RCP = NVLD(Rs("TOTAL_RCP"), 0)
         CREDIT_LIMIT = NVLD(Rs("CREDIT_LIMIT"), 0)
         CUSTOMER_CREDIT = NVLI(Rs("CUSTOMER_CREDIT"), 0)
         
         OLD_CREDIT_AMOUNT = NVLD(Rs("OLD_CREDIT_AMOUNT"), 0)
         
         
         DO_NO = NVLS(Rs("DO_NO"), "")
         TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
         DELIVERY_FEE = NVLD(Rs("DELIVERY_FEE"), 0)
         SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
         RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
         DEPARTMENT_ID = NVLI(Rs("DEPARTMENT_ID"), 0)
         QUE_NO = NVLS(Rs("QUE_NO"), "")
         PR_NO = NVLS(Rs("PR_NO"), "")
         PV_NO = NVLS(Rs("PV_NO"), "")
         
      ElseIf Ind = 2 Then
         ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
      ElseIf Ind = 3 Then
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
         CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
         CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      
         TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
         DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
         CD_AMOUNT = NVLD(Rs("CD_AMOUNT"), 0)
         VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
         TOTAL_RCP = NVLD(Rs("TOTAL_RCP"), 0)
         CREDIT_LIMIT = NVLD(Rs("CREDIT_LIMIT"), 0)
      
         EMP_CODE = NVLS(Rs("EMP_CODE"), "")
         EMP_NAME = NVLS(Rs("SALE_NAME"), "")
         EMP_LNAME = NVLS(Rs("SALE_LNAME"), "")
      ElseIf Ind = 4 Then
         ACCEPT_BY = NVLI(Rs("EMP_ID"), -1)
         EMP_CODE = NVLS(Rs("EMP_CODE"), "")
         EMP_NAME = NVLS(Rs("EMP_NAME"), "")
         EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      ElseIf Ind = 5 Then
         CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
         CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
         CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      ElseIf Ind = 6 Then
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      ElseIf Ind = 7 Then
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      ElseIf Ind = 100 Then
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
         SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
         SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      
         TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
         DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
         CD_AMOUNT = NVLD(Rs("CD_AMOUNT"), 0)
         VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
         TOTAL_RCP = NVLD(Rs("TOTAL_RCP"), 0)
         Credit = NVLD(Rs("CREDIT"), 0)
      ElseIf Ind = 101 Then
         SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
         SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
         SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      End If
End Function

Public Function PatchDocDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID0 As String
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String
Dim TempID8 As String
Dim TempID9 As String
Dim TempID10 As String
Dim TempID11 As String
Dim TempID12 As String
Dim TempID13 As String
Dim TempID14 As String
Dim TempID15 As String
Dim TempID16 As String
Dim TempID17 As String
Dim TempID18 As String
Dim TempID19 As String
Dim TempID20 As String
Dim TempID21 As String

   TempID0 = ACCOUNT_ID
   If ACCOUNT_ID <= 0 Then
      TempID0 = "NULL"
   End If
   
   TempID1 = BILLING_ADDRESS_ID
   If BILLING_ADDRESS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = ENTERPRISE_ADDRESS_ID
   If ENTERPRISE_ADDRESS_ID <= 0 Then
      TempID2 = "NULL"
   End If
      
   TempID3 = DOCUMENT_TYPE
   If DOCUMENT_TYPE <= 0 Then
      TempID3 = "NULL"
   End If
      
   TempID4 = VAT_PERCENT
   If VAT_PERCENT <= 0 Then
      TempID4 = "NULL"
   End If
      
   TempID5 = ENTERPRISE_ID
   If ENTERPRISE_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = WH_PERCENT
   If WH_PERCENT <= 0 Then
      TempID6 = "NULL"
   End If
   
   TempID8 = WH_AMOUNT
   If WH_AMOUNT <= 0 Then
      TempID8 = "NULL"
   End If

   TempID9 = ACCEPT_BY
   If ACCEPT_BY <= 0 Then
      TempID9 = "NULL"
   End If
   
   TempID10 = RECEIVE_BY
   If RECEIVE_BY <= 0 Then
      TempID10 = "NULL"
   End If
   
   TempID11 = PAYMENT_TERM
   If PAYMENT_TERM <= 0 Then
      TempID11 = "NULL"
   End If
   
   TempID12 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID12 = "NULL"
   End If
   
   
   TempID13 = SUPPLIER_ID
   If SUPPLIER_ID <= 0 Then
      TempID13 = "NULL"
   End If
   
   TempID14 = RECEIPT_TYPE
   If RECEIPT_TYPE <= 0 Then
      TempID14 = "NULL"
   End If
   
   TempID15 = RESOURCE_ID
   If RESOURCE_ID <= 0 Then
      TempID15 = "NULL"
   End If
   
   TempID16 = BANK_ID
   If BANK_ID <= 0 Then
      TempID16 = "NULL"
   End If
   
   TempID17 = BBRANCH_ID
   If BBRANCH_ID <= 0 Then
      TempID17 = "NULL"
   End If
   
   TempID18 = PAYMENT_ID
   If PAYMENT_ID <= 0 Then
      TempID18 = "NULL"
   End If
   
   TempID20 = BATCH_ID
   If BATCH_ID <= 0 Then
      TempID20 = "NULL"
   End If
      
   TempID21 = DEPARTMENT_ID
   If DEPARTMENT_ID <= 0 Then
      TempID21 = "NULL"
   End If
   
   If EFFECTIVE_FLAG = "" Then
      EFFECTIVE_FLAG = "N"
   End If
   
   If PAY_FLAG = "" Then
      PAY_FLAG = "N"
   End If
   
'   SIMULATE_FLAG =
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILLING_DOC_SEQ", id, glbErrorLog)
               
      SQL1 = "INSERT INTO BILLING_DOC " & _
                     "( " & _
                     "   BILLING_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   NOTE, " & _
                     "   BILLING_ADDRESS_ID, " & _
                     "   ENTERPRISE_ADDRESS_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DUE_DATE, " & _
                     "   ACCOUNT_ID, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   TOTAL_AMOUNT, " & _
                     "   VAT_PERCENT, "
               SQL1 = SQL1 & _
                     "   VAT_AMOUNT, " & _
                     "   WH_PERCENT, " & _
                     "   WH_AMOUNT, " & _
                      "   ENTERPRISE_ID, " & _
                      "   ACCEPT_BY, " & _
                      "   RECEIVE_BY, " & _
                      "   PAYMENT_TERM, " & _
                      "   PAYEE_NAME, " & _
                      "   DISCOUNT_AMOUNT, " & _
                      "   EXCEPTION_FLAG, " & _
                      "   INVENTORY_DOC_ID, " & _
                      "   COMMIT_FLAG, " & _
                      "   SUPPLIER_ID, " & _
                      "   RECEIPT_TYPE, " & _
                      "   DEPOSIT_AMOUNT, " & _
                      "   TOTAL_PRICE, "
            SQL1 = SQL1 & _
                     "   ESTIMATE_DATE, " & _
                     "   APPROVE_DATE, " & _
                     "   RESOURCE_ID, " & _
                     "   PAYMENT_TYPE, " & _
                     "   BANK_ID, " & _
                     "   BBRANCH_ID, " & _
                     "   CHECK_NO, " & _
                     "   CHECK_DATE, " & _
                     "   VADILITY, " & _
                     "   DELIVERY, " & _
                     "  PAYMENT, " & _
                    "  PAYMENT_DESC, " & _
                    "  SHIPMENT, " & _
                    "  REFER_INV, " & _
                    "  PACKING_OF, " & _
                    "  SHIP_NAME, " & _
                    "  SHIPPING_MARKS, " & _
                    "  SHIPON_DATE, "
        SQL1 = SQL1 & _
                    "  REF, " & _
                    "  CREDIT, " & _
                    "  CD_AMOUNT, " & _
                    "  CD_PERCENT, " & _
                    "  PACKAGE_TYPE, " & _
                    "  TEMP_DO_NO, " & _
                    "  PAID_AMOUNT, " & _
                    "  DEBIT_AMOUNT, " & _
                    "  CREDIT_AMOUNT, " & _
                    "  PAYMENT_ID, " & _
                    "  BANK_NOTE, " & _
                    "  TOTAL_RCP, " & _
                    "  AGREEMENT_DATA, " & _
                    "  AGREEMENT_FINANCE, " & _
                    "  OLD_CREDIT_AMOUNT, " & _
                    "  ENTRY_DATE, " & _
                    "  EXIT_DATE, " & _
                     "  DO_NO, "
         SQL1 = SQL1 & _
                     "  TRUCK_NO, " & _
                     "  DELIVERY_FEE, " & _
                     "  SENDER_NAME, " & _
                     "  RECEIVE_NAME, " & _
                     "  DEPARTMENT_ID, " & _
                     "  QUE_NO, " & _
                     "  PR_NO, " & _
                     "  PV_NO, " & _
                     "  EFFECTIVE_FLAG, " & _
                     "  PAY_FLAG, " & _
                    "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         id & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "'," & _
                         TempID0 & "," & _
                         TempID3 & "," & _
                         TOTAL_AMOUNT & "," & _
                         TempID4 & ","
                  SQL1 = SQL1 & _
                         VAT_AMOUNT & "," & _
                         TempID6 & "," & _
                         TempID8 & "," & _
                         TempID5 & "," & _
                         TempID9 & "," & _
                         TempID10 & "," & _
                         TempID11 & "," & _
                         "'" & ChangeQuote(Trim(PAYEE_NAME)) & "'," & _
                         DISCOUNT_AMOUNT & "," & _
                         "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                         TempID12 & "," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                         TempID13 & "," & _
                         TempID14 & "," & _
                         DEPOSIT_AMOUNT & "," & _
                         TOTAL_PRICE & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ESTIMATE_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(APPROVE_DATE))) & "'," & _
                         TempID15 & ","
                  SQL1 = SQL1 & _
                         PAYMENT_TYPE & "," & _
                         TempID16 & "," & _
                         TempID17 & "," & _
                         "'" & ChangeQuote(Trim(CHECK_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(CHECK_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(VADILITY))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DELIVERY))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PAYMENT))) & "'," & _
                         "'" & ChangeQuote(Trim(PAYMENT_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(SHIPMENT))) & "'," & _
                         "'" & ChangeQuote(Trim(REFER_INV)) & "'," & _
                         "'" & ChangeQuote(PACKING_OF) & "'," & _
                         "'" & ChangeQuote(SHIP_NAME) & "'," & _
                         "'" & ChangeQuote(Trim(SHIPPING_MARKS)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(SHIPON_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(REF)) & "'," & _
                        Credit & "," & _
                        CD_AMOUNT & "," & _
                        CD_PERCENT & ","
            SQL1 = SQL1 & _
                        PACKAGE_TYPE & "," & _
                        "'" & ChangeQuote(Trim(TEMP_DO_NO)) & "'," & _
                        PAID_AMOUNT & "," & _
                        DEBIT_AMOUNT & "," & _
                        CREDIT_AMOUNT & "," & _
                        TempID18 & "," & _
                        "'" & ChangeQuote(Trim(BANK_NOTE)) & "'," & _
                        TOTAL_RCP & "," & _
                        "'" & ChangeQuote(Trim(AGREEMENT_DATA)) & "'," & _
                        "'" & ChangeQuote(Trim(AGREEMENT_FINANCE)) & "'," & _
                        OLD_CREDIT_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(EXIT_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(DO_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(TRUCK_NO)) & "'," & _
                        DELIVERY_FEE & "," & _
                        "'" & ChangeQuote(Trim(SENDER_NAME)) & "'," & _
                        "'" & ChangeQuote(Trim(RECEIVE_NAME)) & "'," & _
                        TempID21 & ","
            SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(QUE_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(PR_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(PV_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(EFFECTIVE_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(PAY_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BILLING_DOC_ID = id
   Else 'Show edit
      SQL1 = "UPDATE BILLING_DOC SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "    BILLING_ADDRESS_ID = " & TempID1 & ", " & _
                     "    ENTERPRISE_ADDRESS_ID = " & TempID2 & ", " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   DUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "', " & _
                     "    ACCOUNT_ID = " & TempID0 & ", " & _
                     "    DOCUMENT_TYPE = " & TempID3 & ", " & _
                     "    TOTAL_AMOUNT = " & TOTAL_AMOUNT & ", " & _
                     "    VAT_PERCENT = " & TempID4 & ", " & _
                     "    VAT_AMOUNT = " & VAT_AMOUNT & ", "
            SQL1 = SQL1 & _
                     "    WH_PERCENT = " & TempID6 & ", " & _
                     "    WH_AMOUNT = " & TempID8 & ", " & _
                     "    ENTERPRISE_ID = " & TempID5 & ", " & _
                     "    ACCEPT_BY = " & TempID9 & ", " & _
                     "    RECEIVE_BY = " & TempID10 & ", " & _
                     "    PAYMENT_TERM = " & TempID11 & ", " & _
                     "   PAYEE_NAME = '" & ChangeQuote(Trim(PAYEE_NAME)) & "', " & _
                     "    DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                     "   EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "    INVENTORY_DOC_ID = " & TempID12 & ", " & _
                     "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "    SUPPLIER_ID = " & TempID13 & ", " & _
                     "    RECEIPT_TYPE = " & TempID14 & ", " & _
                     "    DEPOSIT_AMOUNT = " & DEPOSIT_AMOUNT & ", " & _
                     "    TOTAL_PRICE = " & TOTAL_PRICE & ", " & _
                     "   ESTIMATE_DATE = '" & ChangeQuote(Trim(DateToStringInt(ESTIMATE_DATE))) & "', " & _
                     "   APPROVE_DATE = '" & ChangeQuote(Trim(DateToStringInt(APPROVE_DATE))) & "', " & _
                     "   RESOURCE_ID = " & TempID15 & ", "
            SQL1 = SQL1 & _
                     "    PAYMENT_TYPE = " & PAYMENT_TYPE & ", " & _
                     "    BANK_ID = " & TempID16 & ", " & _
                     "    BBRANCH_ID = " & TempID17 & ", " & _
                     "   CHECK_NO = '" & ChangeQuote(Trim(CHECK_NO)) & "', " & _
                     "   CHECK_DATE = '" & ChangeQuote(Trim(DateToStringInt(CHECK_DATE))) & "', " & _
                     "   VADILITY = '" & ChangeQuote(Trim(DateToStringInt(VADILITY))) & "', " & _
                     "   DELIVERY = '" & ChangeQuote(Trim(DateToStringInt(DELIVERY))) & "', " & _
                     "   PAYMENT = '" & ChangeQuote(Trim(DateToStringInt(PAYMENT))) & "', " & _
                     "   PAYMENT_DESC = '" & ChangeQuote(Trim(PAYMENT_DESC)) & "', " & _
                     "   SHIPMENT = '" & ChangeQuote(Trim(DateToStringInt(SHIPMENT))) & "', " & _
                     "   REFER_INV = '" & ChangeQuote(Trim(REFER_INV)) & "', " & _
                     "   PACKING_OF = '" & ChangeQuote(PACKING_OF) & "', " & _
                     "   SHIP_NAME = '" & ChangeQuote(SHIP_NAME) & "', " & _
                     "   SHIPPING_MARKS = '" & ChangeQuote(Trim(SHIPPING_MARKS)) & "', " & _
                     "   SHIPON_DATE = '" & ChangeQuote(Trim(DateToStringInt(SHIPON_DATE))) & "', " & _
                     "   REF = '" & ChangeQuote(Trim(REF)) & "', " & _
                    "    CREDIT = " & Credit & ", " & _
                    "    CD_AMOUNT = " & CD_AMOUNT & ", " & _
                    "    CD_PERCENT= " & CD_PERCENT & ", " & _
                    "    PACKAGE_TYPE= " & PACKAGE_TYPE & ", " & _
                    "   TEMP_DO_NO = '" & ChangeQuote(Trim(TEMP_DO_NO)) & "', "
         SQL1 = SQL1 & _
                     "    PAID_AMOUNT= " & PAID_AMOUNT & ", " & _
                     "    PAYMENT_ID= " & TempID18 & ", " & _
                     "    DEBIT_AMOUNT= " & DEBIT_AMOUNT & ", " & _
                     "    CREDIT_AMOUNT= " & CREDIT_AMOUNT & ", " & _
                     "   AGREEMENT_DATA = '" & ChangeQuote(AGREEMENT_DATA) & "', " & _
                     "   AGREEMENT_FINANCE = '" & ChangeQuote(AGREEMENT_FINANCE) & "', " & _
                     "   BANK_NOTE = '" & ChangeQuote(BANK_NOTE) & "', " & _
                     "    TOTAL_RCP= " & TOTAL_RCP & ", " & _
                     "    OLD_CREDIT_AMOUNT= " & OLD_CREDIT_AMOUNT & ", " & _
                     "   ENTRY_DATE = '" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "', " & _
                     "   EXIT_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXIT_DATE))) & "', " & _
                     "   DO_NO = '" & ChangeQuote(DO_NO) & "', " & _
                     "   TRUCK_NO = '" & ChangeQuote(TRUCK_NO) & "', " & _
                     "    DELIVERY_FEE= " & DELIVERY_FEE & ", " & _
                     "   SENDER_NAME = '" & ChangeQuote(SENDER_NAME) & "', " & _
                     "   RECEIVE_NAME = '" & ChangeQuote(RECEIVE_NAME) & "', " & _
                     "    DEPARTMENT_ID = " & TempID21 & ", " & _
                     "   QUE_NO = '" & ChangeQuote(QUE_NO) & "', " & _
                     "   PR_NO = '" & ChangeQuote(PR_NO) & "', " & _
                     "   PV_NO = '" & ChangeQuote(PV_NO) & "', " & _
                     "   EFFECTIVE_FLAG = '" & ChangeQuote(EFFECTIVE_FLAG) & "', "
         SQL1 = SQL1 & _
                     "   PAY_FLAG = '" & ChangeQuote(PAY_FLAG) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
Dim Pm As CPayment
Dim TempRs As ADODB.Recordset
Dim Ri As CReceiptItem
Dim iCount As Long
Dim ct As CCashTran
Dim Cq As CCheque
   
   If PAYMENT_ID > 0 Then
      Set Pm = New CPayment
      Pm.PAYMENT_ID = PAYMENT_ID
      Call Pm.DeleteData
      Set Pm = Nothing
   End If
      
   Set TempRs = New ADODB.Recordset
   Set Ri = New CReceiptItem
   Ri.BILLING_DOC_ID = BILLING_DOC_ID
   Call Ri.QueryData(1, TempRs, iCount)
   While Not TempRs.EOF
      Call Ri.PopulateFromRS(1, TempRs)
      Call Ri.DeleteData
      
      TempRs.MoveNext
   Wend
   Set Ri = Nothing
   Set TempRs = Nothing
   
      SQL1 = "DELETE FROM PKGLST_ITEM WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM RECEIPT_ITEM WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM SALE_ORDER WHERE DO_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM SUP_ITEM WHERE DO_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM DO_ITEM WHERE DO_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM BILLING_DISCOUNT WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM BULK_HOLE WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)

'      SQL1 = "DELETE FROM RO_ITEM WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
'      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM GL_DETAIL WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM BILLING_ADJUST WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
   Set TempRs = New ADODB.Recordset
   Set ct = New CCashTran
   Call ct.SetFieldValue("BILLING_DOC_ID", BILLING_DOC_ID)
   Call ct.QueryData(1, TempRs, iCount)
   Set ct = Nothing
   While Not TempRs.EOF
      Set ct = New CCashTran
      Call ct.PopulateFromRS(1, TempRs)
      Call ct.DeleteData
      If ct.GetFieldValue("CHECK_ID") > 0 Then
         Set Cq = New CCheque
         Call Cq.SetFieldValue("CHEQUE_ID", ct.GetFieldValue("CHECK_ID"))
         Call Cq.DeleteData
         Set Cq = Nothing
      End If
      Set ct = Nothing
      Call TempRs.MoveNext
   Wend
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
      
      SQL1 = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID = " & BILLING_DOC_ID & " "
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set DoItems = Nothing
   Set SaleOrders = Nothing
   Set RoItems = Nothing
   Set InventoryDocs = Nothing
   Set ReceiptItems = Nothing
   Set Pkglsts = Nothing
   Set MemoBank = Nothing
   Set BillingDiscounts = Nothing
   Set BulkHoles = Nothing
   Set Payments = Nothing
   Set GlDetails = Nothing
End Sub

Public Function UpdatePaidAmount() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   PAID_AMOUNT =  (SELECT SUM(RI.PAID_AMOUNT) FROM RECEIPT_ITEM RI WHERE (RI.DO_ID = " & BILLING_DOC_ID & ")), " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID

   m_Conn.Execute (SQL1)
End Function

Public Function UpdateCnDnAmount() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   DEBIT_AMOUNT =  (SELECT SUM(RI.DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI, BILLING_DOC BD1 WHERE (RI.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 4) AND (RI.DO_ID = " & BILLING_DOC_ID & ")), " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)

   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   CREDIT_AMOUNT =  (SELECT SUM(RI.DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI, BILLING_DOC BD1 WHERE (RI.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 3) AND (RI.DO_ID = " & BILLING_DOC_ID & ")), " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)
End Function

Public Function UpdateNewAccount()
Dim ServerDateTime As String
Dim SQL1 As String
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
   Else 'Show edit
      SQL1 = "UPDATE BILLING_DOC SET " & _
                     "    ACCOUNT_ID = " & NEW_ACCOUNT_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ACCOUNT_ID = " & ACCOUNT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function


