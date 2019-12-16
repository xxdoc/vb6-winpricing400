VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReceiptItem"
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
' End of generic data member

Public RECEIPT_ITEM_ID As Long
Public BILLING_DOC_ID As Long
Public DO_ID As Long
Public RECEIPT_ITEM_AMOUNT As Double
Public VAT_AMOUNT As Double
Public DISCOUNT_AMOUNT As Double
Public DEPOSIT_AMOUNT As Double
Public DRCR_REASON As Long
Public CASH_DISCOUNT As Double
Public EFFECTIVE_FLAG As String

Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public PAYMENT As Date
Public PAID_AMOUNT As Double
Public DEBIT_CREDIT_AMOUNT As Double

Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_ITEM_DATE As Date
Public TO_ITEM_DATE As Date
Public FROM_ITEM_DUE_DATE As Date
Public TO_ITEM_DUE_DATE As Date
Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date
Public CUSTOMER_ID As Long
Public ACCOUNT_ID As Long
Public DOCUMENT_SUBTYPE As Long
Public RECEIPT_TYPE As Long
Public RECEIPT_NO As String
Public RECEIPT_DATE As Date
Public BILL_NO As String
Public BILL_DATE As Date
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public SUPPLIER_TYPE_NO As String
Public SUPPLIER_TYPE_NAME As String
Public SUPPLIER_ID As Long
Public SUPPLIER_TYPE As Long
Public SUPPLIER_GRADE As Long
Public PAY_FLAG As String

Public DRCR_REASON_DESC As String

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM
Public LINK_ID As Long

Public PART_ITEM_ID As Long

Public PART_NO As String
Public PART_DESC As String
Public FEATURE_CODE As String
Public FEATURE_DESC As String
Public ITEM_DESC As String
Public LOCATION_NAME As String

Public RETURN_AMOUNT As Double
Public LOCATION_ID As Long
Public PART_TYPE As Long
Public RETURN_TOTAL_PRICE As Double
Public RETURN_AVG_PRICE As Double
Public AVG_PRICE As Double
Public AVG_WEIGHT As Double
Public FEATURE_ID As Long
Public RETURN_DISCOUNT_AMOUNT As Double
Public CONFIG_CODE As String
Public DISPLAY_ID As Long
Public COUNTRY_CURRENCY1 As Long
Public COUNTRY_CURRENCY2 As Long
Public WEIGHT_PER_PACK As Double
Public PACK_AMOUNT As Double
Public PRICE_PER_PACK As Double
Public DISCOUNT_PER_PACK As Double
Public MANUAL_FLAG As String
Public MANUAL_CODE As String
Public MANUAL_NAME As String
Public RATE_TYPE As Long
Public TRANSFER_WAGE As Double
Public STD_TRANSFER_CHARGE As Double
Public MAX_DOC_DATE As Date
Public MIN_DOC_DATE As Date
Public SALE_NAME As String
Public SALE_LNAME As String
Public LOCATION_NO As String

Public DocTypeSet As String


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM

'--------------------------------------------------------------------------------------------------------------------------------> FOR FrmExportItem
'For Export  tabel BillingDoc
Public DUE_DATE As Date
Public NOTE As String
Public ENTERPRISE_ID As Long
Public BILLING_ADDRESS_ID As Long
Public ENTERPRISE_ADDRESS_ID As Long
Public TOTAL_AMOUNT As Double
Public VAT_PERCENT As Double
Public WH_PERCENT As Double
Public WH_AMOUNT As Double
Public PAYMENT_TERM As Long
Public ACCEPT_BY As Long
Public RECEIVE_BY As Long
Public EXCEPTION_FLAG As String
Public PAYEE_NAME As String
Public INVENTORY_DOC_ID As Long
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


Public SUM_DISCOUNT_AMOUNT As Double
Public SUM_PAID_AMOUNT As Double
Public TOTAL_DEBIT_AMOUNT As Double
Public TOTAL_CREDIT_AMOUNT As Double
Public SUM_CASH_DISCOUNT As Double
Public OLD_COMMIT_FLAG As String
Public ACCOUNT_NO As String
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public RECEIPT_REF_COUNT As Long
Public DBN_REF_COUNT As Long
Public CDN_REF_COUNT As Long
Public BLS_REF_COUNT As Long
Public RECEIPT_REF_COUNT2 As Long
Public DBN_REF_COUNT2 As Long
Public CDN_REF_COUNT2 As Long
Public TAX_FLAG As String
Public WH_FLAG As String
Public EMP_CODE As String
Public EMP_NAME As String
Public EMP_LNAME As String
Public RESOURCE_NO As String
Public VADILITY As Date
Public DELIVERY As Date
Public PAYMENT_DESC As String
Public SHIPMENT As Date
Public TEMP_PAID_AMOUNT As Double
Public TEMP_CHECK_DATE As String
Public TEMP_CHECK_NO As String
Public PAID_TYPE As Long
Public RECEIPT_PAID_AMOUNT As Double
Public REASON_ID As Long
Public DO_RECEIPT_FLAG As String
Public CUSTOMER_GRADE As Long

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

Public COMMIT_FLAG As String
'--------------------------------------------------------------------------------------------------------------------------------> FOR FrmExportItem

Public DRCR_REASON_NO As String
Public CHEQUE_DOC_ID As Long
Public YYYYMM As String

Public ITEM_AMOUNT As Double
Public TOTAL_WEIGHT As Double
Public BEFORE_DISCOUNT_PRICE As Double
Public EXTRA_DISCOUNT As Double

Public FREELANCE_ID As Long
Public FREELANCE_CODE As String
Public FROM_FREELANCE_CODE As String
Public TO_FREELANCE_CODE As String
Public FREELANCE_NAME As String
Public FREELANCE_LASTNAME As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_ITEM_DATE = -1
   TO_ITEM_DATE = -1
   FROM_ITEM_DUE_DATE = -1
   TO_ITEM_DUE_DATE = -1
   EFFECTIVE_FLAG = ""
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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, IVD1.DOCUMENT_DATE, IVD1.DOCUMENT_NO, IVD1.COMMIT_FLAG, IVD1.DOCUMENT_TYPE, IVD1.VAT_AMOUNT ITEM_VAT_AMOUNT, IVD2.DOCUMENT_TYPE ITEM_DOCUMENT_TYPE,IVD1.PAYMENT "
      SelectStr1 = SelectStr1 & ",PT.PART_NO,PT.PART_DESC,LC.LOCATION_NAME,FT.FEATURE_CODE,FT.FEATURE_DESC "
      SelectStr1 = SelectStr1 & ",MR.KEY_NAME  DRCR_REASON_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PT ON (UG.PART_ITEM_ID = PT.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (UG.DRCR_REASON = MR.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT IVD2.ACCOUNT_ID, IVD2.DOCUMENT_TYPE, IVD2.DOCUMENT_SUBTYPE, IVD2.RECEIPT_TYPE, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      If Len(FROM_CUSTOMER_CODE) > 0 Or Len(TO_CUSTOMER_CODE) > 0 Then
         FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      End If
      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NULL) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT CM2.CUSTOMER_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NULL) "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT IVD2.ACCOUNT_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD1.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD2.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT AC2.CUSTOMER_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD1.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT IVD2.DOCUMENT_NO RECEIPT_NO, IVD2.DOCUMENT_DATE RECEIPT_DATE, IVD2.DOCUMENT_TYPE, IVD1.DOCUMENT_NO BILL_NO, IVD1.DOCUMENT_DATE BILL_DATE, UG.*, CM2.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME, UG.DEBIT_CREDIT_AMOUNT, ENM.LONG_NAME SALE_NAME, ENM.LAST_NAME SALE_LNAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD1.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (CM2.RESPONSE_BY = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT IVD2.DOCUMENT_DATE, CM2.CUSTOMER_ID, SUM(UG.RECEIPT_ITEM_AMOUNT) RECEIPT_ITEM_AMOUNT, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD2.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT IVD2.BILLING_DOC_ID, SUM(UG.RECEIPT_ITEM_AMOUNT) RECEIPT_ITEM_AMOUNT, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD2.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT UG.DO_ID, UG.BILLING_DOC_ID, UG.PAID_AMOUNT, UG.RECEIPT_ITEM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD2.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT MAX(IVD2.DOCUMENT_DATE) MAX_DOC_DATE, MIN(IVD2.DOCUMENT_DATE) MIN_DOC_DATE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD2.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT IVD1.DOCUMENT_NO BILL_NO, IVD1.DOCUMENT_DATE BILL_DATE,UG.PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 100 Then
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      
      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 101 Then
      SelectStr1 = "SELECT IVD2.SUPPLIER_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD2.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 102 Then
      SelectStr1 = "SELECT IVD2.PAY_FLAG, UG.BILLING_DOC_ID, IVD2.DOCUMENT_NO, IVD2.DOCUMENT_DATE, IVD2.SUPPLIER_ID, SP.SUPPLIER_CODE, NM.LONG_NAME SUPPLIER_NAME, ST.SUPPLIER_TYPE_NO, ST.SUPPLIER_TYPE_NAME, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD2.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_TYPE ST ON (ST.SUPPLIER_TYPE_ID = SP.SUPPLIER_TYPE) "
      
      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 103 Then
      SelectStr1 = "SELECT IVD2.SUPPLIER_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD2.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "

      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 104 Then
      SelectStr1 = "SELECT IVD2.SUPPLIER_ID, IVD2.DOCUMENT_TYPE, IVD2.DOCUMENT_SUBTYPE, IVD2.RECEIPT_TYPE, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD2.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = ""
      
   ElseIf Ind = 105 Then            ' สำหรับ Export ไปยังสำนักงานใหญ่
      SelectStr1 = "SELECT UG.*,IVD2.* "
      SelectStr1 = SelectStr1 & ", IVD1.DOCUMENT_NO BILL_NO"
      SelectStr1 = SelectStr1 & ",PT.PART_NO,FT.FEATURE_CODE "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE DRCR_REASON_NO "
      SelectStr1 = SelectStr1 & ",LC.LOCATION_NO  "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD2.SUPPLIER_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PT ON (UG.PART_ITEM_ID = PT.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (UG.DRCR_REASON = MR.KEY_ID) "
   ElseIf Ind = 106 Then
      SelectStr1 = "SELECT IVD2.SUPPLIER_ID, IVD2.DOCUMENT_TYPE, IVD2.DOCUMENT_SUBTYPE, IVD2.RECEIPT_TYPE, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CASH_TRAN CT ON (CT.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD2.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = " WHERE (CT.PAYMENT_TYPE =  1 OR CT.PAYMENT_TYPE =  2) "
   ElseIf Ind = 107 Then  'สำหรับ ประมวลผลประจำปี FrmInitBalance
      SelectStr1 = "SELECT DISTINCT UG.DO_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
   ElseIf Ind = 108 Then  'สำหรับ ตรวจเลขที่เอกสาร
      SelectStr1 = "SELECT IVD2.DOCUMENT_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (IVD2.BILLING_DOC_ID = UG.BILLING_DOC_ID) "
   ElseIf Ind = 115 Then            'Function LoadPaidAmountByCustomer2 จาก CheckMaxCreditDueDate ใน frmAddEditDo
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 116 Then         'Function LoadDnCnAmountByCustomer2 จาก CheckMaxCreditDueDate ใน frmAddEditDo
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (IVD2.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
   ElseIf Ind = 117 Then   'LoadDnCnAmountBySale
      SelectStr1 = "SELECT IVD1.ACCEPT_BY,IVD2.YYYYMM, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
   ElseIf Ind = 120 Then
      SelectStr1 = "SELECT IVD2.DOCUMENT_NO RECEIPT_NO, IVD2.DOCUMENT_DATE RECEIPT_DATE,UG.RECEIPT_ITEM_AMOUNT,UG.PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",IVD1.DOCUMENT_NO BILL_NO, IVD1.DOCUMENT_DATE BILL_DATE,IVD1.CREDIT "
      SelectStr1 = SelectStr1 & ",IVD1.ACCEPT_BY,EM.EMP_CODE,NM.LONG_NAME SALE_NAME, NM.LAST_NAME SALE_LNAME "
      SelectStr1 = SelectStr1 & ",CM.CUSTOMER_ID,CM.CUSTOMER_CODE, NM2.LONG_NAME CUSTOMER_NAME "
      
      SelectStr1 = SelectStr1 & ", DI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ", DI.ITEM_AMOUNT, DI.TOTAL_WEIGHT"
      SelectStr1 = SelectStr1 & ", DI.TOTAL_PRICE, DI.PACK_AMOUNT, DI.DISCOUNT_AMOUNT, DI.EXTRA_DISCOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DO_ITEM DI ON (IVD1.BILLING_DOC_ID = DI.DO_ID) "
       
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (CN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (IVD2.ACCEPT_BY = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = EN.NAME_ID) "
      
      WhereStr = "WHERE (CM.CUSTOMER_ID IS NOT NULL AND EM.EMP_ID IS NOT NULL AND PI.PART_NO IS NOT NULL) "
      WhereStr = WhereStr & " AND (PT.COMMISSION_FLAG = 'Y') "
   ElseIf Ind = 121 Then
      SelectStr1 = "SELECT IVD1.BILLING_DOC_ID,IVD2.DOCUMENT_NO RECEIPT_NO, IVD2.DOCUMENT_DATE RECEIPT_DATE,UG.RECEIPT_ITEM_AMOUNT,UG.PAID_AMOUNT "
      SelectStr1 = SelectStr1 & ",IVD1.DOCUMENT_NO BILL_NO, IVD1.DOCUMENT_DATE BILL_DATE,IVD1.CREDIT "
      SelectStr1 = SelectStr1 & ",IVD1.ACCEPT_BY,FL.FREELANCE_ID,FL.FREELANCE_CODE,FL.FREELANCE_NAME, FREELANCE_LASTNAME "
      SelectStr1 = SelectStr1 & ",CM2.CUSTOMER_ID,CM2.CUSTOMER_CODE, NM2.LONG_NAME CUSTOMER_NAME "
      
      SelectStr1 = SelectStr1 & ", DI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ", DI.ITEM_AMOUNT, DI.TOTAL_WEIGHT"
      SelectStr1 = SelectStr1 & ", DI.TOTAL_PRICE, DI.PACK_AMOUNT, DI.DISCOUNT_AMOUNT, DI.EXTRA_DISCOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD1 ON (UG.DO_ID = IVD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN DO_ITEM DI ON (IVD1.BILLING_DOC_ID = DI.DO_ID) "
       
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (CN.NAME_ID = NM2.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN FREELANCE_ITEM FLI ON (CM2.CUSTOMER_ID = FLI.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FREELANCE FL ON (FLI.FREELANCE_ID = FL.FREELANCE_ID) "
      
      WhereStr = "WHERE (CM2.CUSTOMER_ID IS NOT NULL AND FL.FREELANCE_ID IS NOT NULL AND PI.PART_NO IS NOT NULL) "
     
       WhereStr = WhereStr & " AND (PT.COMMISSION_FLAG = 'Y') "
   ElseIf Ind = 122 Then
      SelectStr1 = "SELECT CM2.CUSTOMER_ID,IVD2.DOCUMENT_DATE,SUM(UG.RECEIPT_ITEM_AMOUNT) RECEIPT_ITEM_AMOUNT, SUM(UG.PAID_AMOUNT) PAID_AMOUNT, SUM(UG.CASH_DISCOUNT) CASH_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD2.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = "WHERE (IVD2.SUPPLIER_ID IS NULL) "
     ElseIf Ind = 123 Then
      SelectStr1 = "SELECT UG.*,IVD2.DOCUMENT_NO "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
   End If
   
   If RECEIPT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID & ") "
      End If
   End If
   
   If Ind = 1 Or Ind = 15 Then
      If BILLING_DOC_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (UG.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
         Else
            WhereStr = WhereStr & "AND (UG.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
         End If
      End If
   Else
      If BILLING_DOC_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD2.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
         Else
            WhereStr = WhereStr & "AND (IVD2.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
         End If
      End If
   End If
   
   If Ind = 108 Or Ind = 123 Then
      If DO_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (UG.DO_ID = " & DO_ID & ") "
         Else
            WhereStr = WhereStr & "AND (UG.DO_ID = " & DO_ID & ") "
         End If
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD2.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (IVD2.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
            
   If FROM_ITEM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_ITEM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD1.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD1.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_ITEM_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_ITEM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD1.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD1.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_ITEM_DUE_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_ITEM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD1.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD1.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_ITEM_DUE_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_ITEM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD1.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD1.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
    If FROM_DOC_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD2.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD2.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DOC_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD2.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD2.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Len(EFFECTIVE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD2.EFFECTIVE_FLAG = '" & EFFECTIVE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (IVD2.EFFECTIVE_FLAG = '" & EFFECTIVE_FLAG & "') "
      End If
   End If
      
   
   If CUSTOMER_ID > 0 Then
      If Ind = 15 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD1.CUSTOMER_ID = " & CUSTOMER_ID & ") "
         Else
            WhereStr = WhereStr & "AND (IVD1.CUSTOMER_ID = " & CUSTOMER_ID & ") "
         End If
      Else
         'If Ind = 15 Then   'แก้ comment ไว้ เนื่องจากวงเงิน DEBIT_CREDIT_AMOUNT เมื่อtest customer ใหม่เข้ามาแล้วมี  DEBIT_CREDIT_AMOUNTการลด การเพิ่มหนี้ ซึ่งไม่ควรมี จะทำการ lock วงเงิน วงวัน ใน frmAddEditDO
            If Len(WhereStr) <= 0 Then
               WhereStr = "WHERE (CM2.CUSTOMER_ID = " & CUSTOMER_ID & ") "
               'WhereStr = "WHERE (IVD2.CUSTOMER_ID = " & CUSTOMER_ID & ") "
            Else
               WhereStr = WhereStr & "AND (CM2.CUSTOMER_ID = " & CUSTOMER_ID & ") "
               'WhereStr = WhereStr & "AND (IVD2.CUSTOMER_ID = " & CUSTOMER_ID & ") "
            End If
         'End If
      End If
   End If
   
   If Trim(CUSTOMER_CODE) <> "" Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CM2.CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CM2.CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM2.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM2.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
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
   
      If Trim(FREELANCE_CODE) <> "" Then
      If WildCard(FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FL.FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FL.FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(FROM_FREELANCE_CODE) <> "" Then
      If WildCard(FROM_FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FROM_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FROM_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FL.FREELANCE_CODE >= '" & ChangeQuote(Trim(FROM_FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FL.FREELANCE_CODE >= '" & ChangeQuote(Trim(FROM_FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(TO_FREELANCE_CODE) <> "" Then
      If WildCard(TO_FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(TO_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(TO_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FL.FREELANCE_CODE <= '" & ChangeQuote(Trim(TO_FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FL.FREELANCE_CODE <= '" & ChangeQuote(Trim(TO_FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD2.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (IVD2.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      End If
   End If
   
   If Trim(SUPPLIER_CODE) <> "" Then
      If WildCard(SUPPLIER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SP.SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SP.SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SP.SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SP.SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
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
   
   GroupByStr = ""
   If Ind = 2 Then
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY IVD2.ACCOUNT_ID, IVD2.DOCUMENT_TYPE, IVD2.DOCUMENT_SUBTYPE, IVD2.RECEIPT_TYPE "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY CM2.CUSTOMER_ID "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY IVD2.ACCOUNT_ID "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY AC2.CUSTOMER_ID "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY IVD2.DOCUMENT_DATE, CM2.CUSTOMER_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY IVD2.BILLING_DOC_ID "
   ElseIf Ind = 100 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 101 Then
      GroupByStr = "GROUP BY IVD2.SUPPLIER_ID "
   ElseIf Ind = 102 Then
      GroupByStr = "GROUP BY IVD2.PAY_FLAG, UG.BILLING_DOC_ID, IVD2.DOCUMENT_NO, IVD2.DOCUMENT_DATE, IVD2.SUPPLIER_ID, SP.SUPPLIER_CODE, NM.LONG_NAME, ST.SUPPLIER_TYPE_NO, ST.SUPPLIER_TYPE_NAME "
   ElseIf Ind = 103 Then
      GroupByStr = "GROUP BY IVD2.SUPPLIER_ID "
   ElseIf Ind = 104 Or Ind = 106 Then
      GroupByStr = "GROUP BY IVD2.SUPPLIER_ID , IVD2.DOCUMENT_TYPE, IVD2.DOCUMENT_SUBTYPE, IVD2.RECEIPT_TYPE "
   ElseIf Ind = 115 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 116 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 117 Then
      GroupByStr = "GROUP BY IVD1.ACCEPT_BY , IVD2.YYYYMM "
   ElseIf Ind = 122 Then
      GroupByStr = "GROUP BY CM2.CUSTOMER_ID,IVD2.DOCUMENT_DATE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LOCATION_ID DESC, DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY LC.LOCATION_NO "
      Else
         OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      'OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY CUSTOMER_ID "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY IVD2.ACCOUNT_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY AC2.CUSTOMER_ID "
   ElseIf Ind = 10 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD1.DOCUMENT_NO ASC, IVD1.DOCUMENT_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CM2.CUSTOMER_CODE ASC, IVD2.DOCUMENT_DATE ASC, IVD2.DOCUMENT_NO, IVD1.DOCUMENT_DATE ASC, IVD1.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY CM2.CUSTOMER_CODE ASC, IVD2.DOCUMENT_DATE ASC, IVD2.DOCUMENT_NO, IVD1.DOCUMENT_DATE ASC, IVD1.DOCUMENT_NO "
      End If
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY IVD2.DOCUMENT_DATE "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY IVD2.BILLING_DOC_ID "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 14 Then
      'OrderByStr = "ORDER BY UG.RECEIPT_ITEM_ID "
   ElseIf Ind = 100 Then
      'OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
   ElseIf Ind = 101 Then
      OrderByStr = "ORDER BY IVD2.SUPPLIER_ID "
   ElseIf Ind = 102 Then
      OrderByStr = "ORDER BY ST.SUPPLIER_TYPE_NO ASC, SP.SUPPLIER_CODE ASC, IVD2.DOCUMENT_DATE ASC, IVD2.DOCUMENT_NO "
   ElseIf Ind = 103 Then
      OrderByStr = "ORDER BY IVD2.SUPPLIER_ID "
   ElseIf Ind = 104 Or Ind = 106 Then
      OrderByStr = "ORDER BY IVD2.SUPPLIER_ID "
   ElseIf Ind = 105 Then
      OrderByStr = "ORDER BY IVD2.DOCUMENT_TYPE,IVD2.DOCUMENT_DATE,IVD2.DOCUMENT_NO "
   ElseIf Ind = 120 Then
      OrderByStr = "ORDER BY EM.EMP_CODE ASC,IVD2.DOCUMENT_NO ASC, IVD2.DOCUMENT_DATE ASC,IVD1.DOCUMENT_NO ASC, IVD1.DOCUMENT_DATE "
   ElseIf Ind = 121 Then
      OrderByStr = "ORDER BY FL.FREELANCE_CODE ASC,CM2.CUSTOMER_CODE ASC,IVD2.DOCUMENT_NO ASC, IVD2.DOCUMENT_DATE ASC,IVD1.DOCUMENT_NO ASC, IVD1.DOCUMENT_DATE "
   End If

   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If OrderByStr = "" Then
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RECEIPT_ITEM_ID = NVLI(Rs("RECEIPT_ITEM_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      VAT_AMOUNT = NVLD(Rs("ITEM_VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
      DRCR_REASON = NVLI(Rs("DRCR_REASON"), -1)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      PAYMENT = InternalDateToDate(NVLS(Rs("PAYMENT"), ""))
      
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      RETURN_AMOUNT = NVLD(Rs("RETURN_AMOUNT"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      
      RETURN_TOTAL_PRICE = NVLD(Rs("RETURN_TOTAL_PRICE"), 0)
      RETURN_AVG_PRICE = NVLD(Rs("RETURN_AVG_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      AVG_WEIGHT = NVLD(Rs("AVG_WEIGHT"), 0)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      RETURN_DISCOUNT_AMOUNT = NVLD(Rs("RETURN_DISCOUNT_AMOUNT"), 0)
      CONFIG_CODE = NVLS(Rs("CONFIG_CODE"), "")
      DISPLAY_ID = NVLI(Rs("DISPLAY_ID"), -1)
      
      COUNTRY_CURRENCY1 = NVLD(Rs("COUNTRY_CURRENCY1"), 0)
      COUNTRY_CURRENCY2 = NVLD(Rs("COUNTRY_CURRENCY2"), 0)
      WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      PRICE_PER_PACK = NVLD(Rs("PRICE_PER_PACK"), 0)
      DISCOUNT_PER_PACK = NVLD(Rs("DISCOUNT_PER_PACK"), 0)
      
      MANUAL_FLAG = NVLS(Rs("MANUAL_FLAG"), "")
      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")
      RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
      TRANSFER_WAGE = NVLD(Rs("TRANSFER_WAGE"), 0)
      STD_TRANSFER_CHARGE = NVLD(Rs("STD_TRANSFER_CHARGE"), 0)
      
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")
      DRCR_REASON_DESC = NVLS(Rs("DRCR_REASON_DESC"), "")
   ElseIf Ind = 3 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
   ElseIf Ind = 4 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
'If CASH_DISCOUNT > 0 Then
''Debug.Print
'End If
   ElseIf Ind = 5 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
   ElseIf Ind = 6 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 7 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 9 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 10 Then
      RECEIPT_NO = NVLS(Rs("RECEIPT_NO"), "")
      RECEIPT_DATE = InternalDateToDate(NVLS(Rs("RECEIPT_DATE"), ""))
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      BILL_DATE = InternalDateToDate(NVLS(Rs("BILL_DATE"), ""))
      
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
      
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      SALE_LNAME = NVLS(Rs("SALE_LNAME"), "")
      SALE_NAME = SALE_NAME & " " & SALE_LNAME
   ElseIf Ind = 11 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 13 Then
      DO_ID = NVLI(Rs("DO_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 14 Then
      MAX_DOC_DATE = InternalDateToDate(NVLS(Rs("MAX_DOC_DATE"), ""))
      MIN_DOC_DATE = InternalDateToDate(NVLS(Rs("MIN_DOC_DATE"), ""))
   ElseIf Ind = 15 Then
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      BILL_DATE = InternalDateToDate(NVLS(Rs("BILL_DATE"), ""))
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 100 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
   ElseIf Ind = 101 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
   ElseIf Ind = 102 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), 0)
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      SUPPLIER_TYPE_NO = NVLS(Rs("SUPPLIER_TYPE_NO"), "")
      SUPPLIER_TYPE_NAME = NVLS(Rs("SUPPLIER_TYPE_NAME"), "")
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = NVLS(Rs("DOCUMENT_DATE"), "")
      PAY_FLAG = NVLS(Rs("PAY_FLAG"), "")
   ElseIf Ind = 103 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 104 Or Ind = 106 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
   ElseIf Ind = 105 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      NOTE = NVLS(Rs("NOTE"), "")
      BILLING_ADDRESS_ID = NVLI(Rs("BILLING_ADDRESS_ID"), 0)
      ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), 0)
      ENTERPRISE_ADDRESS_ID = NVLI(Rs("ENTERPRISE_ADDRESS_ID"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      WH_PERCENT = NVLD(Rs("WH_PERCENT"), 0)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      
      PAYMENT_TERM = NVLI(Rs("PAYMENT_TERM"), 0)
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), 0)
      RECEIVE_BY = NVLI(Rs("RECEIVE_BY"), 0)
      
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      
      '"INVENTORY_DOC_ID" "ID_TYPE",
      
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      
      APPROVE_DATE = InternalDateToDate(NVLS(Rs("APPROVE_DATE"), ""))
      ESTIMATE_DATE = InternalDateToDate(NVLS(Rs("ESTIMATE_DATE"), ""))
      
      RESOURCE_ID = NVLI(Rs("RESOURCE_ID"), 0)
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), 0)
      
      BANK_ID = NVLI(Rs("BANK_ID"), 0)
      BBRANCH_ID = NVLI(Rs("BBRANCH_ID"), 0)
      
      CHECK_NO = NVLS(Rs("CHECK_NO"), "")
      CHECK_DATE = InternalDateToDate(NVLS(Rs("CHECK_DATE"), ""))
      VADILITY = InternalDateToDate(NVLS(Rs("VADILITY"), ""))
      DELIVERY = InternalDateToDate(NVLS(Rs("DELIVERY"), ""))
      PAYMENT = InternalDateToDate(NVLS(Rs("PAYMENT"), ""))
      
      PAYMENT_DESC = NVLS(Rs("PAYMENT_DESC"), "")
      SHIPMENT = (NVLS(Rs("SHIPMENT"), ""))
      REFER_INV = NVLS(Rs("REFER_INV"), "")
      PACKING_OF = NVLS(Rs("PACKING_OF"), "")
      SHIPON_DATE = InternalDateToDate(NVLS(Rs("SHIPON_DATE"), ""))
      REF = NVLS(Rs("REF"), "")
      
      Credit = NVLD(Rs("CREDIT"), 0)
      'NUM_NO = NVLI(Rs("NUM_NO"), 0)
      AREA_CODE = NVLS(Rs("AREA_CODE"), "")
      'PO_DATE = InternalDateToDate(NVLS(Rs("PO_DATE"), ""))
      
      SHIP_NAME = NVLS(Rs("SHIP_NAME"), "")
      SHIPPING_MARKS = NVLS(Rs("SHIPPING_MARKS"), "")
      CD_PERCENT = NVLD(Rs("CD_PERCENT"), 0)
      CD_AMOUNT = NVLD(Rs("CD_AMOUNT"), 0)
      PACKAGE_TYPE = NVLI(Rs("PACKAGE_TYPE"), 0)
      TEMP_DO_NO = NVLS(Rs("TEMP_DO_NO"), "")
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      DO_TOTAL_PRICE = NVLD(Rs("DO_TOTAL_PRICE"), 0)
      REVENUE_TOTAL_PRICE = NVLD(Rs("REVENUE_TOTAL_PRICE"), 0)
      'DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      PAYMENT_ID = NVLI(Rs("PAYMENT_ID"), 0)
      BANK_BRANCH_ID = NVLI(Rs("BANK_BRANCH_ID"), 0)
      
      BANK_NOTE = NVLS(Rs("BANK_NOTE"), "")
      TOTAL_RCP = NVLD(Rs("TOTAL_RCP"), 0)
      RUNNING_NO = NVLI(Rs("RUNNING_NO"), 0)
      AGREEMENT_DATA = NVLS(Rs("AGREEMENT_DATA"), "")
      AGREEMENT_FINANCE = NVLS(Rs("AGREEMENT_FINANCE"), "")
      
      OLD_CREDIT_AMOUNT = NVLD(Rs("OLD_CREDIT_AMOUNT"), 0)
      ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
      EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
      BILL_NO = NVLS(Rs("BILL_NO"), "")
            
      '---------------------'---------------------'---------------------'---------------------'---------------------'--------------------->
      
      RECEIPT_ITEM_ID = NVLI(Rs("RECEIPT_ITEM_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)                                                                                                                               'ใช้เป็น NO
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      'VAT_AMOUNT = NVLD(Rs("ITEM_VAT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
      DRCR_REASON = NVLI(Rs("DRCR_REASON"), -1)                                                                                        ''ใช้เป็น NO
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      PAYMENT = InternalDateToDate(NVLS(Rs("PAYMENT"), ""))
      
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)                                                                                      'ใช้เป็น NO
      RETURN_AMOUNT = NVLD(Rs("RETURN_AMOUNT"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)                                                                                        'ใช้เป็น NO
      
      RETURN_TOTAL_PRICE = NVLD(Rs("RETURN_TOTAL_PRICE"), 0)
      RETURN_AVG_PRICE = NVLD(Rs("RETURN_AVG_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      AVG_WEIGHT = NVLD(Rs("AVG_WEIGHT"), 0)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)                                                                                                 'ใช้เป็น NO
      RETURN_DISCOUNT_AMOUNT = NVLD(Rs("RETURN_DISCOUNT_AMOUNT"), 0)
      CONFIG_CODE = NVLS(Rs("CONFIG_CODE"), "")
      DISPLAY_ID = NVLI(Rs("DISPLAY_ID"), -1)
      
      COUNTRY_CURRENCY1 = NVLD(Rs("COUNTRY_CURRENCY1"), 0)
      COUNTRY_CURRENCY2 = NVLD(Rs("COUNTRY_CURRENCY2"), 0)
      WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      PRICE_PER_PACK = NVLD(Rs("PRICE_PER_PACK"), 0)
      DISCOUNT_PER_PACK = NVLD(Rs("DISCOUNT_PER_PACK"), 0)
      
      MANUAL_FLAG = NVLS(Rs("MANUAL_FLAG"), "")
      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")
      RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
      TRANSFER_WAGE = NVLD(Rs("TRANSFER_WAGE"), 0)
      STD_TRANSFER_CHARGE = NVLD(Rs("STD_TRANSFER_CHARGE"), 0)
      
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")
      
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      DRCR_REASON_NO = NVLS(Rs("DRCR_REASON_NO"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
   ElseIf Ind = 107 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
   ElseIf Ind = 108 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   ElseIf Ind = 115 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
   ElseIf Ind = 116 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 117 Then
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 120 Then
      RECEIPT_NO = NVLS(Rs("RECEIPT_NO"), "")
      RECEIPT_DATE = InternalDateToDate(NVLS(Rs("RECEIPT_DATE"), ""))
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      BILL_DATE = InternalDateToDate(NVLS(Rs("BILL_DATE"), ""))
      
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      SALE_LNAME = NVLS(Rs("SALE_LNAME"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      BEFORE_DISCOUNT_PRICE = TOTAL_PRICE + DISCOUNT_AMOUNT
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      
      Credit = NVLI(Rs("CREDIT"), -1)
   ElseIf Ind = 121 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      RECEIPT_NO = NVLS(Rs("RECEIPT_NO"), "")
      RECEIPT_DATE = InternalDateToDate(NVLS(Rs("RECEIPT_DATE"), ""))
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      BILL_DATE = InternalDateToDate(NVLS(Rs("BILL_DATE"), ""))
      
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      FREELANCE_ID = NVLI(Rs("FREELANCE_ID"), -1)
      FREELANCE_CODE = NVLS(Rs("FREELANCE_CODE"), "")
      FREELANCE_NAME = NVLS(Rs("FREELANCE_NAME"), "")
      FREELANCE_LASTNAME = NVLS(Rs("FREELANCE_LASTNAME"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      BEFORE_DISCOUNT_PRICE = TOTAL_PRICE + DISCOUNT_AMOUNT
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      
      Credit = NVLI(Rs("CREDIT"), -1)
   ElseIf Ind = 122 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
     ElseIf Ind = 123 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim Bd As CBillingDoc
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = DRCR_REASON
   If TempID1 <= 0 Then
      TempID1 = "NULL"
   End If
   
   If MANUAL_FLAG = "" Then
      MANUAL_FLAG = "Y"
   End If
   
   TempID2 = CHEQUE_DOC_ID
   If TempID2 <= 0 Then
        TempID2 = "NULL"
   End If
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("RECEIPT_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO RECEIPT_ITEM " & _
                     "( " & _
                     "   RECEIPT_ITEM_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   DO_ID, " & _
                     "   RECEIPT_ITEM_AMOUNT, " & _
                     "   DISCOUNT_AMOUNT, " & _
                     "   DEPOSIT_AMOUNT, " & _
                     "   PAID_AMOUNT, " & _
                     "   DEBIT_CREDIT_AMOUNT, " & _
                     "   DRCR_REASON, " & _
                     "   CASH_DISCOUNT, "
      SQL1 = SQL1 & _
                     "   PART_ITEM_ID, " & _
                     "   RETURN_AMOUNT, " & _
                     "   LOCATION_ID, " & _
                     "   PART_TYPE, " & _
                     "   RETURN_TOTAL_PRICE, " & _
                     "   RETURN_AVG_PRICE, " & _
                     "   AVG_WEIGHT, " & _
                     "   FEATURE_ID, " & _
                     "   RETURN_DISCOUNT_AMOUNT, " & _
                     "   CONFIG_CODE, "
      SQL1 = SQL1 & _
                     "   DISPLAY_ID, " & _
                     "   COUNTRY_CURRENCY1, " & _
                     "   COUNTRY_CURRENCY2, " & _
                     "   WEIGHT_PER_PACK, " & _
                     "   PACK_AMOUNT, " & _
                     "   PRICE_PER_PACK, " & _
                     "   DISCOUNT_PER_PACK, " & _
                     "   MANUAL_FLAG, " & _
                     "   MANUAL_CODE, " & _
                     "   MANUAL_NAME, "
      SQL1 = SQL1 & _
                     "   RATE_TYPE, " & _
                     "   TRANSFER_WAGE, " & _
                     "   STD_TRANSFER_CHARGE, " & _
                     "   ITEM_DESC, " & _
                     "   LINK_ID, "
      SQL1 = SQL1 & _
                     "   AVG_PRICE, " & _
                     "   CHEQUE_DOC_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( "
      SQL1 = SQL1 & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                         "" & DO_ID & "," & _
                         "" & RECEIPT_ITEM_AMOUNT & "," & _
                        "" & DISCOUNT_AMOUNT & "," & _
                        "" & DEPOSIT_AMOUNT & "," & _
                        "" & PAID_AMOUNT & "," & _
                        "" & DEBIT_CREDIT_AMOUNT & "," & _
                        TempID1 & "," & _
                        CASH_DISCOUNT & ","
      SQL1 = SQL1 & _
                        PART_ITEM_ID & "," & _
                        RETURN_AMOUNT & "," & _
                        LOCATION_ID & "," & _
                        PART_TYPE & "," & _
                        RETURN_TOTAL_PRICE & "," & _
                        RETURN_AVG_PRICE & "," & _
                        AVG_WEIGHT & "," & _
                        FEATURE_ID & "," & _
                        RETURN_DISCOUNT_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(CONFIG_CODE)) & "',"
      SQL1 = SQL1 & _
                        DISPLAY_ID & "," & _
                        COUNTRY_CURRENCY1 & "," & _
                        COUNTRY_CURRENCY2 & "," & _
                        WEIGHT_PER_PACK & "," & _
                        PACK_AMOUNT & "," & _
                        PRICE_PER_PACK & "," & _
                        DISCOUNT_PER_PACK & "," & _
                        "'" & ChangeQuote(Trim(MANUAL_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(MANUAL_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(MANUAL_NAME)) & "',"
      SQL1 = SQL1 & _
                        RATE_TYPE & "," & _
                        TRANSFER_WAGE & "," & _
                        STD_TRANSFER_CHARGE & "," & _
                        "'" & ChangeQuote(Trim(ITEM_DESC)) & "'," & _
                        LINK_ID & ","
      SQL1 = SQL1 & _
                        AVG_PRICE & "," & _
                          TempID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         RECEIPT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE RECEIPT_ITEM SET " & _
                     "   BILLING_DOC_ID = " & BILLING_DOC_ID & ", " & _
                     "   DO_ID = " & DO_ID & ", " & _
                     "   RECEIPT_ITEM_AMOUNT = " & RECEIPT_ITEM_AMOUNT & ", " & _
                     "   DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                     "   DEPOSIT_AMOUNT = " & DEPOSIT_AMOUNT & ", " & _
                     "   PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "   DEBIT_CREDIT_AMOUNT = " & DEBIT_CREDIT_AMOUNT & ", " & _
                     "   DRCR_REASON = " & TempID1 & ", " & _
                     "   CASH_DISCOUNT = " & CASH_DISCOUNT & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   RETURN_AMOUNT = " & RETURN_AMOUNT & ", " & _
                     "   LOCATION_ID = " & LOCATION_ID & ", " & _
                     "   PART_TYPE = " & PART_TYPE & ", " & _
                     "   RETURN_TOTAL_PRICE = " & RETURN_TOTAL_PRICE & ", " & _
                     "   RETURN_AVG_PRICE = " & RETURN_AVG_PRICE & ", " & _
                     "   AVG_WEIGHT = " & AVG_WEIGHT & ", " & _
                     "   FEATURE_ID = " & FEATURE_ID & ", " & _
                     "   RETURN_DISCOUNT_AMOUNT = " & RETURN_DISCOUNT_AMOUNT & ", " & _
                     "   CONFIG_CODE = '" & ChangeQuote(Trim(CONFIG_CODE)) & "', "
               SQL1 = SQL1 & _
                     "   DISPLAY_ID = " & DISPLAY_ID & ", " & _
                     "   COUNTRY_CURRENCY1 = " & COUNTRY_CURRENCY1 & ", " & _
                     "   COUNTRY_CURRENCY1 = " & COUNTRY_CURRENCY1 & ", " & _
                     "   WEIGHT_PER_PACK = " & WEIGHT_PER_PACK & ", " & _
                     "   PACK_AMOUNT = " & PACK_AMOUNT & ", " & _
                     "   PRICE_PER_PACK = " & PRICE_PER_PACK & ", " & _
                     "   DISCOUNT_PER_PACK = " & DISCOUNT_PER_PACK & ", " & _
                     "   MANUAL_FLAG = '" & ChangeQuote(Trim(MANUAL_FLAG)) & "', " & _
                     "   MANUAL_CODE = '" & ChangeQuote(Trim(MANUAL_CODE)) & "', " & _
                     "   MANUAL_NAME = '" & ChangeQuote(Trim(MANUAL_NAME)) & "', " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   TRANSFER_WAGE = " & TRANSFER_WAGE & ", " & _
                     "   STD_TRANSFER_CHARGE = " & STD_TRANSFER_CHARGE & ", " & _
                     "   ITEM_DESC = '" & ChangeQuote(Trim(ITEM_DESC)) & "', " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                      "   CHEQUE_DOC_ID = '" & TempID2 & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
      
      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = DO_ID
      Call Bd.UpdatePaidAmount
      Call Bd.UpdateCnDnAmount
      Set Bd = Nothing
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
Dim Bd As CBillingDoc

   SQL1 = "DELETE FROM RECEIPT_ITEM WHERE RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID
   m_Conn.Execute (SQL1)
   
   Set Bd = New CBillingDoc
   Bd.BILLING_DOC_ID = DO_ID
   Call Bd.UpdatePaidAmount
   Call Bd.UpdateCnDnAmount
   Set Bd = Nothing
End Function
Public Function CopyObjectFromDo(Ind As Long, TempObj As CDoItem)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM
    PART_ITEM_ID = TempObj.PART_ITEM_ID
    RETURN_AMOUNT = TempObj.ITEM_AMOUNT
    LOCATION_ID = TempObj.LOCATION_ID
    PART_TYPE = TempObj.PART_TYPE
    RETURN_TOTAL_PRICE = TempObj.TOTAL_PRICE
    AVG_PRICE = TempObj.AVG_PRICE
    AVG_WEIGHT = TempObj.AVG_WEIGHT
    FEATURE_ID = TempObj.FEATURE_ID
    RETURN_DISCOUNT_AMOUNT = TempObj.DISCOUNT_AMOUNT
    CONFIG_CODE = TempObj.CONFIG_CODE
    DISPLAY_ID = TempObj.DISPLAY_ID
    PART_ITEM_ID = TempObj.PART_ITEM_ID
    COUNTRY_CURRENCY2 = TempObj.COUNTRY_CURRENCY2
    WEIGHT_PER_PACK = TempObj.WEIGHT_PER_PACK
    PACK_AMOUNT = TempObj.PACK_AMOUNT
    PRICE_PER_PACK = TempObj.PRICE_PER_PACK
    DISCOUNT_PER_PACK = TempObj.DISCOUNT_PER_PACK
    MANUAL_FLAG = TempObj.MANUAL_FLAG
    MANUAL_CODE = TempObj.MANUAL_CODE
    MANUAL_NAME = TempObj.MANUAL_NAME
    RATE_TYPE = TempObj.RATE_TYPE
    TRANSFER_WAGE = TempObj.TRANSFER_WAGE
    STD_TRANSFER_CHARGE = TempObj.STD_TRANSFER_CHARGE
    ITEM_DESC = TempObj.ITEM_DESC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM
   
End Function
Public Function CopyObject(Ind As Long, TempObj As CReceiptItem)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM
   DO_ID = TempObj.DO_ID
   
    PART_ITEM_ID = TempObj.PART_ITEM_ID
    RETURN_AMOUNT = TempObj.RETURN_AMOUNT
    LOCATION_ID = TempObj.LOCATION_ID
    PART_TYPE = TempObj.PART_TYPE
    RETURN_TOTAL_PRICE = TempObj.RETURN_TOTAL_PRICE
    DEBIT_CREDIT_AMOUNT = TempObj.RETURN_TOTAL_PRICE
    AVG_PRICE = TempObj.AVG_PRICE
    AVG_WEIGHT = TempObj.AVG_WEIGHT
    FEATURE_ID = TempObj.FEATURE_ID
    RETURN_DISCOUNT_AMOUNT = TempObj.RETURN_DISCOUNT_AMOUNT
    CONFIG_CODE = TempObj.CONFIG_CODE
    DISPLAY_ID = TempObj.DISPLAY_ID
    PART_ITEM_ID = TempObj.PART_ITEM_ID
    COUNTRY_CURRENCY2 = TempObj.COUNTRY_CURRENCY2
    WEIGHT_PER_PACK = TempObj.WEIGHT_PER_PACK
    PACK_AMOUNT = TempObj.PACK_AMOUNT
    PRICE_PER_PACK = TempObj.PRICE_PER_PACK
    DISCOUNT_PER_PACK = TempObj.DISCOUNT_PER_PACK
    MANUAL_FLAG = TempObj.MANUAL_FLAG
    MANUAL_CODE = TempObj.MANUAL_CODE
    MANUAL_NAME = TempObj.MANUAL_NAME
    RATE_TYPE = TempObj.RATE_TYPE
    TRANSFER_WAGE = TempObj.TRANSFER_WAGE
    STD_TRANSFER_CHARGE = TempObj.STD_TRANSFER_CHARGE
    ITEM_DESC = TempObj.ITEM_DESC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM
   
End Function

Public Function ShowDescText() As String
   If MANUAL_FLAG = "N" Then
      If DISPLAY_ID = 1 Then
         ShowDescText = ITEM_DESC
      ElseIf DISPLAY_ID = 2 Then
         ShowDescText = FEATURE_DESC
      ElseIf DISPLAY_ID = 3 Then
         ShowDescText = PART_DESC
      End If
   Else
      ShowDescText = MANUAL_NAME
   End If
End Function

Public Function ShowCodeText() As String
   If MANUAL_FLAG = "N" Then
      If DISPLAY_ID = 1 Then
         ShowCodeText = ""
      ElseIf DISPLAY_ID = 2 Then
         ShowCodeText = FEATURE_CODE
      ElseIf DISPLAY_ID = 3 Then
         ShowCodeText = PART_NO
      End If
   Else
      ShowCodeText = MANUAL_CODE
   End If
End Function
Public Function CopyObjectFromSup(Ind As Long, TempObj As CSupItem)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR Supitem
    PART_ITEM_ID = TempObj.PART_ITEM_ID
    RETURN_AMOUNT = TempObj.TX_AMOUNT
    LOCATION_ID = TempObj.LOCATION_ID
    PART_TYPE = TempObj.PART_TYPE
    RETURN_TOTAL_PRICE = TempObj.TOTAL_INCLUDE_PRICE
    AVG_PRICE = TempObj.INCLUDE_UNIT_PRICE
    'AVG_WEIGHT = TempObjAVG_WEIGHT
    'FEATURE_ID = TempObj.FEATURE_ID
    RETURN_DISCOUNT_AMOUNT = TempObj.DISCOUNT_AMOUNT
    'CONFIG_CODE = TempObj.CONFIG_CODE
    'DISPLAY_ID = TempObj.DISPLAY_ID
    PART_ITEM_ID = TempObj.PART_ITEM_ID
    'COUNTRY_CURRENCY2 = TempObj.COUNTRY_CURRENCY2
    'WEIGHT_PER_PACK = TempObj.WEIGHT_PER_PACK
    'PACK_AMOUNT = TempObj.PACK_AMOUNT
    'PRICE_PER_PACK = TempObj.PRICE_PER_PACK
    'DISCOUNT_PER_PACK = TempObj.DISCOUNT_PER_PACK
    'MANUAL_FLAG = TempObj.MANUAL_FLAG
    'MANUAL_CODE = TempObj.MANUAL_CODE
    MANUAL_NAME = TempObj.PART_DESC
'    RATE_TYPE = TempObj.RATE_TYPE
'    TRANSFER_WAGE = TempObj.TRANSFER_WAGE
'    STD_TRANSFER_CHARGE = TempObj.STD_TRANSFER_CHARGE
    ITEM_DESC = TempObj.ITEM_DESC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FOR DOITEM
   
End Function
Public Sub GenerateBDHeader(FileID As Long)
Dim TempStr As String
   
   TempStr = "BD;"
   TempStr = TempStr & DOCUMENT_NO & ";"
   TempStr = TempStr & DateToStringInt(DOCUMENT_DATE) & ";"
   TempStr = TempStr & DOCUMENT_TYPE & ";"
   TempStr = TempStr & DateToStringInt(DUE_DATE) & ";"
   TempStr = TempStr & NOTE & ";"
   
   TempStr = TempStr & BILLING_ADDRESS_ID & ";"
   TempStr = TempStr & ENTERPRISE_ID & ";"
   TempStr = TempStr & ENTERPRISE_ADDRESS_ID & ";"
   TempStr = TempStr & TOTAL_AMOUNT & ";"
   TempStr = TempStr & VAT_PERCENT & ";"
   
   TempStr = TempStr & VAT_AMOUNT & ";"
   TempStr = TempStr & WH_PERCENT & ";"
   TempStr = TempStr & WH_AMOUNT & ";"
   TempStr = TempStr & DISCOUNT_AMOUNT & ";"
   TempStr = TempStr & PAYMENT_TERM & ";"
   
   TempStr = TempStr & ACCEPT_BY & ";"
   TempStr = TempStr & RECEIVE_BY & ";"
   TempStr = TempStr & EXCEPTION_FLAG & ";"
   TempStr = TempStr & PAYEE_NAME & ";"
   TempStr = TempStr & COMMIT_FLAG & ";"
   
   'TempStr = TempStr & SUPPLIER_ID & ";"
   TempStr = TempStr & SUPPLIER_CODE & ";"
   TempStr = TempStr & RECEIPT_TYPE & ";"
   TempStr = TempStr & ACCOUNT_ID & ";"
   TempStr = TempStr & DEPOSIT_AMOUNT & ";"
   TempStr = TempStr & TOTAL_PRICE & ";"
   
   TempStr = TempStr & DateToStringInt(APPROVE_DATE) & ";"
   TempStr = TempStr & DateToStringInt(ESTIMATE_DATE) & ";"
   
   TempStr = TempStr & RESOURCE_ID & ";"
   TempStr = TempStr & PAYMENT_TYPE & ";"
   TempStr = TempStr & BANK_ID & ";"
   TempStr = TempStr & BBRANCH_ID & ";"
   TempStr = TempStr & CHECK_NO & ";"
   
   TempStr = TempStr & DateToStringInt(CHECK_DATE) & ";"
   TempStr = TempStr & VADILITY & ";"
   TempStr = TempStr & DELIVERY & ";"
   TempStr = TempStr & PAYMENT & ";"
   TempStr = TempStr & PAYMENT_DESC & ";"
   
   TempStr = TempStr & SHIPMENT & ";"
   TempStr = TempStr & REFER_INV & ";"
   TempStr = TempStr & PACKING_OF & ";"
   TempStr = TempStr & DateToStringInt(SHIPON_DATE) & ";"
   TempStr = TempStr & REF & ";"
   
   TempStr = TempStr & Credit & ";"
   TempStr = TempStr & AREA_CODE & ";"
   TempStr = TempStr & SHIP_NAME & ";"
   TempStr = TempStr & SHIPPING_MARKS & ";"
   TempStr = TempStr & CD_PERCENT & ";"
   
   TempStr = TempStr & CD_AMOUNT & ";"
   TempStr = TempStr & PACKAGE_TYPE & ";"
   TempStr = TempStr & TEMP_DO_NO & ";"
   TempStr = TempStr & PAID_AMOUNT & ";"
   TempStr = TempStr & DEBIT_AMOUNT & ";"
      
   TempStr = TempStr & CREDIT_AMOUNT & ";"
   TempStr = TempStr & DO_TOTAL_PRICE & ";"
   TempStr = TempStr & REVENUE_TOTAL_PRICE & ";"
   TempStr = TempStr & PAYMENT_ID & ";"
   TempStr = TempStr & BANK_BRANCH_ID & ";"
   
   TempStr = TempStr & BANK_NOTE & ";"
   TempStr = TempStr & TOTAL_RCP & ";"
   TempStr = TempStr & RUNNING_NO & ";"
   TempStr = TempStr & AGREEMENT_DATA & ";"
   TempStr = TempStr & AGREEMENT_FINANCE & ";"
   
   TempStr = TempStr & OLD_CREDIT_AMOUNT & ";"
   TempStr = TempStr & DateToStringInt(ENTRY_DATE) & ";"
   TempStr = TempStr & DateToStringInt(EXIT_DATE) & ";"
      
   Print #FileID, TempStr
End Sub
Public Sub GenerateRcpTailer(FileID As Long)
Dim TempStr As String
   
   TempStr = "RCP;"
   
   TempStr = TempStr & DRCR_REASON_NO & ";"
   TempStr = TempStr & LOCATION_NO & ";"
   TempStr = TempStr & PART_NO & ";"
   
   TempStr = TempStr & RECEIPT_ITEM_AMOUNT & ";"
   TempStr = TempStr & DISCOUNT_AMOUNT & ";"
   TempStr = TempStr & DEPOSIT_AMOUNT & ";"
   TempStr = TempStr & PAID_AMOUNT & ";"
   TempStr = TempStr & DEBIT_CREDIT_AMOUNT & ";"
   
   TempStr = TempStr & CASH_DISCOUNT & ";"
   TempStr = TempStr & LINK_ID & ";"
   TempStr = TempStr & RETURN_AMOUNT & ";"
   TempStr = TempStr & RETURN_TOTAL_PRICE & ";"
   
   TempStr = TempStr & RETURN_AVG_PRICE & ";"
   TempStr = TempStr & AVG_PRICE & ";"
   TempStr = TempStr & AVG_WEIGHT & ";"
   TempStr = TempStr & RETURN_DISCOUNT_AMOUNT & ";"
   TempStr = TempStr & CONFIG_CODE & ";"
   
   TempStr = TempStr & DISPLAY_ID & ";"
   TempStr = TempStr & COUNTRY_CURRENCY1 & ";"
   TempStr = TempStr & COUNTRY_CURRENCY2 & ";"
   TempStr = TempStr & WEIGHT_PER_PACK & ";"
   TempStr = TempStr & PACK_AMOUNT & ";"
   
   TempStr = TempStr & PRICE_PER_PACK & ";"
   TempStr = TempStr & DISCOUNT_PER_PACK & ";"
   TempStr = TempStr & MANUAL_FLAG & ";"
   TempStr = TempStr & MANUAL_CODE & ";"
   TempStr = TempStr & MANUAL_NAME & ";"
   
   TempStr = TempStr & RATE_TYPE & ";"
   TempStr = TempStr & TRANSFER_WAGE & ";"
   TempStr = TempStr & STD_TRANSFER_CHARGE & ";"
   TempStr = TempStr & ITEM_DESC & ";"
   
   TempStr = TempStr & BBRANCH_ID & ";"
   TempStr = TempStr & BILL_NO & ";"
      
   Print #FileID, TempStr
End Sub
Public Function DeleteFromBillInv()
Dim SQL1 As String

      SQL1 = "DELETE FROM RECEIPT_ITEM WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
End Function
Public Function DeleteDataFromDoID()
Dim SQL1 As String
   SQL1 = "DELETE FROM RECEIPT_ITEM WHERE DO_ID = " & DO_ID
   m_Conn.Execute (SQL1)
End Function
