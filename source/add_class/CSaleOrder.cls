VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSaleOrder"
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

Public SALE_ORDER_ID As Long
Public DO_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public ITEM_AMOUNT As Double
Public TOTAL_WEIGHT As Double
Public AVG_WEIGHT As Double
Public TOTAL_PRICE As Double
Public AVG_PRICE As Double
Public GUI_ID As Long
Public LINK_ID As Long
Public PIG_STATUS As Long
Public FEATURE_ID As Long
Public SOC_ID As Long
Public DISCOUNT_AMOUNT As Double
Public DEPOSIT_AMOUNT As Double
Public CONFIG_CODE As String
Public ITEM_DESC As String
Public DISPLAY_ID As Long
Public FROM_PERIOD As Date
Public TO_PERIOD As Date
Public WEIGHT_PER_PACK As Double
Public WEIGHT_PER_PACK_SO As Double
Public PACK_AMOUNT As Double
Public PRICE_PER_PACK As Double
Public DISCOUNT_PER_PACK As Double
Public MANUAL_FLAG As String
Public MANUAL_CODE As String
Public MANUAL_NAME As String
Public EXTRA_DISCOUNT As Double
Public RATE_TYPE As Long
Public PARCEL_TYPE As Long
Public TRANSFER_WAGE As Double
Public STD_TRANSFER_CHARGE As Double
 
Public EMP_CODE As String
Public FEATURE_CODE As String
Public FEATURE_DESC As String
Public SOC_CODE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public COMMIT_FLAG As String
Public PIG_TYPE As String
Public PART_NO As String
Public PART_DESC As String
Public PIG_STATUS_NAME As String
Public LOCATION_NAME As String
Public PART_TYPE_NAME As String
Public PART_TYPE As Long
Public PART_TYPE_OVER As String
Public PART_TYPE_NO As Long
Public LOT_MANUAL As String
Public LOT_ITEM_ID As Long
Public INVENTORY_DOC_ID As Long
Public INVENTORY_WH_DOC_ID As Long
Public CUSTOMER_CODE As String

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Public CUSTOMER_NAME As String
Public CUSTOMER_ID As Long
Public SERVICE_FLAG As String

Public UC_AMOUNT As Double
Public AC_AMOUNT As Double
Public MINIMUM_FLAG As String
Public CALCULATED_AMOUNT As Double
Public SUBSCRIBER_ID As Long
Public ACCOUNT_ID As Long
Public USAGE_DATE As Date
Public PO_NO As String
Public PO_ID As Long
Public BILLING_DOC_ID As Long
Public DOCUMENT_NO As String
Public BILLING_DOC_SO_ID As Long
Public BILLING_DOC_SO_NO As String
Public DOLLAR As Double
Public COUNTRY_CURRENCY1 As Long
Public COUNTRY_CURRENCY2 As Long
Public SubLotItems As Collection
Public UNIT_NAME As String
Public QTYP As Double
Public PACK As Double
Public DISCOUNT_PERCENT As Double
Public CUSTOMER_TYPE As Long
Public BEFORE_DISCOUNT_PRICE As Double
Public CURRENT_POSITION As Long
Public ACCEPT_BY As Long
Public LOGISTIC_FLAG As String
Public DOCUMENT_DATE As Date
Public TRUCK_NO As String
Public SUPPLIER_ID As Long
Public PACKAGE_TYPE As PACKAGE_TYPE
Public DOCUMENT_TYPE As Long
Public DOCUMENT_SUBTYPE As Long
Public RECEIPT_TYPE As Long
Public EMP_NAME As String
Public DocTypeSet As String

Public PRO_COMMISSION_BAG  As Double
Public PRO_CHEER_BAG   As Double
Public PRO_DST_BAG   As Double
Public PRO_OTHER1_BAG   As Double
Public PRO_OTHER2_BAG   As Double
Public PRO_OTHER3_BAG   As Double
Public PRO_COMMISSION_KG   As Double
Public PRO_CHEER_KG   As Double
Public PRO_DST_KG   As Double
Public PRO_OTHER1_KG   As Double
Public PRO_OTHER2_KG   As Double
Public PRO_OTHER3_KG   As Double

Public EX_PROMOTION_PART_ITEM_ID As Long
Public EX_PROMOTION_DLC_ITEM_ID As Long
Public DISCOUNT_AMOUNT_PART As Double
Public DISCOUNT_AMOUNT_DLC As Double

Public EX_WORKS_PRICE_ITEM_ID As Long
Public EX_DELIVERY_COST_ITEM_ID As Long
Public PACKAGE_RATE As Double
Public RATE_CUSTOMER As Double
Public USER_APPLOVE_PRICE As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   Set SubLotItems = New Collection
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
      SelectStr1 = "SELECT UG.*, PT.*, LC.*, PI.*, UN.*, PS.*, FT.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE, IVD.INVENTORY_DOC_ID,UN.UNIT_NAME, UG.WEIGHT_PER_PACK DO_WEIGHT_PER_PACK "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, UG.ITEM_DESC, UG.FEATURE_ID, UG.PART_ITEM_ID, SUM(UG.ITEM_AMOUNT) ITEM_AMOUNT, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = "WHERE (CUSTOMER_ID IS NOT NULL) AND (DOCUMENT_TYPE IN (1, 2, 5)) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, UG.FEATURE_ID, FT.FEATURE_CODE, FT.FEATURE_DESC, SUM(UG.ITEM_AMOUNT) ITEM_AMOUNT, SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(UG.TOTAL_PRICE) TOTAL_PRICE, SUM(UG.PACK_AMOUNT) PACK_AMOUNT, SUM(UG.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(UG.EXTRA_DISCOUNT) EXTRA_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      
      WhereStr = "WHERE (CUSTOMER_ID IS NOT NULL) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT IVD.ACCEPT_BY, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, UG.FEATURE_ID, SUM(UG.ITEM_AMOUNT) ITEM_AMOUNT, SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(UG.TOTAL_PRICE) TOTAL_PRICE, SUM(UG.PACK_AMOUNT) PACK_AMOUNT, SUM(UG.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(UG.EXTRA_DISCOUNT) EXTRA_DISCOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (IVD.ACCEPT_BY = EM.EMP_ID) "

      WhereStr = "WHERE (CUSTOMER_ID IS NOT NULL) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT IVD.*, UG.*, CM.CUSTOMER_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, FT.FEATURE_CODE, FT.FEATURE_DESC "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE_TYPE FTT ON (FT.FEATURE_TYPE = FTT.FEATURE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (CUSTOMER_ID IS NOT NULL) "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT IVD.PACKAGE_TYPE, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, UG.FEATURE_ID, FT.FEATURE_CODE, FT.FEATURE_DESC, SUM(UG.ITEM_AMOUNT) ITEM_AMOUNT, SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(UG.TOTAL_PRICE) TOTAL_PRICE, SUM(UG.PACK_AMOUNT) PACK_AMOUNT, SUM(UG.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(UG.EXTRA_DISCOUNT) EXTRA_DISCOUNT "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE_TYPE FTT ON (FT.FEATURE_TYPE = FTT.FEATURE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (CUSTOMER_ID IS NOT NULL) "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME "
      SelectStr1 = SelectStr1 & ", IVD.*, UG.*, FT.FEATURE_CODE, FT.FEATURE_DESC, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE,ENM.LONG_NAME EMP_LONG_NAME,ENM.LAST_NAME EMP_LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE_TYPE FTT ON (FT.FEATURE_TYPE = FTT.FEATURE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EM.EMP_ID = IVD.ACCEPT_BY) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      
      WhereStr = "WHERE (CM.CUSTOMER_ID IS NOT NULL) "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, SUM(DI.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER DI "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.DO_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT BD.ACCOUNT_ID, BD.DOCUMENT_TYPE, BD.DOCUMENT_SUBTYPE, BD.RECEIPT_TYPE, SUM(DI.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER DI "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.DO_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT DI.DO_ID, SUM(DI.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SALE_ORDER DI "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (DI.DO_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT DISTINCT UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (UG.PART_ITEM_ID > 0) "
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT IVD.ACCEPT_BY, PI.PART_ITEM_ID, SUM(UG.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(UG.EXTRA_DISCOUNT) EXTRA_DISCOUNT, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = ""
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, PI.PART_ITEM_ID, SUM(UG.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(UG.EXTRA_DISCOUNT) EXTRA_DISCOUNT, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT DISTINCT ADR.INVENTORY_DOC_ID,ADR.DOCUMENT_NO,ADR.DOCUMENT_DATE,CM.CUSTOMER_ID,NM.LONG_NAME CUSTOMER_NAME,BD.BILLING_DOC_ID BILLING_DOC_SO_ID,BD.DOCUMENT_NO BILLING_DOC_SO_NO,UG.WEIGHT_PER_PACK WEIGHT_PER_PACK_SO,UG.*,PT.*, LC.*, PI.*, UN.*, PS.*, FT.* "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SALE_ORDER UG "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (UG.DO_ID =LIW.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (LIW.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC ADR ON (LIW.INVENTORY_WH_DOC_ID =ADR.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (ADR.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (UG.FEATURE_ID = FT.FEATURE_ID) "
      


      WhereStr = ""
   End If

   If SALE_ORDER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SALE_ORDER_ID = " & SALE_ORDER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SALE_ORDER_ID = " & SALE_ORDER_ID & ") "
      End If
   End If
   
   If DO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DO_ID = " & DO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DO_ID = " & DO_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
      
   If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
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
   
   If Trim(EMP_CODE) <> "" Then
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         End If
      End If
   End If
         
   If Trim(TRUCK_NO) <> "" Then
      If WildCard(TRUCK_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(IVD.NOTE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(IVD.NOTE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (IVD.NOTE = '" & ChangeQuote(Trim(TRUCK_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (IVD.NOTE = '" & ChangeQuote(Trim(TRUCK_NO)) & "')"
         End If
      End If
   End If
         
   If Trim(PART_NO) <> "" Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
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
         
   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE IN " & DocTypeSet & ") "
      End If
   End If
   
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((PART_TYPE = " & PART_TYPE & ") ) "
      Else
         WhereStr = WhereStr & "AND ((PART_TYPE = " & PART_TYPE & ")) " 'OR (PART_TYPE IS NULL)
      End If
   End If
   
   If Len(PART_TYPE_OVER) > 0 Then ' กรณีพิเศษ ใช้สำหรับ การดึงใบขึ้นอาหารจาก Sale order ที่เป็นการขายอื่น ที่มีทั้งสินค้าและบริการ
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((PART_TYPE NOT IN " & PART_TYPE_OVER & ") "
         WhereStr = WhereStr & "OR (UG.FEATURE_ID IS NOT NULL)) "
      Else
         WhereStr = WhereStr & "AND ((PART_TYPE NOT IN " & PART_TYPE_OVER & ") "
         WhereStr = WhereStr & "OR (UG.FEATURE_ID IS NOT NULL)) "
      End If
   End If
   
   If CURRENT_POSITION > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CURRENT_POSITION = " & CURRENT_POSITION & ") "
      Else
         WhereStr = WhereStr & "AND (CURRENT_POSITION = " & CURRENT_POSITION & ") "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If INVENTORY_WH_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      End If
   End If
   
   
   If Len(LOGISTIC_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOGISTIC_FLAG = '" & LOGISTIC_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (LOGISTIC_FLAG = '" & LOGISTIC_FLAG & "') "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY CM.CUSTOMER_ID, UG.ITEM_DESC, UG.FEATURE_ID, UG.PART_ITEM_ID "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY CM.CUSTOMER_ID, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, UG.FEATURE_ID, FT.FEATURE_CODE, FT.FEATURE_DESC "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY IVD.ACCEPT_BY, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, UG.FEATURE_ID "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY IVD.PACKAGE_TYPE, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PARCEL_TYPE, UG.FEATURE_ID, FT.FEATURE_CODE, FT.FEATURE_DESC "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY CM.CUSTOMER_ID "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY BD.ACCOUNT_ID, BD.DOCUMENT_TYPE, BD.DOCUMENT_SUBTYPE, BD.RECEIPT_TYPE "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY DI.DO_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY IVD.ACCEPT_BY, PI.PART_ITEM_ID "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY CM.CUSTOMER_ID, PI.PART_ITEM_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SALE_ORDER_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SALE_ORDER_ID "
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
         OrderByStr = "ORDER BY SALE_ORDER_ID "
      End If
   ElseIf Ind = 2 Then
      If Ind = 1 Then
         OrderByStr = "ORDER BY CUSTOMER_CODE "
      ElseIf Ind = 2 Then
         OrderByStr = "ORDER BY LONG_NAME "
      Else
         OrderByStr = "ORDER BY CUSTOMER_CODE "
      End If
   ElseIf Ind = 3 Then
      If Ind = 1 Then
         OrderByStr = "ORDER BY CUSTOMER_CODE "
      ElseIf Ind = 2 Then
         OrderByStr = "ORDER BY PART_NO "
      Else
         OrderByStr = "ORDER BY CUSTOMER_CODE "
      End If
   ElseIf Ind = 4 Then
      If Ind = 1 Then
         OrderByStr = "ORDER BY EMP_CODE "
      ElseIf Ind = 2 Then
         OrderByStr = "ORDER BY PART_NO "
      Else
         OrderByStr = "ORDER BY EMP_CODE "
      End If
   ElseIf Ind = 5 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.NOTE ASC, IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO ASC, PART_NO ASC, FEATURE_CODE "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY IVD.NOTE ASC, IVD.DOCUMENT_NO  "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, IVD.DOCUMENT_NO "
      End If
   ElseIf Ind = 6 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_NO ASC, FEATURE_CODE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PACKAGE_TYPE ASC, PART_NO ASC, FEATURE_CODE "
      End If
   ElseIf Ind = 7 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, PI.PART_NO ASC, FT.FEATURE_CODE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY NM.LONG_NAME ASC, PI.PART_NO ASC, FT.FEATURE_CODE  "
      Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, PI.PART_NO ASC, FT.FEATURE_CODE  "
      End If
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY SALE_ORDER_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY SALE_ORDER_ID "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY DI.DO_ID "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY PI.PART_ITEM_ID "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY PI.PART_ITEM_ID "
   ElseIf Ind = 14 Then
      OrderByStr = "ORDER BY PI.PART_ITEM_ID "
   End If

   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function CopyObject(Ind As Long, TempObj As CDoItem)
   
   DO_ID = TempObj.DO_ID                                                             '2
   PART_ITEM_ID = TempObj.PART_ITEM_ID                          '3
   LOCATION_ID = TempObj.LOCATION_ID                          '4
   ITEM_AMOUNT = TempObj.ITEM_AMOUNT                          '5
   TOTAL_WEIGHT = TempObj.TOTAL_WEIGHT                          '6
   AVG_WEIGHT = TempObj.AVG_WEIGHT                          '7
   TOTAL_PRICE = TempObj.TOTAL_PRICE                          '8
   AVG_PRICE = TempObj.AVG_PRICE                          '9
   
   GUI_ID = TempObj.GUI_ID                          '10
   PIG_STATUS = TempObj.PIG_STATUS                          '11
   LINK_ID = TempObj.LINK_ID                          '12
   DEPOSIT_AMOUNT = TempObj.DEPOSIT_AMOUNT                          '13
   DISCOUNT_AMOUNT = TempObj.DISCOUNT_AMOUNT                          '14
   CONFIG_CODE = TempObj.CONFIG_CODE                          '15
   ITEM_DESC = TempObj.ITEM_DESC                          '16
   FEATURE_ID = TempObj.FEATURE_ID                          '17
   SOC_ID = TempObj.SOC_ID                          '18
   DISPLAY_ID = TempObj.DISPLAY_ID                          '19
   
   FEATURE_CODE = TempObj.FEATURE_CODE                          '20
   FEATURE_DESC = TempObj.FEATURE_DESC                          '21
   COMMIT_FLAG = TempObj.COMMIT_FLAG                          '22
   PART_TYPE = TempObj.PART_TYPE                          '23
   PART_NO = TempObj.PART_NO                          '24
   PART_DESC = TempObj.PART_DESC                          '25
   PART_TYPE_NAME = TempObj.PART_TYPE_NAME                          '26
   LOCATION_NAME = TempObj.LOCATION_NAME                          '27
   INVENTORY_DOC_ID = TempObj.INVENTORY_DOC_ID                          '28
   PO_NO = TempObj.PO_NO                          '29
   PO_ID = TempObj.PO_ID                          '30
   DOLLAR = TempObj.DOLLAR                  '31
      COUNTRY_CURRENCY1 = TempObj.COUNTRY_CURRENCY1                 '32
      COUNTRY_CURRENCY2 = TempObj.COUNTRY_CURRENCY2                 '33
     UNIT_NAME = TempObj.UNIT_NAME                                                                  '34
     QTYP = TempObj.QTYP                                                                                                '35
    PACK = TempObj.PACK                                                                                               '36
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      SALE_ORDER_ID = NVLI(Rs("SALE_ORDER_ID"), -1)                          '1
      DO_ID = NVLI(Rs("DO_ID"), -1)                          '2
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)                          '3
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)                          '4
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)                          '5
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)                          '6
      AVG_WEIGHT = NVLD(Rs("AVG_WEIGHT"), 0)                          '7
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)                          '8
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)                          '9
      GUI_ID = NVLI(Rs("GUI_ID"), -1)                          '10
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)                          '11
      LINK_ID = NVLI(Rs("LINK_ID"), -1)                          '12
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)                          '13
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)                          '14
      CONFIG_CODE = NVLS(Rs("CONFIG_CODE"), "")                          '15
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")                          '16
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)                          '17
      SOC_ID = NVLI(Rs("SOC_ID"), -1)                          '18
      DISPLAY_ID = NVLI(Rs("DISPLAY_ID"), -1)                          '19
      FROM_PERIOD = InternalDateToDate(NVLS(Rs("FROM_PERIOD"), ""))                          '20
      TO_PERIOD = InternalDateToDate(NVLS(Rs("TO_PERIOD"), ""))                          '21
      WEIGHT_PER_PACK = NVLD(Rs("DO_WEIGHT_PER_PACK"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      PRICE_PER_PACK = NVLD(Rs("PRICE_PER_PACK"), 0)
      DISCOUNT_PER_PACK = NVLD(Rs("DISCOUNT_PER_PACK"), 0)
      MANUAL_FLAG = NVLS(Rs("MANUAL_FLAG"), "")
      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")
      RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      TRANSFER_WAGE = NVLD(Rs("TRANSFER_WAGE"), 0)
      STD_TRANSFER_CHARGE = NVLD(Rs("STD_TRANSFER_CHARGE"), 0)
      
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")                          '22
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")                          '23
      SERVICE_FLAG = NVLS(Rs("SERVICE_FLAG"), "")
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")                          '24
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)                          '25
      PART_NO = NVLS(Rs("PART_NO"), "")                          '26
      PART_DESC = NVLS(Rs("PART_DESC"), "")                          '27
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")                          '28
      PART_TYPE_NO = NVLI(Rs("PART_TYPE_NO"), -1)                         '25
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")                          '29
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)                          '30
      PO_NO = NVLS(Rs("PO_NO"), "")                          '31
      PO_ID = NVLI(Rs("PO_ID"), -1)                          '32
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")                          '33
      DOLLAR = NVLD(Rs("DOLLAR"), 0)                          '34
      COUNTRY_CURRENCY1 = NVLI(Rs("COUNTRY_CURRENCY1"), -1)                          '35
     COUNTRY_CURRENCY2 = NVLI(Rs("COUNTRY_CURRENCY2"), -1)                          '36
     UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")                          '37
      QTYP = NVLD(Rs("QTYP"), 0)                          '38
      PACK = NVLD(Rs("PACK"), 0)                          '39
      DISCOUNT_PERCENT = NVLD(Rs("DISCOUNT_PERCENT"), 0)                          '40
      PACK_AMOUNT = NVLI(Rs("PACK_AMOUNT"), -1)                          '41
      
      EX_WORKS_PRICE_ITEM_ID = NVLI(Rs("EX_WORKS_PRICE_ITEM_ID"), -1)
      PACKAGE_RATE = NVLD(Rs("PACKAGE_RATE"), 0)
      EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
      RATE_CUSTOMER = NVLD(Rs("RATE_CUSTOMER"), 0)
      PRO_COMMISSION_BAG = NVLD(Rs("PRO_COMMISSION_BAG"), 0)
      PRO_CHEER_BAG = NVLD(Rs("PRO_CHEER_BAG"), 0)
      PRO_DST_BAG = NVLD(Rs("PRO_DST_BAG"), 0)
      PRO_OTHER1_BAG = NVLD(Rs("PRO_OTHER1_BAG"), 0)
      PRO_OTHER2_BAG = NVLD(Rs("PRO_OTHER2_BAG"), 0)
      PRO_OTHER3_BAG = NVLD(Rs("PRO_OTHER3_BAG"), 0)
      PRO_COMMISSION_KG = NVLD(Rs("PRO_COMMISSION_KG"), 0)
      PRO_CHEER_KG = NVLD(Rs("PRO_CHEER_KG"), 0)
      PRO_DST_KG = NVLD(Rs("PRO_DST_KG"), 0)
      PRO_OTHER1_KG = NVLD(Rs("PRO_OTHER1_KG"), 0)
      PRO_OTHER2_KG = NVLD(Rs("PRO_OTHER2_KG"), 0)
      PRO_OTHER3_KG = NVLD(Rs("PRO_OTHER3_KG"), 0)
      USER_APPLOVE_PRICE = NVLS(Rs("USER_APPLOVE_PRICE"), "")
      
      EX_PROMOTION_PART_ITEM_ID = NVLI(Rs("EX_PROMOTION_PART_ITEM_ID"), -1)
      EX_PROMOTION_DLC_ITEM_ID = NVLI(Rs("EX_PROMOTION_DLC_ITEM_ID"), -1)
      DISCOUNT_AMOUNT_PART = NVLD(Rs("DISCOUNT_AMOUNT_PART"), 0)
      DISCOUNT_AMOUNT_DLC = NVLD(Rs("DISCOUNT_AMOUNT_DLC"), 0)
   ElseIf Ind = 2 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 3 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      BEFORE_DISCOUNT_PRICE = TOTAL_PRICE + DISCOUNT_AMOUNT
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), 1)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
'      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
'      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")
   ElseIf Ind = 4 Then
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      BEFORE_DISCOUNT_PRICE = TOTAL_PRICE + DISCOUNT_AMOUNT
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), 1)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
   ElseIf Ind = 5 Then
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      PRICE_PER_PACK = NVLD(Rs("PRICE_PER_PACK"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      TRANSFER_WAGE = NVLD(Rs("TRANSFER_WAGE"), 0)
      STD_TRANSFER_CHARGE = NVLD(Rs("STD_TRANSFER_CHARGE"), 0)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TRUCK_NO = NVLS(Rs("NOTE"), "")
      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")
   ElseIf Ind = 6 Then
      PACKAGE_TYPE = NVLI(Rs("PACKAGE_TYPE"), -1)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      BEFORE_DISCOUNT_PRICE = TOTAL_PRICE + DISCOUNT_AMOUNT
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
   ElseIf Ind = 7 Then
      SALE_ORDER_ID = NVLI(Rs("SALE_ORDER_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      AVG_WEIGHT = NVLD(Rs("AVG_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      CONFIG_CODE = NVLS(Rs("CONFIG_CODE"), "")
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      SOC_ID = NVLI(Rs("SOC_ID"), -1)
      DISPLAY_ID = NVLI(Rs("DISPLAY_ID"), -1)
      FROM_PERIOD = InternalDateToDate(NVLS(Rs("FROM_PERIOD"), ""))
      TO_PERIOD = InternalDateToDate(NVLS(Rs("TO_PERIOD"), ""))
      WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      PRICE_PER_PACK = NVLD(Rs("PRICE_PER_PACK"), 0)
      DISCOUNT_PER_PACK = NVLD(Rs("DISCOUNT_PER_PACK"), 0)
      MANUAL_FLAG = NVLS(Rs("MANUAL_FLAG"), "")
      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")

      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DISCOUNT_PERCENT = NVLD(Rs("DISCOUNT_PERCENT"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), 1)
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("EMP_LONG_NAME"), "") & " " & NVLS(Rs("EMP_LAST_NAME"), "")
      
   ElseIf Ind = 8 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
   ElseIf Ind = 9 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
   ElseIf Ind = 10 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
   ElseIf Ind = 11 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 12 Then
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 13 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 14 Then
      SALE_ORDER_ID = NVLI(Rs("SALE_ORDER_ID"), -1)                          '1
      DO_ID = NVLI(Rs("DO_ID"), -1)                          '2
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)                          '3
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)                          '4
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)                          '5
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)                          '6
      AVG_WEIGHT = NVLD(Rs("AVG_WEIGHT"), 0)                          '7
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)                          '8
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)                          '9
      GUI_ID = NVLI(Rs("GUI_ID"), -1)                          '10
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)                          '11
      LINK_ID = NVLI(Rs("LINK_ID"), -1)                          '12
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)                          '13
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)                          '14
      CONFIG_CODE = NVLS(Rs("CONFIG_CODE"), "")                          '15
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")                          '16
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)                          '17
      SOC_ID = NVLI(Rs("SOC_ID"), -1)                          '18
      DISPLAY_ID = NVLI(Rs("DISPLAY_ID"), -1)                          '19
      FROM_PERIOD = InternalDateToDate(NVLS(Rs("FROM_PERIOD"), ""))                          '20
      TO_PERIOD = InternalDateToDate(NVLS(Rs("TO_PERIOD"), ""))                          '21
      WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
      WEIGHT_PER_PACK_SO = NVLD(Rs("WEIGHT_PER_PACK_SO"), 0)
      PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
      PRICE_PER_PACK = NVLD(Rs("PRICE_PER_PACK"), 0)
      DISCOUNT_PER_PACK = NVLD(Rs("DISCOUNT_PER_PACK"), 0)
      MANUAL_FLAG = NVLS(Rs("MANUAL_FLAG"), "")
      MANUAL_CODE = NVLS(Rs("MANUAL_CODE"), "")
      MANUAL_NAME = NVLS(Rs("MANUAL_NAME"), "")
      RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
      EXTRA_DISCOUNT = NVLD(Rs("EXTRA_DISCOUNT"), 0)
      TRANSFER_WAGE = NVLD(Rs("TRANSFER_WAGE"), 0)
      STD_TRANSFER_CHARGE = NVLD(Rs("STD_TRANSFER_CHARGE"), 0)
      
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")                          '22
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")                          '23
      SERVICE_FLAG = NVLS(Rs("SERVICE_FLAG"), "")
'      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")                          '24
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)                          '25
      PART_NO = NVLS(Rs("PART_NO"), "")                          '26
      PART_DESC = NVLS(Rs("PART_DESC"), "")                          '27
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")                          '28
      PART_TYPE_NO = NVLI(Rs("PART_TYPE_NO"), -1)                         '25
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")                          '29
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)                          '30
      PO_NO = NVLS(Rs("PO_NO"), "")                          '31
      PO_ID = NVLI(Rs("PO_ID"), -1)                          '32
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")                          '33
      DOLLAR = NVLD(Rs("DOLLAR"), 0)                          '34
      COUNTRY_CURRENCY1 = NVLI(Rs("COUNTRY_CURRENCY1"), -1)                          '35
     COUNTRY_CURRENCY2 = NVLI(Rs("COUNTRY_CURRENCY2"), -1)                          '36
     UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")                          '37
      QTYP = NVLD(Rs("QTYP"), 0)                          '38
      PACK = NVLD(Rs("PACK"), 0)                          '39
      DISCOUNT_PERCENT = NVLD(Rs("DISCOUNT_PERCENT"), 0)                          '40
      PACK_AMOUNT = NVLI(Rs("PACK_AMOUNT"), -1)                          '41
      
      EX_WORKS_PRICE_ITEM_ID = NVLI(Rs("EX_WORKS_PRICE_ITEM_ID"), -1)
      PACKAGE_RATE = NVLD(Rs("PACKAGE_RATE"), 0)
      EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
      RATE_CUSTOMER = NVLD(Rs("RATE_CUSTOMER"), 0)
      PRO_COMMISSION_BAG = NVLD(Rs("PRO_COMMISSION_BAG"), 0)
      PRO_CHEER_BAG = NVLD(Rs("PRO_CHEER_BAG"), 0)
      PRO_DST_BAG = NVLD(Rs("PRO_DST_BAG"), 0)
      PRO_OTHER1_BAG = NVLD(Rs("PRO_OTHER1_BAG"), 0)
      PRO_OTHER2_BAG = NVLD(Rs("PRO_OTHER2_BAG"), 0)
      PRO_OTHER3_BAG = NVLD(Rs("PRO_OTHER3_BAG"), 0)
      PRO_COMMISSION_KG = NVLD(Rs("PRO_COMMISSION_KG"), 0)
      PRO_CHEER_KG = NVLD(Rs("PRO_CHEER_KG"), 0)
      PRO_DST_KG = NVLD(Rs("PRO_DST_KG"), 0)
      PRO_OTHER1_KG = NVLD(Rs("PRO_OTHER1_KG"), 0)
      PRO_OTHER2_KG = NVLD(Rs("PRO_OTHER2_KG"), 0)
      PRO_OTHER3_KG = NVLD(Rs("PRO_OTHER3_KG"), 0)
      USER_APPLOVE_PRICE = NVLS(Rs("USER_APPLOVE_PRICE"), "")
      BILLING_DOC_SO_ID = NVLI(Rs("BILLING_DOC_SO_ID"), -1)
      BILLING_DOC_SO_NO = NVLS(Rs("BILLING_DOC_SO_NO"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID7 As String
Dim TempID8 As String
Dim TempID9 As String
Dim TempID10 As String

   TempID1 = PIG_STATUS
   If PIG_STATUS <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = FEATURE_ID
   If FEATURE_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = SOC_ID
   If SOC_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PO_ID
   If PO_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID7 = EX_WORKS_PRICE_ITEM_ID
   If EX_WORKS_PRICE_ITEM_ID <= 0 Then
      TempID7 = "NULL"
   End If
   
   TempID8 = EX_DELIVERY_COST_ITEM_ID
   If EX_DELIVERY_COST_ITEM_ID <= 0 Then
      TempID8 = "NULL"
   End If
   
   TempID9 = EX_PROMOTION_PART_ITEM_ID
   If EX_PROMOTION_PART_ITEM_ID <= 0 Then
      TempID9 = "NULL"
   End If
   
   TempID10 = EX_PROMOTION_DLC_ITEM_ID
   If EX_PROMOTION_DLC_ITEM_ID <= 0 Then
      TempID10 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SALE_ORDER_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SALE_ORDER " & _
                     "( " & _
                     "   SALE_ORDER_ID, " & _
                     "   DO_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   ITEM_AMOUNT, " & _
                     "   TOTAL_WEIGHT, " & _
                     "   AVG_WEIGHT, " & _
                     "   TOTAL_PRICE, " & _
                     "   AVG_PRICE, "
            SQL1 = SQL1 & _
                     "   GUI_ID, " & _
                     "   LINK_ID, " & _
                     "   PIG_STATUS, " & _
                     "   DISCOUNT_AMOUNT, " & _
                     "   DEPOSIT_AMOUNT, " & _
                     "   CONFIG_CODE, " & _
                     "   ITEM_DESC, " & _
                     "   FEATURE_ID, " & _
                     "   SOC_ID, " & _
                     "   DISPLAY_ID, " & _
                     "   PO_NO, " & _
                     "   PO_ID, " & _
                     "   FROM_PERIOD, " & _
                     "   TO_PERIOD, " & _
                     "   DOLLAR, " & _
                     "   COUNTRY_CURRENCY1, " & _
                     "   COUNTRY_CURRENCY2, "
                    SQL1 = SQL1 & _
                     "   QTYP, " & _
                     "   PACK, " & _
                     "   DISCOUNT_PERCENT, " & _
                     "   WEIGHT_PER_PACK, " & _
                     "   PACK_AMOUNT, " & _
                     "   PRICE_PER_PACK, " & _
                     "   DISCOUNT_PER_PACK, " & _
                     "   MANUAL_FLAG, " & _
                     "   MANUAL_CODE, " & _
                     "   MANUAL_NAME, " & _
                     "   RATE_TYPE, " & _
                     "   EXTRA_DISCOUNT, " & _
                     "   TRANSFER_WAGE, " & _
                     "   STD_TRANSFER_CHARGE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY, "
               SQL1 = SQL1 & "   PRO_COMMISSION_BAG, " & _
                     "   PRO_CHEER_BAG, " & _
                     "   PRO_DST_BAG, " & _
                     "   PRO_OTHER1_BAG, " & _
                    "    PRO_OTHER2_BAG, " & _
                    "    PRO_OTHER3_BAG, " & _
                     "   PRO_COMMISSION_KG, " & _
                     "   PRO_CHEER_KG, " & _
                     "   PRO_DST_KG, " & _
                     "   PRO_OTHER1_KG, " & _
                    "    PRO_OTHER2_KG, " & _
                    "    PRO_OTHER3_KG, " & _
                    "    EX_WORKS_PRICE_ITEM_ID, " & _
                    "    EX_DELIVERY_COST_ITEM_ID, " & _
                    "    EX_PROMOTION_PART_ITEM_ID, " & _
                    "    EX_PROMOTION_DLC_ITEM_ID, " & _
                    "    PACKAGE_RATE, " & _
                    "    RATE_CUSTOMER, " & _
                    "    DISCOUNT_AMOUNT_PART, " & _
                    "    DISCOUNT_AMOUNT_DLC, " & _
                    "    USER_APPLOVE_PRICE" & _
                     ") " & _
                     "VALUES "
            SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "" & DO_ID & "," & _
                         "" & PART_ITEM_ID & "," & _
                         "" & LOCATION_ID & "," & _
                         "" & ITEM_AMOUNT & "," & _
                         "" & TOTAL_WEIGHT & "," & _
                         "" & AVG_WEIGHT & "," & _
                         "" & TOTAL_PRICE & "," & _
                         "" & AVG_PRICE & "," & _
                         "" & GUI_ID & "," & _
                         "" & LINK_ID & "," & _
                         TempID1 & "," & _
                         "" & DISCOUNT_AMOUNT & "," & _
                         "" & DEPOSIT_AMOUNT & "," & _
                         "'" & CONFIG_CODE & "'," & _
                         "'" & ITEM_DESC & "',"
            SQL1 = SQL1 & _
                         TempID2 & "," & _
                         TempID3 & "," & _
                         DISPLAY_ID & "," & _
                         "'" & PO_NO & "'," & _
                         TempID4 & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(FROM_PERIOD))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(TO_PERIOD))) & "'," & _
                         "'" & DOLLAR & "'," & _
                         "'" & COUNTRY_CURRENCY1 & "'," & _
                         "'" & COUNTRY_CURRENCY2 & "'," & _
                         "'" & QTYP & "'," & _
                         "'" & PACK & "'," & _
                         "'" & DISCOUNT_PERCENT & "',"
            SQL1 = SQL1 & _
                        WEIGHT_PER_PACK & "," & _
                        PACK_AMOUNT & "," & _
                        PRICE_PER_PACK & "," & _
                        DISCOUNT_PER_PACK & "," & _
                        "'" & ChangeQuote(Trim(MANUAL_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(MANUAL_CODE)) & "'," & _
                        "'" & ChangeQuote(Trim(MANUAL_NAME)) & "'," & _
                        RATE_TYPE & "," & _
                        EXTRA_DISCOUNT & "," & _
                        TRANSFER_WAGE & "," & _
                        STD_TRANSFER_CHARGE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                        glbUser.USER_ID & ","
                  SQL1 = SQL1 & PRO_COMMISSION_BAG & "," & _
                        PRO_CHEER_BAG & "," & _
                        PRO_DST_BAG & "," & _
                        PRO_OTHER1_BAG & "," & _
                        PRO_OTHER2_BAG & "," & _
                        PRO_OTHER3_BAG & "," & _
                        PRO_COMMISSION_KG & "," & _
                        PRO_CHEER_KG & "," & _
                        PRO_DST_KG & "," & _
                        PRO_OTHER1_KG & "," & _
                        PRO_OTHER2_KG & "," & _
                        PRO_OTHER3_KG & "," & _
                        TempID7 & "," & _
                        TempID8 & "," & _
                        TempID9 & "," & _
                        TempID10 & "," & _
                        PACKAGE_RATE & "," & _
                        RATE_CUSTOMER & "," & _
                        DISCOUNT_AMOUNT_PART & "," & _
                        DISCOUNT_AMOUNT_DLC & "," & _
                        "'" & USER_APPLOVE_PRICE & "'" & _
                     ") "
         SALE_ORDER_ID = ID
   Else 'Show edit
'      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "UPDATE SALE_ORDER SET " & _
                     "   DO_ID = " & DO_ID & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   LOCATION_ID = " & LOCATION_ID & ", " & _
                     "   ITEM_AMOUNT = " & ITEM_AMOUNT & ", " & _
                     "   TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   AVG_WEIGHT = " & AVG_WEIGHT & ", " & _
                     "   TOTAL_PRICE = " & TOTAL_PRICE & ", " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                     "   GUI_ID = " & GUI_ID & ", " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                     "   PIG_STATUS = " & TempID1 & ", " & _
                     "   DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                     "   DEPOSIT_AMOUNT = " & DEPOSIT_AMOUNT & ", " & _
                     "   CONFIG_CODE = '" & CONFIG_CODE & "', " & _
                     "   ITEM_DESC = '" & ITEM_DESC & "', " & _
                     "   FEATURE_ID = " & TempID2 & ", " & _
                     "   SOC_ID = " & TempID3 & ", " & _
                     "   DISPLAY_ID = " & DISPLAY_ID & ", " & _
                     "   PO_NO = '" & PO_NO & "', " & _
                     "   PO_ID = " & TempID4 & ", "
               SQL1 = SQL1 & _
                     "   FROM_PERIOD = '" & ChangeQuote(Trim(DateToStringInt(FROM_PERIOD))) & "', " & _
                     "   TO_PERIOD = '" & ChangeQuote(Trim(DateToStringInt(TO_PERIOD))) & "', " & _
                    "   DOLLAR = '" & DOLLAR & "', " & _
                    "   COUNTRY_CURRENCY1 = '" & COUNTRY_CURRENCY1 & "', " & _
                    "   COUNTRY_CURRENCY2 = '" & COUNTRY_CURRENCY2 & "', " & _
                    "   QTYP = '" & QTYP & "', " & _
                    "   PACK = '" & PACK & "', " & _
                    "   DISCOUNT_PERCENT = '" & DISCOUNT_PERCENT & "', " & _
                     "   WEIGHT_PER_PACK = " & WEIGHT_PER_PACK & ", " & _
                     "   PACK_AMOUNT = " & PACK_AMOUNT & ", " & _
                     "   PRICE_PER_PACK = " & PRICE_PER_PACK & ", " & _
                     "   DISCOUNT_PER_PACK = " & DISCOUNT_PER_PACK & ", " & _
                     "   MANUAL_FLAG = '" & ChangeQuote(Trim(MANUAL_FLAG)) & "', " & _
                     "   MANUAL_CODE = '" & ChangeQuote(Trim(MANUAL_CODE)) & "', " & _
                     "   MANUAL_NAME = '" & ChangeQuote(Trim(MANUAL_NAME)) & "', " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   EXTRA_DISCOUNT = " & EXTRA_DISCOUNT & ", " & _
                     "   TRANSFER_WAGE = " & TRANSFER_WAGE & ", " & _
                     "   STD_TRANSFER_CHARGE = " & STD_TRANSFER_CHARGE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & ","
             SQL1 = SQL1 & " PRO_COMMISSION_BAG = " & PRO_COMMISSION_BAG & ", " & _
                "   PRO_CHEER_BAG = " & PRO_CHEER_BAG & ", " & _
                "   PRO_DST_BAG = " & PRO_DST_BAG & ", " & _
                "   PRO_OTHER1_BAG = " & PRO_OTHER1_BAG & ", " & _
                "   PRO_OTHER2_BAG = " & PRO_OTHER2_BAG & ", " & _
                "   PRO_OTHER3_BAG = " & PRO_OTHER3_BAG & ", " & _
                "   PRO_COMMISSION_KG = " & PRO_COMMISSION_KG & ", " & _
                "   PRO_CHEER_KG = " & PRO_CHEER_KG & ", " & _
                "   PRO_DST_KG = " & PRO_DST_KG & ", " & _
                "   PRO_OTHER1_KG = " & PRO_OTHER1_KG & ", " & _
                "   PRO_OTHER2_KG = " & PRO_OTHER2_KG & ", " & _
                "   PRO_OTHER3_KG = " & PRO_OTHER3_KG & ", " & _
                "   EX_WORKS_PRICE_ITEM_ID = " & TempID7 & ", " & _
                "   EX_DELIVERY_COST_ITEM_ID = " & TempID8 & ", " & _
                "   EX_PROMOTION_PART_ITEM_ID = " & TempID9 & ", " & _
                "   EX_PROMOTION_DLC_ITEM_ID = " & TempID10 & ", " & _
                "   PACKAGE_RATE = " & PACKAGE_RATE & ", " & _
                "   RATE_CUSTOMER = " & RATE_CUSTOMER & ", " & _
                "   DISCOUNT_AMOUNT_PART = " & DISCOUNT_AMOUNT_PART & ", " & _
                "   DISCOUNT_AMOUNT_DLC = " & DISCOUNT_AMOUNT_DLC & ", " & _
                "   USER_APPLOVE_PRICE = '" & USER_APPLOVE_PRICE & "'" & _
                 " WHERE SALE_ORDER_ID = " & SALE_ORDER_ID '
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM SALE_ORDER WHERE SALE_ORDER_ID = " & SALE_ORDER_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set SubLotItems = Nothing
End Sub
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
