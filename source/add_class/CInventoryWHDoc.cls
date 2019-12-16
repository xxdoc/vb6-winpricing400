VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryWHDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DocumentType ต่างๆ ที่ใช้ในการรับเข้า InventoryWHDoc
'13= รับเข้า Bulk
'14= รับเข้า Bag
'15= ปรับยอด Bag
'16= ปรับยอด Bulk
'17= Rebag to Bag
'18= Rebag to Bulk
'19= Rebag to RM
'20= รับเข้าจากการโอนย้ายคลัง Bag
'21= รับเข้าจากการโอนย้ายคลัง Bulk


'DocumentType ต่างๆ ที่ใช้ในการจ่ายออก InventoryWHDoc
'2000= จ่ายออกอาหาร Bag
'2001= จ่ายออกอาหาร Bulk
'2002= โอน bag เพื่อ Rebag
'2003= โอน Bulk เพื่อ ผลิตใหม่ to Rebag
'2004= จ่ายออกสินค้าอื่นๆ

Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public INVENTORY_WH_DOC_ID As Long
Public INVENTORY_WH_TYPE As Long
Public DOCUMENT_NO As String
Public DOCUMENT_NO_SO As String
Public DOCUMENT_NO_RQ As String
Public DOCUMENT_DATE As Date
Public MAX_BALANCE_DATE As Date
Public PACK_DATE As Date
Public DOCUMENT_DESC As String
Public BILLING_DOC_ID As Long
Public BILL_NO As String
Public DO_NO As String
Public SO_NO As String
Public TRUCK_NO As String
Public CUSTOMER_ID As Long
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String
Public CUSTOMER_TYPE As Long
Public CAPACITY_AMOUNT As Double
Public TX_AMOUNT As Double
Public DELIVERY_ID As Long
Public DELIVERY_FEE As Double
Public SENDER_NAME As String
Public RECEIVE_NAME As String
Public DOCUMENT_TYPE As Long
Public DOCUMENT_TYPE_SET As String
Public PALLET_NO As Long
Public EMP_ID As Long
Public COMMIT_FLAG As String
Public SALE_FLAG As String
Public REASON_ID As Long
Public QUE_NO As String
Public ENTRY_DATE As Date
Public EXIT_DATE As Date
Public EXCEPTION_FLAG As String
Public DEPARTMENT_ID As Long
Public ADJUST_FLAG As String
Public Credit As Double
Public YYYYMM As String
Public CONDITION1 As Long
Public CONDITION2 As Long
Public CONDITION3 As Long
Public EMP_CHECK_CAR_ID As Long
Public EMP_CHECK_CAR_NAME As String
Public EMP_CHECK_CAR_LNAME As String
Public EMP_CHECK_PRODUCT_ID As Long
Public EMP_CHECK_PRODUCT_NAME As String
Public EMP_CHECK_PRODUCT_LNAME As String
Public RESPONSE_NAME As String
Public RESPONSE_LNAME As String

Public WEIGHT_ID As String
Public ENTRY_WEIGHT As Double
Public EXIT_WEIGHT As Double
Public TOTAL_WEIGHT As Double
Public WEIGHT_NOTE As String
Public ENTRY_WEIGHT_DATE As String
Public ENTRY_WEIGHT_TIME As String
Public EXIT_WEIGHT_DATE As String
Public EXIT_WEIGHT_TIME As String

Public NOTE As String
Public LOT_ID As Long
Public BATCH_NO As String
Public PROCESS_ID As Long
Public APPROVED_BY As Long
Public RESPONSE_BY As Long
Public START_DATE As Date
Public FINISH_DATE As Date
Public SUCCESS_FLAG As String
Public LOAD_FLAG As String
Public LOAD_GOODS_NO As String

Public VERIFY_FLAG As String
Public VERIFY_NAME As String

Public LOT_NO As String
Public BIN_ID As Long
Public BIN_NO As String
Public HEAD_PACK_NO As Long

Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_TYPE As Long
Public TIME_PACK_BEGIN As String

Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As CONFIG_DOC_TYPE

Public OLD_COMMIT_FLAG As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DEPARTMENT_NAME As String
Public CONSIGNMENT_FLAG As String
Public INVENTORY_DOC_ID As Long

Public MODIFY_DATE As Date
Public CREATE_NAME As String
Public MODIFY_NAME As String
Public CREATE_BY As Long
Public MODIFY_BY As Long

Public ImportItems As Collection
Public ExportItems As Collection
Public ImportExports As Collection
Public TransferItems As Collection
Public C_LotItemsWH As Collection


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   DOCUMENT_DATE = -1
   PACK_DATE = -1
   COMMIT_FLAG = "N"
   SALE_FLAG = "N"
   MAX_BALANCE_DATE = -1
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   Set ImportExports = New Collection
   Set TransferItems = New Collection
   Set C_LotItemsWH = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
   If Ind = 1 Then 'ใช้แล้ว
      SelectStr1 = "SELECT IWD.*,CM.CUSTOMER_CODE,NM1.LONG_NAME CUSTOMER_NAME,CN.CUSTOMER_ID, NM3.LONG_NAME EMP_CHECK_CAR_NAME, NM3.LAST_NAME EMP_CHECK_CAR_LNAME, NM4.LONG_NAME EMP_CHECK_PRODUCT_NAME, NM4.LAST_NAME EMP_CHECK_PRODUCT_LNAME,NM5.LONG_NAME RESPONSE_NAME,NM5.LAST_NAME RESPONSE_LNAME,IVD.INVENTORY_DOC_ID,IVD.DOCUMENT_NO DOCUMENT_NO_RQ,J.PROCESS_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (IWD.INVENTORY_WH_DOC_ID = J.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = IWD.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = IWD.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (CN.NAME_ID = NM1.NAME_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (IWD.DEPARTMENT_ID = LO.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = IWD.EMP_CHECK_CAR_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EN2.EMP_ID = IWD.EMP_CHECK_PRODUCT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN2.NAME_ID = NM4.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN3 ON (EN3.EMP_ID =IWD.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM5 ON (EN3.NAME_ID = NM5.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD  ON (IWD.INVENTORY_WH_DOC_ID = IVD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 2 Then '
      SelectStr1 = "SELECT IWD.INVENTORY_WH_DOC_ID,IWD.DOCUMENT_NO,IWD.DOCUMENT_DATE,IWD.DOCUMENT_DESC,IWD.TRUCK_NO, CM.CUSTOMER_CODE,NM1.LONG_NAME CUSTOMER_NAME,CN.CUSTOMER_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = IWD.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = IWD.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (CN.NAME_ID = NM1.NAME_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (IWD.DEPARTMENT_ID = LO.LAY_OUT_ID) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT ADR.*, NM1.LONG_NAME SUPPLIER_NAME, NM3.LONG_NAME RESPONSE_NAME, NM3.LAST_NAME RESPONSE_LNAME, SP.SUPPLIER_CODE, LO.LAY_OUT_NAME DEPARTMENT_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC ADR "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = ADR.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = ADR.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = ADR.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (ADR.DEPARTMENT_ID = LO.LAY_OUT_ID) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT(IWD.INVENTORY_WH_DOC_ID),IWD.DOCUMENT_NO,IWD.DOCUMENT_DATE,BD.DOCUMENT_NO DOCUMENT_NO_SO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (IWD.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (IWD.INVENTORY_WH_DOC_ID = LIW.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (LIW.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
    ElseIf Ind = 5 Then 'ใช้แล้ว
      SelectStr1 = "SELECT IWD.*,LIW.HEAD_PACK_NO  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = IWD.EMP_CHECK_CAR_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EN2.EMP_ID = IWD.EMP_CHECK_PRODUCT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN2.NAME_ID = NM4.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN3 ON (EN3.EMP_ID =IWD.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM5 ON (EN3.NAME_ID = NM5.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (IWD.INVENTORY_WH_DOC_ID =LIW.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB J ON (IWD.INVENTORY_WH_DOC_ID =J.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 6 Then  'ใช้แล้ว
      SelectStr1 = "SELECT IWD.*,CM.CUSTOMER_CODE,NM6.LONG_NAME CUSTOMER_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = IWD.EMP_CHECK_CAR_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EN2.EMP_ID = IWD.EMP_CHECK_PRODUCT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN2.NAME_ID = NM4.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN3 ON (EN3.EMP_ID =IWD.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM5 ON (EN3.NAME_ID = NM5.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (IWD.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM6 ON (CN.NAME_ID = NM6.NAME_ID) "
   ElseIf Ind = 7 Then  'ใช้แล้ว
      SelectStr1 = "SELECT DISTINCT IWD.INVENTORY_WH_DOC_ID,LIW.BILLING_DOC_ID,IWD.DOCUMENT_NO LOAD_GOODS_NO,IWD.DO_NO,IWD.TRUCK_NO,IWD.LOAD_FLAG,IWD.NOTE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (IWD.INVENTORY_WH_DOC_ID =LIW.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 8 Then  'ใช้แล้ว
      SelectStr1 = "SELECT MAX(IWD.DOCUMENT_DATE) MAX_BALANCE_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
   ElseIf Ind = 9 Then  'ใช้แล้ว
      SelectStr1 = "SELECT IWD.DOCUMENT_NO,IWD.DOCUMENT_DATE,IWD.TRUCK_NO,IWD.ENTRY_WEIGHT_DATE,IWD.ENTRY_WEIGHT_TIME,IWD.EXIT_WEIGHT_DATE, "
      SelectStr1 = SelectStr1 & "IWD.EXIT_WEIGHT_TIME,LIW.TX_AMOUNT,PI.PART_NO,PI.PART_TYPE,CM.CUSTOMER_CODE,NM6.LONG_NAME CUSTOMER_NAME, SUM(PD.CAPACITY_AMOUNT) CAPACITY_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INVENTORY_WH_DOC IWD "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (IWD.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM6 ON (CN.NAME_ID = NM6.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (IWD.INVENTORY_WH_DOC_ID = LIW.INVENTORY_WH_DOC_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LIW.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LD ON (LIW.LOT_ITEM_WH_ID = LD.LOT_ITEM_WH_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN PALLET_DOC PD ON (LD.LOT_DOC_ID = PD.LOT_DOC_ID)"
   End If
   
   If INVENTORY_WH_DOC_ID > -1 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (IWD.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      End If
   End If
   
'   If INVENTORY_DOC_ID > -1 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (IWD.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (IWD.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
'      End If
'   End If
   
   'INVENTORY_DOC_ID
   
   If Len(CUSTOMER_CODE) > 1 Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "CUSTOMER_CODE", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "CUSTOMER_CODE", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
  If Len(FROM_CUSTOMER_CODE) > 1 Then
      If WildCard(FROM_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "CUSTOMER_CODE", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "CUSTOMER_CODE", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM.CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   If Len(TO_CUSTOMER_CODE) > 1 Then
      If WildCard(TO_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "CUSTOMER_CODE", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "CUSTOMER_CODE", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM.CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
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
   
   If Len(TRUCK_NO) > 1 Then
      If WildCard(TRUCK_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "IWD.TRUCK_NO", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "IWD.TRUCK_NO", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (IWD.TRUCK_NO = '" & ChangeQuote(Trim(TRUCK_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (IWD.TRUCK_NO = '" & ChangeQuote(Trim(TRUCK_NO)) & "')"
         End If
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
      
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(DOCUMENT_TYPE_SET) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      End If
   End If
   
   If Len(LOAD_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOAD_FLAG = '" & LOAD_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (LOAD_FLAG = '" & LOAD_FLAG & "') "
      End If
   End If
   
   If Len(VERIFY_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.VERIFY_FLAG = '" & VERIFY_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (J.VERIFY_FLAG = '" & VERIFY_FLAG & "') "
      End If
   End If
   
   'VERIFY_FLAG
   If Len(SUCCESS_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.SUCCESS_FLAG = '" & SUCCESS_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (IWD.SUCCESS_FLAG = '" & SUCCESS_FLAG & "') "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (IWD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Trim(DOCUMENT_NO) <> "" Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(IWD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(IWD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (IWD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (IWD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
   

      
      'ตรงนี้
   If DOCUMENT_DATE > -1 Then
      TempDate = Trim(DateToStringInt(DOCUMENT_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IWD.DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If PACK_DATE > -1 Then
      TempDate = Trim(DateToStringInt(PACK_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LIW.PACK_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LIW.PACK_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If

      
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IWD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IWD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If HEAD_PACK_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.HEAD_PACK_NO = " & HEAD_PACK_NO & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.HEAD_PACK_NO = " & HEAD_PACK_NO & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 9 Then
      GroupByStr = "GROUP BY IWD.DOCUMENT_NO,IWD.DOCUMENT_DATE,IWD.TRUCK_NO,IWD.ENTRY_WEIGHT_DATE,IWD.ENTRY_WEIGHT_TIME,IWD.EXIT_WEIGHT_DATE, "
      GroupByStr = GroupByStr & "IWD.EXIT_WEIGHT_TIME,LIW.TX_AMOUNT,PI.PART_NO,PI.PART_TYPE,CM.CUSTOMER_CODE,NM6.LONG_NAME "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
 If Ind = 5 Then
    If OrderBy = 1 Then
      OrderByStr = "ORDER BY LIW.PART_ITEM_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LIW.TIME_PACK_BEGIN "
   Else
      OrderByStr = "ORDER BY IWD.DOCUMENT_NO "
   End If
ElseIf Ind = 6 Then
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY IWD.DOCUMENT_NO   "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM6.LONG_NAME "
   Else
      OrderByStr = "ORDER BY IWD.DOCUMENT_NO "
   End If
ElseIf Ind = 8 Then 'ปล่อยว่าง
ElseIf Ind = 9 Then
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY IWD.DOCUMENT_NO  "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY CM.CUSTOMER_CODE  "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY NM6.LONG_NAME  "
   Else
      OrderByStr = "ORDER BY IWD.DOCUMENT_NO "
   End If
 Else
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY IWD.DOCUMENT_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY IWD.DOCUMENT_DATE "
   Else
      OrderByStr = "ORDER BY IWD.INVENTORY_WH_DOC_ID "
   End If

End If
   
   If Ind <> 8 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then 'ใช้แล้ว
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DELIVERY_ID = NVLI(Rs("DELIVERY_ID"), 0)
      DELIVERY_FEE = NVLD(Rs("DELIVERY_FEE"), 0)
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      DO_NO = NVLS(Rs("DO_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      OLD_COMMIT_FLAG = COMMIT_FLAG
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
      REASON_ID = NVLI(Rs("REASON_ID"), -1)
      QUE_NO = NVLS(Rs("QUE_NO"), "")
      ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
      EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      DEPARTMENT_ID = NVLI(Rs("DEPARTMENT_ID"), -1)
      ADJUST_FLAG = NVLS(Rs("ADJUST_FLAG"), "")
      Credit = NVLD(Rs("CREDIT"), 0)
'      LOT_ID = NVLI(Rs("LOT_ID"), -1)
      BATCH_NO = NVLS(Rs("BATCH_NO"), "")
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      CONDITION1 = NVLI(Rs("CONDITION1"), -1)
      CONDITION2 = NVLI(Rs("CONDITION2"), -1)
      CONDITION3 = NVLI(Rs("CONDITION3"), -1)
      EMP_CHECK_CAR_ID = NVLI(Rs("EMP_CHECK_CAR_ID"), -1)
      EMP_CHECK_CAR_NAME = NVLS(Rs("EMP_CHECK_CAR_NAME"), "")
      EMP_CHECK_CAR_LNAME = NVLS(Rs("EMP_CHECK_CAR_LNAME"), "")
      EMP_CHECK_PRODUCT_ID = NVLI(Rs("EMP_CHECK_PRODUCT_ID"), -1)
      EMP_CHECK_PRODUCT_NAME = NVLS(Rs("EMP_CHECK_PRODUCT_NAME"), "")
      EMP_CHECK_PRODUCT_LNAME = NVLS(Rs("EMP_CHECK_PRODUCT_LNAME"), "")
      RESPONSE_NAME = NVLS(Rs("RESPONSE_NAME"), "")
      RESPONSE_LNAME = NVLS(Rs("RESPONSE_LNAME"), "")
      WEIGHT_ID = NVLS(Rs("WEIGHT_ID"), "")
      ENTRY_WEIGHT = NVLI(Rs("ENTRY_WEIGHT"), -1)
      EXIT_WEIGHT = NVLI(Rs("EXIT_WEIGHT"), -1)
      TOTAL_WEIGHT = NVLI(Rs("TOTAL_WEIGHT"), -1)
      WEIGHT_NOTE = NVLS(Rs("WEIGHT_NOTE"), "")
      ENTRY_WEIGHT_DATE = NVLS(Rs("ENTRY_WEIGHT_DATE"), "")
      ENTRY_WEIGHT_TIME = NVLS(Rs("ENTRY_WEIGHT_TIME"), "")
      EXIT_WEIGHT_DATE = NVLS(Rs("EXIT_WEIGHT_DATE"), "")
      EXIT_WEIGHT_TIME = NVLS(Rs("EXIT_WEIGHT_TIME"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      LOAD_FLAG = NVLS(Rs("LOAD_FLAG"), "")
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
      CREATE_BY = NVLI(Rs("CREATE_BY"), -1)
      MODIFY_BY = NVLI(Rs("MODIFY_BY"), -1)
      CONSIGNMENT_FLAG = NVLS(Rs("CONSIGNMENT_FLAG"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      DOCUMENT_NO_RQ = NVLS(Rs("DOCUMENT_NO_RQ"), "")
      PROCESS_ID = NVLI(Rs("PROCESS_ID"), -1)
   ElseIf Ind = 3 Then
       INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DELIVERY_ID = NVLI(Rs("DELIVERY_ID"), 0)
      DELIVERY_FEE = NVLD(Rs("DELIVERY_FEE"), 0)
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      DO_NO = NVLS(Rs("DO_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      OLD_COMMIT_FLAG = COMMIT_FLAG
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
      REASON_ID = NVLI(Rs("REASON_ID"), -1)
      QUE_NO = NVLS(Rs("QUE_NO"), "")
      ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
      EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      DEPARTMENT_ID = NVLI(Rs("DEPARTMENT_ID"), -1)
      ADJUST_FLAG = NVLS(Rs("ADJUST_FLAG"), "")
      Credit = NVLD(Rs("CREDIT"), 0)
      LOT_ID = NVLI(Rs("LOT_ID"), -1)
      BATCH_NO = NVLS(Rs("BATCH_NO"), "")
      ElseIf Ind = 4 Then
       INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
       DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
       DOCUMENT_NO_SO = NVLS(Rs("DOCUMENT_NO_SO"), "")
       DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      ElseIf Ind = 5 Then  'ใช้แล้ว
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      BATCH_NO = NVLS(Rs("BATCH_NO"), "")
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
    ElseIf Ind = 6 Then  'ใช้แล้ว
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      BATCH_NO = NVLS(Rs("BATCH_NO"), "")
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      NOTE = NVLS(Rs("NOTE"), "")
   ElseIf Ind = 7 Then
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      LOAD_GOODS_NO = NVLS(Rs("LOAD_GOODS_NO"), "")
      DO_NO = NVLS(Rs("DO_NO"), "")
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
       LOAD_FLAG = NVLS(Rs("LOAD_FLAG"), "")
       NOTE = NVLS(Rs("NOTE"), "")
       BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
    ElseIf Ind = 8 Then
      MAX_BALANCE_DATE = InternalDateToDate(NVLS(Rs("MAX_BALANCE_DATE"), ""))
   ElseIf Ind = 9 Then  'ใช้แล้ว
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      ENTRY_WEIGHT_DATE = NVLS(Rs("ENTRY_WEIGHT_DATE"), "")
      ENTRY_WEIGHT_TIME = NVLS(Rs("ENTRY_WEIGHT_TIME"), "")
      EXIT_WEIGHT_DATE = NVLS(Rs("EXIT_WEIGHT_DATE"), "")
      EXIT_WEIGHT_TIME = NVLS(Rs("EXIT_WEIGHT_TIME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String

   If ADJUST_FLAG = "" Then
      ADJUST_FLAG = "N"
   End If

   TempID1 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = DELIVERY_ID
   If DELIVERY_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = EMP_ID
   If EMP_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = REASON_ID
   If REASON_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = DEPARTMENT_ID
   If DEPARTMENT_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
'   TempID6 = INVENTORY_DOC_ID
'   If INVENTORY_DOC_ID <= 0 Then
'      TempID6 = "NULL"
'   End If

   If Len(EXCEPTION_FLAG) = 0 Then
      EXCEPTION_FLAG = "N"
   End If
   
   If Len(SUCCESS_FLAG) = 0 Then
      SUCCESS_FLAG = "N"
   End If
      
   If Len(CONSIGNMENT_FLAG) = 0 Then
      CONSIGNMENT_FLAG = "N"
   End If
   
   If Len(COMMIT_FLAG) = 0 Then
      COMMIT_FLAG = "N"
   End If
   
   
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("INVENTORY_WH_DOC_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO INVENTORY_WH_DOC " & _
                     "( " & _
                     "   INVENTORY_WH_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   DELIVERY_ID, " & _
                     "   DOCUMENT_DESC, " & _
                     "   RECEIVE_NAME, " & _
                     "   BILL_NO, " & _
                     "   DO_NO, " & _
                     "   TRUCK_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   SENDER_NAME, " & _
                     "   DELIVERY_FEE, " & _
                     "   EMP_ID, " & _
                     "   COMMIT_FLAG, " & _
                     "   SALE_FLAG, " & _
                     "   REASON_ID, "
         SQL1 = SQL1 & _
                     "   QUE_NO, " & _
                     "   ENTRY_DATE, " & _
                     "   EXIT_DATE, " & _
                     "   EXCEPTION_FLAG, " & _
                     "   DEPARTMENT_ID, " & _
                     "   ADJUST_FLAG, " & _
                     "   CREDIT, " & _
                     "   YYYYMM,CONDITION1,CONDITION2,CONDITION3,EMP_CHECK_CAR_ID,EMP_CHECK_PRODUCT_ID,ENTRY_WEIGHT,EXIT_WEIGHT,TOTAL_WEIGHT,CUSTOMER_ID,WEIGHT_NOTE,WEIGHT_ID,ENTRY_WEIGHT_DATE,ENTRY_WEIGHT_TIME,EXIT_WEIGHT_DATE,EXIT_WEIGHT_TIME,NOTE,BATCH_NO, " & _
                     "   PROCESS_ID,APPROVED_BY,RESPONSE_BY,START_DATE,FINISH_DATE,SUCCESS_FLAG,LOAD_FLAG,CONSIGNMENT_FLAG," & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(RECEIVE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(BILL_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DO_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(TRUCK_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         DOCUMENT_TYPE & "," & _
                         "'" & ChangeQuote(Trim(SENDER_NAME)) & "'," & _
                         DELIVERY_FEE & "," & _
                         TempID3 & "," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(SALE_FLAG)) & "'," & _
                         TempID4 & "," & _
                         "'" & ChangeQuote(Trim(QUE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "',"
            SQL1 = SQL1 & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EXIT_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                         TempID5 & "," & _
                         "'" & ChangeQuote(Trim(ADJUST_FLAG)) & "'," & _
                         Credit & "," & _
                         "'" & ChangeQuote(Trim(YYYYMM)) & "'," & CONDITION1 & "," & CONDITION2 & "," & CONDITION3 & "," & EMP_CHECK_CAR_ID & "," & EMP_CHECK_PRODUCT_ID & "," & ENTRY_WEIGHT & "," & EXIT_WEIGHT & "," & TOTAL_WEIGHT & "," & TempID1 & ",'" & WEIGHT_NOTE & "','" & WEIGHT_ID & "','" & ENTRY_WEIGHT_DATE & "','" & ENTRY_WEIGHT_TIME & "','" & EXIT_WEIGHT_DATE & "','" & EXIT_WEIGHT_TIME & "','" & NOTE & "','" & BATCH_NO & "'," & _
                         "" & PROCESS_ID & "," & APPROVED_BY & "," & RESPONSE_BY & ",'" & ChangeQuote(Trim(DateToStringInt(START_DATE))) & "','" & ChangeQuote(Trim(DateToStringInt(FINISH_DATE))) & "','" & ChangeQuote(Trim(SUCCESS_FLAG)) & "','" & ChangeQuote(Trim(LOAD_FLAG)) & "','" & ChangeQuote(Trim(CONSIGNMENT_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         INVENTORY_WH_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_WH_DOC SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   DELIVERY_ID = " & TempID2 & ", " & _
                     "   DOCUMENT_DESC = '" & ChangeQuote(Trim(DOCUMENT_DESC)) & "', " & _
                     "   RECEIVE_NAME = '" & ChangeQuote(Trim(RECEIVE_NAME)) & "', " & _
                     "   BILL_NO = '" & ChangeQuote(Trim(BILL_NO)) & "', " & _
                     "   DO_NO = '" & ChangeQuote(Trim(DO_NO)) & "', " & _
                     "   TRUCK_NO = '" & ChangeQuote(Trim(TRUCK_NO)) & "', " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "   SENDER_NAME = '" & ChangeQuote(Trim(SENDER_NAME)) & "', " & _
                     "   DELIVERY_FEE = '" & DELIVERY_FEE & "', " & _
                     "   EMP_ID = " & TempID3 & ", " & _
                     "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   SALE_FLAG = '" & ChangeQuote(Trim(SALE_FLAG)) & "', " & _
                     "   REASON_ID = " & TempID4 & ", " & _
                     "   QUE_NO = '" & ChangeQuote(Trim(QUE_NO)) & "', " & _
                     "   ENTRY_DATE = '" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "', " & _
                     "   EXIT_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXIT_DATE))) & "', " & _
                     "   EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "   DEPARTMENT_ID = " & TempID5 & ", "
            SQL1 = SQL1 & _
                     "   ADJUST_FLAG = '" & ChangeQuote(Trim(ADJUST_FLAG)) & "', " & _
                     "   CREDIT = " & Credit & ", " & _
                     "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "   CONDITION1 = " & CONDITION1 & ",CONDITION2 = " & CONDITION2 & ",CONDITION3 = " & CONDITION3 & ",EMP_CHECK_CAR_ID = " & EMP_CHECK_CAR_ID & ",EMP_CHECK_PRODUCT_ID = " & EMP_CHECK_PRODUCT_ID & ", " & _
                     "   ENTRY_WEIGHT = " & ENTRY_WEIGHT & ",EXIT_WEIGHT = " & EXIT_WEIGHT & ",TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   CUSTOMER_ID = " & TempID1 & ",WEIGHT_NOTE = '" & WEIGHT_NOTE & "',NOTE = '" & NOTE & "', " & _
                     "   WEIGHT_ID = '" & WEIGHT_ID & "',ENTRY_WEIGHT_DATE = '" & ENTRY_WEIGHT_DATE & "',ENTRY_WEIGHT_TIME = '" & ENTRY_WEIGHT_TIME & "',EXIT_WEIGHT_DATE = '" & EXIT_WEIGHT_DATE & "',EXIT_WEIGHT_TIME = '" & EXIT_WEIGHT_TIME & "'," & _
                     "   BATCH_NO = '" & BATCH_NO & "',PROCESS_ID = " & PROCESS_ID & ",APPROVED_BY = " & APPROVED_BY & ",RESPONSE_BY = " & RESPONSE_BY & ", " & _
                     "   START_DATE = '" & ChangeQuote(Trim(DateToStringInt(START_DATE))) & "',FINISH_DATE = '" & ChangeQuote(Trim(DateToStringInt(FINISH_DATE))) & "',SUCCESS_FLAG = '" & ChangeQuote(Trim(SUCCESS_FLAG)) & "',LOAD_FLAG = '" & ChangeQuote(Trim(LOAD_FLAG)) & "',CONSIGNMENT_FLAG = '" & ChangeQuote(Trim(CONSIGNMENT_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PALLET_DOC WHERE LOT_DOC_ID IN (SELECT LOT_DOC_ID FROM LOT_DOC WHERE LOT_ITEM_WH_ID IN (SELECT LOT_ITEM_WH_ID FROM LOT_ITEM_WH WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ")) "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_DOC WHERE LOT_ITEM_WH_ID IN (SELECT LOT_ITEM_WH_ID FROM LOT_ITEM_WH WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM_WH WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID
      m_Conn.Execute (SQL1)
      

         SQL1 = "UPDATE   JOB SET INVENTORY_WH_DOC_ID_INPUT = NULL WHERE INVENTORY_WH_DOC_ID_INPUT=" & INVENTORY_WH_DOC_ID
         m_Conn.Execute (SQL1)

         SQL1 = "UPDATE   JOB SET INVENTORY_WH_DOC_ID= NULL WHERE INVENTORY_WH_DOC_ID =" & INVENTORY_WH_DOC_ID
         m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM INVENTORY_WH_DOC WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID
      m_Conn.Execute (SQL1)

End Function

Public Function PatchDocDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_WH_DOC SET " & _
                  "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function
Public Function UpdateSuccessFlag() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
   UpdateSuccessFlag = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_WH_DOC SET " & _
                  "   SUCCESS_FLAG = '" & SUCCESS_FLAG & "'," & _
                  "   LOAD_FLAG = '" & LOAD_FLAG & "'" & _
                 "WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ""

    m_Conn.Execute (SQL1)
    UpdateSuccessFlag = True
End Function

Public Function UpdateWeight() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

UpdateWeight = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                  
       SQL1 = "UPDATE INVENTORY_WH_DOC SET " & _
                 "WEIGHT_ID = '" & WEIGHT_ID & "', " & _
                 "ENTRY_WEIGHT = " & ENTRY_WEIGHT & ", " & _
                 "EXIT_WEIGHT = " & EXIT_WEIGHT & ", " & _
                 "TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                 "WEIGHT_NOTE = '" & WEIGHT_NOTE & "', " & _
                 "ENTRY_WEIGHT_DATE = '" & ENTRY_WEIGHT_DATE & "'," & _
                 "ENTRY_WEIGHT_TIME = '" & ENTRY_WEIGHT_TIME & "'," & _
                 "EXIT_WEIGHT_DATE = '" & EXIT_WEIGHT_DATE & "'," & _
                 "EXIT_WEIGHT_TIME = '" & EXIT_WEIGHT_TIME & "'," & _
                 "MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID

    m_Conn.Execute (SQL1)
   UpdateWeight = True
End Function
Private Sub Class_Terminate()
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set ImportExports = Nothing
   Set TransferItems = Nothing
   Set C_LotItemsWH = Nothing

End Sub
