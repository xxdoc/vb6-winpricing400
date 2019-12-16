VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryDoc"
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

Public INVENTORY_DOC_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_DESC As String
Public BILL_NO As String
Public DO_NO As String
Public TRUCK_NO As String
Public SUPPLIER_ID As Long
Public SUPPLIER_ID_TRUE As Long
Public DELIVERY_ID As Long
Public DELIVERY_FEE As Double
Public SENDER_NAME As String
Public RECEIVE_NAME As String
Public DOCUMENT_TYPE As Long
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
Public CUSTOMER_ID As Long
Public CUSTOMER_CODE  As String
Public CUSTOMER_NAME  As String

Public OLD_COMMIT_FLAG As String
Public PART_NO As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public SUPPLIER_CODE_TRUE As String
Public SUPPLIER_NAME_TRUE As String
Public RESPONSE_NAME As String
Public RESPONSE_LNAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DEPARTMENT_NAME As String
Public PR_NO As String
Public INVENTORY_WH_DOC_ID As Long

Public BILLING_DOC_ID As Long
Public PART_ITEM_ID As Long
Public WEIGHT_PER_PACK As Long
Public PACK_AMOUNT As Double
Public LOAD_GOODS_NO As String
Public DOCUMENT_NO_SO As String
Public NOTE As String
Public LOAD_FLAG As String
Public B_SUCCESS_FLAG  As String
Public DELIVERY_CUS_ITEM_ID As Long
Public PRICE_THINK_TYPE As Long
Public USER_APPLOVE_PRICE_THINK As Long

Public ImportItems As Collection
Public ExportItems As Collection
Public ImportExports As Collection
Public TransferItems As Collection


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   DOCUMENT_DATE = -1
   COMMIT_FLAG = "N"
   SALE_FLAG = "N"
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   Set ImportExports = New Collection
   Set TransferItems = New Collection

End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long, Optional Ind As Long = 1)
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
   SelectStr1 = "SELECT ADR.*, NM1.LONG_NAME SUPPLIER_NAME, NM4.LONG_NAME SUPPLIER_NAME_TRUE, NM3.LONG_NAME RESPONSE_NAME, NM3.LAST_NAME RESPONSE_LNAME, SP.SUPPLIER_CODE,SP2.SUPPLIER_CODE SUPPLIER_CODE_TRUE, LO.LAY_OUT_NAME DEPARTMENT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM INVENTORY_DOC ADR "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
   
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP2 ON (SP2.SUPPLIER_ID = ADR.SUPPLIER_ID_TRUE) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN2 ON (SN2.SUPPLIER_ID = SP2.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (SN2.NAME_ID = NM4.NAME_ID) "
   
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = ADR.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (ADR.DEPARTMENT_ID = LO.LAY_OUT_ID) "
ElseIf Ind = 2 Then
   SelectStr1 = "SELECT ADR.*, NM1.LONG_NAME SUPPLIER_NAME, NM3.LONG_NAME RESPONSE_NAME, NM3.LAST_NAME RESPONSE_LNAME, SP.SUPPLIER_CODE, LO.LAY_OUT_NAME DEPARTMENT_NAME,CM.CUSTOMER_ID, CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM INVENTORY_DOC ADR "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = ADR.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (ADR.DEPARTMENT_ID = LO.LAY_OUT_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (ADR.CUSTOMER_ID = CM.CUSTOMER_ID)"
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID)"
ElseIf Ind = 3 Then
   SelectStr1 = "SELECT DISTINCT ADR.INVENTORY_DOC_ID,ADR.DOCUMENT_NO,ADR.DOCUMENT_DATE,ADR.INVENTORY_WH_DOC_ID,CM.CUSTOMER_ID,NM.LONG_NAME CUSTOMER_NAME,LIW.BILLING_DOC_ID,  "
   SelectStr1 = SelectStr1 & "IWD.TRUCK_NO,IWD.DOCUMENT_NO LOAD_GOODS_NO,BD.DOCUMENT_NO DOCUMENT_NO_SO,BD.NOTE,BD.DO_NO,IWD.LOAD_FLAG,"
   SelectStr1 = SelectStr1 & "BD.SUCCESS_FLAG B_SUCCESS_FLAG,BD.DELIVERY_CUS_ITEM_ID,BD.PRICE_THINK_TYPE,BD.USER_APPLOVE_PRICE_THINK "
   SelectStr2 = "SELECT COUNT(*)  "

   FromStr = "FROM INVENTORY_DOC ADR "
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (ADR.CUSTOMER_ID = CM.CUSTOMER_ID)"
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN LOT_ITEM_WH LIW ON (ADR.INVENTORY_WH_DOC_ID = LIW.INVENTORY_WH_DOC_ID)  "
   FromStr = FromStr & " LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID)  "
   FromStr = FromStr & " LEFT OUTER JOIN BILLING_DOC BD ON (LIW.BILLING_DOC_ID =BD.BILLING_DOC_ID) "


ElseIf Ind = 4 Then
   SelectStr1 = "SELECT  DISTINCT ADR.INVENTORY_DOC_ID,ADR.DOCUMENT_NO,ADR.DOCUMENT_DATE,CM.CUSTOMER_ID,NM.LONG_NAME CUSTOMER_NAME,LIW.PART_ITEM_ID,LIW.BILLING_DOC_ID,SO.PART_ITEM_ID,SO.PACK_AMOUNT,SO.WEIGHT_PER_PACK  "
   SelectStr2 = "SELECT COUNT(*)  "

   FromStr = "FROM INVENTORY_DOC ADR "
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (ADR.CUSTOMER_ID = CM.CUSTOMER_ID)"
   FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN LOT_ITEM_WH LIW ON (ADR.INVENTORY_WH_DOC_ID = LIW.INVENTORY_WH_DOC_ID)  "
   FromStr = FromStr & " LEFT OUTER JOIN SALE_ORDER SO ON (LIW.BILLING_DOC_ID =SO.DO_ID)  "
End If

   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If INVENTORY_WH_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
'
'   If INVENTORY_WH_DOC_ID = -3 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (INVENTORY_WH_DOC_ID IS  NULL ) "
'      Else
'         WhereStr = WhereStr & "AND (INVENTORY_WH_DOC_ID IS  NULL ) "
'      End If
'   End If
      
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (ADR.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Trim(DOCUMENT_NO) <> "" Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ADR.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (ADR.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
      
   If DOCUMENT_DATE > -1 Then
      TempDate = Trim(DateToStringInt(DOCUMENT_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM LOT_ITEM LI LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (LI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM LOT_ITEM LI, PART_ITEM PI1 WHERE (LI.PART_ITEM_ID = PI1.PART_ITEM_ID) AND (LI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI1.PART_NO = '" & PART_NO & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM LOT_ITEM LI LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (LI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM LOT_ITEM LI, PART_ITEM PI1 WHERE (LI.PART_ITEM_ID = PI1.PART_ITEM_ID) AND (LI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI1.PART_NO = '" & PART_NO & "')) > 0) "
         End If
      End If
   End If
            
          
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ADR.DOCUMENT_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ADR.DOCUMENT_DATE "
   Else
      OrderByStr = "ORDER BY ADR.INVENTORY_DOC_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
   SUPPLIER_ID_TRUE = NVLI(Rs("SUPPLIER_ID_TRUE"), -1)
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
   PR_NO = NVLS(Rs("PR_NO"), "")
   
   SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
   SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
   SUPPLIER_CODE_TRUE = NVLS(Rs("SUPPLIER_CODE_TRUE"), "")
   SUPPLIER_NAME_TRUE = NVLS(Rs("SUPPLIER_NAME_TRUE"), "")
   RESPONSE_NAME = NVLS(Rs("RESPONSE_NAME"), "")
   RESPONSE_LNAME = NVLS(Rs("RESPONSE_LNAME"), "")
   DEPARTMENT_NAME = NVLS(Rs("DEPARTMENT_NAME"), "")
ElseIf Ind = 2 Then
   INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
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
   PR_NO = NVLS(Rs("PR_NO"), "") '
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   
   SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
   SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
   RESPONSE_NAME = NVLS(Rs("RESPONSE_NAME"), "")
   RESPONSE_LNAME = NVLS(Rs("RESPONSE_LNAME"), "")
   DEPARTMENT_NAME = NVLS(Rs("DEPARTMENT_NAME"), "")
ElseIf Ind = 3 Then
   INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
   BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
   TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
   LOAD_GOODS_NO = NVLS(Rs("LOAD_GOODS_NO"), "")
   DOCUMENT_NO_SO = NVLS(Rs("DOCUMENT_NO_SO"), "")
   NOTE = NVLS(Rs("NOTE"), "")
   DO_NO = NVLS(Rs("DO_NO"), "")
   LOAD_FLAG = NVLS(Rs("LOAD_FLAG"), "")
   B_SUCCESS_FLAG = NVLS(Rs("B_SUCCESS_FLAG"), "")
   DELIVERY_CUS_ITEM_ID = NVLI(Rs("DELIVERY_CUS_ITEM_ID"), -1)
   PRICE_THINK_TYPE = NVLI(Rs("PRICE_THINK_TYPE"), -1)
   USER_APPLOVE_PRICE_THINK = NVLI(Rs("USER_APPLOVE_PRICE_THINK"), -1)
   
ElseIf Ind = 4 Then
   INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
   WEIGHT_PER_PACK = NVLI(Rs("WEIGHT_PER_PACK"), -1)
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
Dim TempID8 As String

   If ADJUST_FLAG = "" Then
      ADJUST_FLAG = "N"
   End If
   
   TempID1 = SUPPLIER_ID
   If SUPPLIER_ID <= 0 Then
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
   
   TempID6 = INVENTORY_WH_DOC_ID
   If INVENTORY_WH_DOC_ID <= 0 Then
      TempID6 = "NULL"
   End If
   
   TempID7 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID7 = "NULL"
   End If
   
   TempID8 = SUPPLIER_ID_TRUE
   If SUPPLIER_ID_TRUE <= 0 Then
      TempID8 = "NULL"
   End If


   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("INVENTORY_DOC_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO INVENTORY_DOC " & _
                     "( " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   SUPPLIER_ID, " & _
                     "   SUPPLIER_ID_TRUE, " & _
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
                     "   PR_NO, " & _
                     "   YYYYMM, " & _
                     "   INVENTORY_WH_DOC_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         TempID1 & "," & _
                         TempID8 & "," & _
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
                         "'" & ChangeQuote(Trim(PR_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                         "" & TempID6 & "," & _
                         "" & TempID7 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         INVENTORY_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_DOC SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   SUPPLIER_ID = " & TempID1 & ", " & _
                     "   SUPPLIER_ID_TRUE = " & TempID8 & ", " & _
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
                     "   PR_NO = '" & ChangeQuote(Trim(PR_NO)) & "', " & _
                     "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "   INVENTORY_WH_DOC_ID = " & TempID6 & ", " & _
                     "   CUSTOMER_ID = " & TempID7 & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUBLOT_ITEM WHERE LOT_ITEM_ID IN (SELECT LOT_ITEM_ID FROM LOT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM LOT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
            
      SQL1 = "DELETE FROM INVENTORY_DOC WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
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

Private Sub Class_Terminate()
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set ImportExports = Nothing
   Set ImportExports = Nothing
   Set TransferItems = Nothing
End Sub
