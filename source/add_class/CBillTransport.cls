VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillTransport"
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
Public BILL_TRANSPORT_ID As Long
Public BILL_TRANSPORT_ITEM_ID As Long
Public BILLING_DOC_ID As Long
Public SUPPLIER_TRANSPORT_ID As Long
Public SUPPLIER_TRANSPORT_CODE As String
Public NOT_CAL_VAT As String
Public CAL_RATE_DELIVERY_TYPE As Long
Public EX_DELIVERY_COST_ITEM_ID As Long


Public AROUND As Long
Public SHIPPING As String
Public SHIPPING_DATE As Date
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public NOTE_SHIPPING As String
Public PAYMENT_DESC As String
Public NOTE As String
Public NOTE_ITEM  As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DATE_BD As Date
Public TO_DATE_BD As Date

Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public CUSTOMER_CODE As String
Public KEY_CODE As String
Public RATE_PER_UNIT As Double
Public CAL_DRIVER As String
Public CAL_CUSTOMER As String
Public CAL_PRICE_IN_PRODUCT As String
Public TOTAL_PRICE As Double
Public WEIGHT_PER_UNIT As Long
Public PACK_AMOUNT As Long
Public CUSTOMER_NAME As String
Public FIX_RATE_DETAIL As String

Public C_BillTransportItem  As Collection
Public CREATE_DATE As Date

Public Truck As String
Public ROUND As Long
Public C10 As Double
Public C20 As Double
Public C30 As Double
Public C50 As Double
Public Cb As Double
Public RATE_DRIVER As Double
Public RATE_CUSTOMER As Double
Public RATE_SALE As Double
Public RATE_FAC As Double
Public RATE_IMCOME_OTHER As Double
Public RATE_EXPEND_OTHER As Double
Public PRICE_DRIVER As Double
Public PRICE_CUSTOMER As Double
Public PRICE_SALE  As Double
Public PRICE_FAC  As Double
Public PRICE_IMCOME_OTHER As Double
Public PRICE_EXPEND_OTHER As Double
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   SHIPPING_DATE = -1
   FROM_DATE_BD = -1
   TO_DATE_BD = -1
   Set C_BillTransportItem = New Collection
End Sub
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
Dim GroupByStr As String
Dim TempDate As String

   If Ind = 1 Then
      SelectStr1 = "SELECT BT.* "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM BILL_TRANSPORT BT "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT IVD.BILLING_DOC_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,NM.LONG_NAME CUSTOMER_NAME,CM.CUSTOMER_CODE,BT.BILL_TRANSPORT_ID,BT.AROUND,BT.SHIPPING,BT.FIX_RATE_DETAIL,BT.NOT_CAL_VAT,BT.NOTE,BT.SHIPPING_DATE,BT.CAL_RATE_DELIVERY_TYPE,BT.EX_DELIVERY_COST_ITEM_ID,  "
      SelectStr1 = SelectStr1 & "BTI.BILL_TRANSPORT_ITEM_ID,BTI.BILL_TYPE_ID,BTI.WEIGHT_PER_UNIT,BTI.PACK_AMOUNT,BTI.RATE_PER_UNIT,BTI.TOTAL_PRICE,BTI.CAL_DRIVER,BTI.CAL_CUSTOMER,BTI.CAL_PRICE_IN_PRODUCT,BTI.NOTE NOTE_ITEM,ST.*,MR.KEY_CODE  "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM BILL_TRANSPORT BT "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (BT.BILLING_DOC_ID = IVD.BILLING_DOC_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN BILL_TRANSPORT_ITEM BTI ON (BT.BILL_TRANSPORT_ID = BTI.BILL_TRANSPORT_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_TRANSPORT ST ON (BT.SUPPLIER_TRANSPORT_ID =ST.SUPPLIER_TRANSPORT_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (ST.SUPPLIER_ID =SP.SUPPLIER_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (BTI.BILL_TYPE_ID=MR.KEY_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT IVD.BILLING_DOC_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,IVD.NOTE NOTE_SHIPPING,IVD.PAYMENT_DESC,NM.LONG_NAME CUSTOMER_NAME,CM.CUSTOMER_CODE,BT.BILL_TRANSPORT_ID,BT.AROUND,BT.SHIPPING,BT.FIX_RATE_DETAIL,BT.NOT_CAL_VAT,BT.NOTE,BT.SHIPPING_DATE,BT.CAL_RATE_DELIVERY_TYPE,BT.EX_DELIVERY_COST_ITEM_ID, "
      SelectStr1 = SelectStr1 & "BTI.BILL_TRANSPORT_ITEM_ID,BTI.BILL_TYPE_ID,BTI.WEIGHT_PER_UNIT,BTI.PACK_AMOUNT,BTI.RATE_PER_UNIT,BTI.TOTAL_PRICE,BTI.CAL_DRIVER,BTI.CAL_CUSTOMER,BTI.CAL_PRICE_IN_PRODUCT,BTI.NOTE NOTE_ITEM,ST.*,MR.KEY_CODE,SNM.LONG_NAME SUPPLIER_NAME, SP.SUPPLIER_CODE  "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM BILLING_DOC IVD  "
      FromStr = FromStr & "LEFT OUTER JOIN BILL_TRANSPORT BT  ON (IVD.BILLING_DOC_ID = BT.BILLING_DOC_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN BILL_TRANSPORT_ITEM BTI ON (BT.BILL_TRANSPORT_ID = BTI.BILL_TRANSPORT_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_TRANSPORT ST ON (BT.SUPPLIER_TRANSPORT_ID =ST.SUPPLIER_TRANSPORT_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (ST.SUPPLIER_ID =SP.SUPPLIER_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SP.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME SNM ON (SN.NAME_ID = SNM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (BTI.BILL_TYPE_ID=MR.KEY_ID)  "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (IVD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   End If
   
   If BILL_TRANSPORT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BT.BILL_TRANSPORT_ID = " & BILL_TRANSPORT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BT.BILL_TRANSPORT_ID = " & BILL_TRANSPORT_ID & ") "
      End If
   End If
   
   If Len(SUPPLIER_TRANSPORT_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ST.SUPPLIER_TRANSPORT_CODE = '" & SUPPLIER_TRANSPORT_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (ST.SUPPLIER_TRANSPORT_CODE = '" & SUPPLIER_TRANSPORT_CODE & "') "
      End If
   End If
   
   'SUPPLIER_TRANSPORT_CODE
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BT.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BT.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
     If Len(Trim(DOCUMENT_NO)) > 0 Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(IVD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(IVD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (IVD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (IVD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If

   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   'SUPPLIER_CODE
   
    If Len(KEY_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MR.KEY_CODE = '" & KEY_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (MR.KEY_CODE = '" & KEY_CODE & "') "
      End If
   End If
   
   If Len(SUPPLIER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SP.SUPPLIER_CODE = '" & SUPPLIER_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (SP.SUPPLIER_CODE = '" & SUPPLIER_CODE & "') "
      End If
   End If
   
   If Len(CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_CODE = '" & CUSTOMER_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE = '" & CUSTOMER_CODE & "') "
      End If
   End If
  
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BT.SHIPPING_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BT.SHIPPING_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BT.SHIPPING_DATE  <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BT.SHIPPING_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_DATE_BD > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE_BD))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE_BD > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE_BD))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If SHIPPING_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(SHIPPING_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BT.SHIPPING_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BT.SHIPPING_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   
   GroupByStr = ""
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BT.BILL_TRANSPORT_ID "
      Else
        OrderByStr = "ORDER BY BT.BILL_TRANSPORT_ID "
      End If
   ElseIf Ind = 2 Then
       If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY BT.BILL_TRANSPORT_ID "
      End If
   ElseIf Ind = 3 Then
       If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO  "
      End If
   Else
      OrderByStr = "ORDER BY BT.BILL_TRANSPORT_ID "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
         BILL_TRANSPORT_ID = NVLI(Rs("BILL_TRANSPORT_ID"), -1)
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         SUPPLIER_TRANSPORT_ID = NVLI(Rs("SUPPLIER_TRANSPORT_ID"), -1)
         AROUND = NVLI(Rs("AROUND"), -1)
         SHIPPING = NVLS(Rs("SHIPPING"), "")
         SHIPPING_DATE = InternalDateToDate(NVLS(Rs("SHIPPING_DATE"), ""))
         FIX_RATE_DETAIL = NVLS(Rs("FIX_RATE_DETAIL"), "")
         NOT_CAL_VAT = NVLS(Rs("NOT_CAL_VAT"), "")
         NOTE = NVLS(Rs("NOTE"), "")
         CAL_RATE_DELIVERY_TYPE = NVLI(Rs("CAL_RATE_DELIVERY_TYPE"), -1)
         EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
      ElseIf Ind = 2 Then
         BILL_TRANSPORT_ID = NVLI(Rs("BILL_TRANSPORT_ID"), -1)
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         SUPPLIER_TRANSPORT_ID = NVLI(Rs("SUPPLIER_TRANSPORT_ID"), -1)
         SUPPLIER_TRANSPORT_CODE = NVLS(Rs("SUPPLIER_TRANSPORT_CODE"), "")
         AROUND = NVLI(Rs("AROUND"), -1)
         SHIPPING = NVLS(Rs("SHIPPING"), "")
         SHIPPING_DATE = InternalDateToDate(NVLS(Rs("SHIPPING_DATE"), ""))
         FIX_RATE_DETAIL = NVLS(Rs("FIX_RATE_DETAIL"), "")
         NOT_CAL_VAT = NVLS(Rs("NOT_CAL_VAT"), "")
         NOTE = NVLS(Rs("NOTE"), "")
         NOTE_ITEM = NVLS(Rs("NOTE_ITEM"), "")
         CAL_DRIVER = NVLS(Rs("CAL_DRIVER"), "")
         CAL_CUSTOMER = NVLS(Rs("CAL_CUSTOMER"), "")
         CAL_PRICE_IN_PRODUCT = NVLS(Rs("CAL_PRICE_IN_PRODUCT"), "")
         RATE_PER_UNIT = NVLD(Rs("RATE_PER_UNIT"), 0)
         TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
         WEIGHT_PER_UNIT = NVLI(Rs("WEIGHT_PER_UNIT"), -1)
         PACK_AMOUNT = NVLI(Rs("PACK_AMOUNT"), -1)
         KEY_CODE = NVLS(Rs("KEY_CODE"), "")
         CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
         CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
         BILL_TRANSPORT_ITEM_ID = NVLI(Rs("BILL_TRANSPORT_ITEM_ID"), -1)
         CAL_RATE_DELIVERY_TYPE = NVLI(Rs("CAL_RATE_DELIVERY_TYPE"), -1)
         EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
      ElseIf Ind = 3 Then
         BILL_TRANSPORT_ID = NVLI(Rs("BILL_TRANSPORT_ID"), -1)
         BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         NOTE_SHIPPING = NVLS(Rs("NOTE_SHIPPING"), "")
         PAYMENT_DESC = NVLS(Rs("PAYMENT_DESC"), "")
         SUPPLIER_TRANSPORT_ID = NVLI(Rs("SUPPLIER_TRANSPORT_ID"), -1)
         SUPPLIER_TRANSPORT_CODE = NVLS(Rs("SUPPLIER_TRANSPORT_CODE"), "")
         SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
         SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
         AROUND = NVLI(Rs("AROUND"), -1)
         SHIPPING = NVLS(Rs("SHIPPING"), "")
         SHIPPING_DATE = InternalDateToDate(NVLS(Rs("SHIPPING_DATE"), ""))
         FIX_RATE_DETAIL = NVLS(Rs("FIX_RATE_DETAIL"), "")
         NOT_CAL_VAT = NVLS(Rs("NOT_CAL_VAT"), "")
         NOTE = NVLS(Rs("NOTE"), "")
         NOTE_ITEM = NVLS(Rs("NOTE_ITEM"), "")
         CAL_DRIVER = NVLS(Rs("CAL_DRIVER"), "")
         CAL_CUSTOMER = NVLS(Rs("CAL_CUSTOMER"), "")
         CAL_PRICE_IN_PRODUCT = NVLS(Rs("CAL_PRICE_IN_PRODUCT"), "")
         RATE_PER_UNIT = NVLD(Rs("RATE_PER_UNIT"), 0)
         TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
         WEIGHT_PER_UNIT = NVLI(Rs("WEIGHT_PER_UNIT"), -1)
         PACK_AMOUNT = NVLI(Rs("PACK_AMOUNT"), -1)
         KEY_CODE = NVLS(Rs("KEY_CODE"), "")
         CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
         CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
         BILL_TRANSPORT_ITEM_ID = NVLI(Rs("BILL_TRANSPORT_ITEM_ID"), -1)
         CAL_RATE_DELIVERY_TYPE = NVLI(Rs("CAL_RATE_DELIVERY_TYPE"), -1)
         EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2  As String

   TempID1 = BILLING_DOC_ID
   If BILLING_DOC_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = SUPPLIER_TRANSPORT_ID
   If SUPPLIER_TRANSPORT_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILL_TRANSPORT_SEQ", id, glbErrorLog)
      SQL1 = "INSERT INTO BILL_TRANSPORT " & _
                     "( " & _
                     "BILL_TRANSPORT_ID, " & _
                     "BILLING_DOC_ID, " & _
                     "SUPPLIER_TRANSPORT_ID, " & _
                     "AROUND, " & _
                     "SHIPPING,SHIPPING_DATE, FIX_RATE_DETAIL,NOT_CAL_VAT,NOTE,CAL_RATE_DELIVERY_TYPE,EX_DELIVERY_COST_ITEM_ID, " & _
                     "CREATE_DATE, " & _
                    "CREATE_BY, " & _
                    "MODIFY_DATE, " & _
                    "MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     id & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     AROUND & "," & _
                     "'" & SHIPPING & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(SHIPPING_DATE))) & "'," & _
                     "'" & FIX_RATE_DETAIL & "','" & NOT_CAL_VAT & "','" & NOTE & "'," & CAL_RATE_DELIVERY_TYPE & "," & EX_DELIVERY_COST_ITEM_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
                     BILL_TRANSPORT_ID = id
   Else 'Show edit
                     SQL1 = "UPDATE BILL_TRANSPORT SET " & _
                     "  BILLING_DOC_ID  = " & TempID1 & ", " & _
                     "  SUPPLIER_TRANSPORT_ID  = " & TempID2 & ", " & _
                     "  AROUND    = " & AROUND & ", " & _
                     "  SHIPPING     = '" & SHIPPING & "', " & _
                     "  SHIPPING_DATE     = '" & ChangeQuote(Trim(DateToStringInt(SHIPPING_DATE))) & "', " & _
                     "  FIX_RATE_DETAIL     = '" & FIX_RATE_DETAIL & "', " & _
                     "  NOT_CAL_VAT     = '" & NOT_CAL_VAT & "', " & _
                     "  NOTE     = '" & NOTE & "', " & _
                     "  CAL_RATE_DELIVERY_TYPE    = " & CAL_RATE_DELIVERY_TYPE & ", " & _
                     "  EX_DELIVERY_COST_ITEM_ID    = " & EX_DELIVERY_COST_ITEM_ID & ", " & _
                     "  MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "  MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "  WHERE BILL_TRANSPORT_ID  = " & BILL_TRANSPORT_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
   SQL1 = "DELETE FROM BILL_TRANSPORT_ITEM WHERE BILL_TRANSPORT_ID = " & BILL_TRANSPORT_ID
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM BILL_TRANSPORT WHERE BILL_TRANSPORT_ID = " & BILL_TRANSPORT_ID
   m_Conn.Execute (SQL1)
End Function


Private Sub Class_Terminate()
   Set C_BillTransportItem = Nothing
End Sub
