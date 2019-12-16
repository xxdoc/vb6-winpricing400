VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartItem"
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
'
Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_DESC As String
Public UNIT_COUNT As Long
Public PART_TYPE As Long
Public MINIMUM_ALLOW As Double
Public MAXIMUM_ALLOW As Double
Public PIG_FLAG As String
Public PIG_TYPE As String
Public UNIT_WEIGHT As Double
Public BARCODE_NO As String
Public BILL_DESC As String
Public WEIGHT_PER_PACK As Double
Public PARCEL_TYPE As Long
Public OLD_PART_ID As Long
Public NUMBER_PLC_ID As String
Public NUMBER_LAB_ID As String
Public PART_MASTER_ID  As Long
Public PRODUCT_TYPE_ID As Long

Public EXCEPTION_FLAG As String

Public DuptCheck As Boolean

Public ANIMAL_TYPE As Long
Public ANIMAL_NAME As String
'Public CUSTOMER_SALE_TYPE As Long
Public CUSTOMER_SALE_NAME As String
Public DEFAULT_LOCATION As Long
Public DEFAULT_LOCATION_NAME As String

Public LOCATION_ID As Long
Public UNIT_NAME As String

Public PART_TYPE_NAME As String
Public PART_TYPE_NO As String
Public PART_NO_PRODUCT As String
Public PART_TYPE_BAG As Long
Public PART_RX As Long
Public STOCK_DATE As String
Public PART_LOT As Long

Public PART_GROUP_ID As Long
Public PART_GROUP_NO As String
Public PART_GROUP_NAME As String
Public FROM_PART_NO As String
Public TO_PART_NO As String

Public FROM_DATE As Date
Public TO_DATE As Date

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public SUCCESS_FLAG As String
Public JOB_DATE As String
Public JOB_NO  As String
Public MODIFY_DATE As Date
Public TX_AMOUNT As Long
Public CANCEL_FLAG As String

Public CREATE_DATE As Date
Public CREATE_BY As Long
Public MODIFY_BY As Long

Public PartLocations As Collection
Public Suppliers As Collection
Public HumidRates As Collection
Public Pictures As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartLocations = New Collection
   Set Suppliers = New Collection
   Set HumidRates = New Collection
   Set Pictures = New Collection
   
   FROM_DATE = -1
   TO_DATE = -1
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
   SelectStr1 = "SELECT ADR.*,ADR.CREATE_DATE CD,ADR.MODIFY_DATE MD, PT.PART_TYPE_NO,PT.PART_TYPE_NAME,PT.RAW_FLAG,PT.PART_GROUP_ID,PT.COMMISSION_FLAG,UN.UNIT_NAME,UN.UNIT_NO,UN.PERIOD_TYPE,LC.LOCATION_NAME DEFAULT_LOCATION_NAME  " ',MR1.KEY_NAME CUSTOMER_SALE_NAME,MR2.KEY_NAME ANIMAL_NAME
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM ADR  "
   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.UNIT_COUNT = UN.UNIT_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (ADR.DEFAULT_LOCATION = LC.LOCATION_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = ADR.CUSTOMER_SALE_TYPE) "
'   FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = ADR.ANIMAL_TYPE) "
   
ElseIf Ind = 2 Then
   SelectStr1 = "SELECT ADR.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM ADR  "
ElseIf Ind = 3 Then
   SelectStr1 = "SELECT DISTINCT  JIO.TX_AMOUNT , ADR.*, PT.*, UN.* , JV.MODIFY_DATE  MD,J.JOB_NO   "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM ADR  "
   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.UNIT_COUNT = UN.UNIT_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN JOB_VERIFY JV ON ( JV.PART_ITEM_ID = ADR.PART_ITEM_ID ) "
   FromStr = FromStr & " LEFT OUTER JOIN JOB_INOUT JIO ON ( JIO.PART_ITEM_ID = JV.PART_ITEM_ID ) "
   FromStr = FromStr & " LEFT OUTER JOIN JOB  J  ON ( JIO.JOB_ID = J.JOB_ID ) "
   
ElseIf Ind = 4 Then
   SelectStr1 = "SELECT ADR.*, PT.*,PG.*, UN.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM ADR  "
   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.UNIT_COUNT = UN.UNIT_ID) "

 ElseIf Ind = 5 Then  'CReportPlanning003
      SelectStr1 = "SELECT ADR.PART_ITEM_ID,ADR.PART_NO,ADR.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM PART_ITEM ADR "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
End If
  
   
   WhereStr = ""
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If PART_MASTER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_MASTER_ID = " & PART_MASTER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_MASTER_ID = " & PART_MASTER_ID & ") "
      End If
   End If
   
   If PRODUCT_TYPE_ID > 0 Then
         If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID & ") "
      End If
   End If

      
   If UNIT_COUNT > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UNIT_COUNT = " & UNIT_COUNT & ") "
      Else
         WhereStr = WhereStr & "AND (UNIT_COUNT = " & UNIT_COUNT & ") "
      End If
   End If
      
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If Len(CANCEL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      End If
   End If
   
   If Len(PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_TYPE = '" & PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_TYPE = '" & PIG_TYPE & "') "
      End If
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   If Trim(PART_DESC) <> "" Then
      If WildCard(PART_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "PART_DESC", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "PART_DESC", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(BARCODE_NO)) > 0 Then
      If WildCard(BARCODE_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(BARCODE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(BARCODE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "')"
         End If
      End If
   End If
   
   If Len(SUCCESS_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUCCESS_FLAG = '" & SUCCESS_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SUCCESS_FLAG = '" & SUCCESS_FLAG & "') "
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   
   If PART_TYPE_NO <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE_NO = '" & PART_TYPE_NO & "') "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE_NO = '" & PART_TYPE_NO & "') "
      End If
   End If
   
   If EXCEPTION_FLAG <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.CREATE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.CREATE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
    If Len(Trim(FROM_PART_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.PART_NO >= '" & ChangeQuote(Trim(FROM_PART_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.PART_NO >= '" & ChangeQuote(Trim(FROM_PART_NO)) & "')"
      End If
   End If
   
   If Len(Trim(TO_PART_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.PART_NO <= '" & ChangeQuote(Trim(TO_PART_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.PART_NO <= '" & ChangeQuote(Trim(TO_PART_NO)) & "')"
      End If
   End If
         
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 4 Then
      OrderByStr = "ORDER BY  PART_GROUP_NO,PART_TYPE_NO,PART_NO "
   Else
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY  PART_ITEM_ID  "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY  PART_DESC  "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY  PART_TYPE_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY  PART_TYPE_NAME "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY  PART_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY  PART_DESC "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY  BARCODE_NO "
      Else
         OrderByStr = "ORDER BY  PART_NO "
      End If
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
Public Function QueryData2(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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

If Ind = 101 Then
   SelectStr1 = "SELECT PI.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM PI  "
   
End If
  
   
   WhereStr = ""
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If Len(CANCEL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      End If
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(PART_DESC)) > 0 Then
      If WildCard(PART_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "PART_DESC", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "PART_DESC", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         End If
      End If
   End If
   
         
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY  PART_ITEM_ID  "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY  PART_DESC  "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY  PART_TYPE_NO "
   ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY  PART_TYPE_NAME "
   ElseIf OrderBy = 5 Then
      OrderByStr = "ORDER BY  PART_NO "
   ElseIf OrderBy = 6 Then
      OrderByStr = "ORDER BY  PART_DESC "
   ElseIf OrderBy = 7 Then
      OrderByStr = "ORDER BY  BARCODE_NO "
   Else
      OrderByStr = "ORDER BY  PART_NO "
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
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   PART_NO = NVLS(Rs("PART_NO"), "")
   UNIT_COUNT = NVLI(Rs("UNIT_COUNT"), -1)
   MINIMUM_ALLOW = NVLD(Rs("MINIMUM_ALLOW"), 0)
   MAXIMUM_ALLOW = NVLD(Rs("MAXIMUM_ALLOW"), 0)
   PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
   PIG_FLAG = NVLS(Rs("PIG_FLAG"), "")
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
   UNIT_WEIGHT = NVLD(Rs("UNIT_WEIGHT"), 0)
   BARCODE_NO = NVLS(Rs("BARCODE_NO"), "")
   BILL_DESC = NVLS(Rs("BILL_DESC"), "")
   WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
   PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), -1)
   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
   OLD_PART_ID = NVLI(Rs("OLD_PART_ID"), -1)
   CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
   NUMBER_PLC_ID = NVLS(Rs("NUMBER_PLC_ID"), "")
   NUMBER_LAB_ID = NVLS(Rs("NUMBER_LAB_ID"), "")
   PART_MASTER_ID = NVLI(Rs("PART_MASTER_ID"), -1)
   PRODUCT_TYPE_ID = NVLI(Rs("PRODUCT_TYPE_ID"), -1)
   
   DEFAULT_LOCATION = NVLI(Rs("DEFAULT_LOCATION"), -1)
   DEFAULT_LOCATION_NAME = NVLS(Rs("DEFAULT_LOCATION_NAME"), "")
   
   KEY_ID = PART_ITEM_ID
   KEY_LOOKUP = PART_NO
   
   CREATE_DATE = InternalDateToDate(NVLS(Rs("CD"), ""))
   MODIFY_DATE = InternalDateToDate(NVLS(Rs("MD"), ""))
   CREATE_BY = NVLI(Rs("CREATE_BY"), -1)
   MODIFY_BY = NVLI(Rs("MODIFY_BY"), -1)
 ' SUCCESS_FLAG = NVLS(Rs("SUCCESS_FLAG"), "")
 ' PART_TYPE_NO = NVLI(Rs("PART_TYPE_NO"), -1)
   If Ind = 1 Then
   ElseIf Ind = 2 Then
      JOB_DATE = NVLS(Rs("JOB_DATE"), "")
   ElseIf Ind = 3 Then
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MD"), ""))
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), -1)
   ElseIf Ind = 4 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
      PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
      PART_GROUP_NO = NVLS(Rs("PART_GROUP_NO"), "")
    
   End If
   
End Function
Public Function PopulateFromRS2(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 101 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 6 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      KEY_ID = PART_ITEM_ID
      KEY_LOOKUP = PART_NO
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

   TempID1 = PART_TYPE
   If PART_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
      
   TempID2 = UNIT_COUNT
   If UNIT_COUNT <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = NUMBER_PLC_ID
   If Len(NUMBER_PLC_ID) <= 0 Then
      TempID3 = "NULL"
   Else
      TempID3 = "'" & NUMBER_PLC_ID & "'"
   End If
   
   TempID4 = DEFAULT_LOCATION
   If DEFAULT_LOCATION <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = NUMBER_LAB_ID
   If Len(NUMBER_LAB_ID) <= 0 Then
      TempID5 = "NULL"
   Else
      TempID5 = "'" & NUMBER_LAB_ID & "'"
   End If
   
  TempID6 = PART_MASTER_ID
   If PART_MASTER_ID <= 0 Then
      TempID6 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PART_ITEM " & _
                     "( " & _
                     "   PART_ITEM_ID, " & _
                     "   PART_NO, " & _
                     "   UNIT_COUNT, " & _
                     "   MINIMUM_ALLOW, " & _
                     "   MAXIMUM_ALLOW, " & _
                     "   PART_TYPE, " & _
                     "   PIG_FLAG, " & _
                     "   PART_DESC, " & _
                     "   PIG_TYPE, " & _
                     "   UNIT_WEIGHT, " & _
                     "   BARCODE_NO, " & _
                     "   BILL_DESC, " & _
                     "   WEIGHT_PER_PACK, " & _
                     "   PARCEL_TYPE, " & _
                     "   OLD_PART_ID, " & _
                     "   CANCEL_FLAG,DEFAULT_LOCATION, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY, " & _
                    "    NUMBER_PLC_ID,NUMBER_LAB_ID,PART_MASTER_ID,PRODUCT_TYPE_ID " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PART_NO)) & "'," & _
                         "" & TempID2 & "," & _
                         "'" & ChangeQuote(Trim(MINIMUM_ALLOW)) & "'," & _
                         "'" & ChangeQuote(Trim(MAXIMUM_ALLOW)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PIG_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(PIG_TYPE)) & "'," & _
                         UNIT_WEIGHT & "," & _
                         "'" & ChangeQuote(Trim(BARCODE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(BILL_DESC)) & "'," & _
                         WEIGHT_PER_PACK & "," & _
                         PARCEL_TYPE & "," & _
                         OLD_PART_ID & "," & _
                         "'" & ChangeQuote(Trim(CANCEL_FLAG)) & "'," & "" & TempID4 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "" & TempID3 & "," & TempID5 & "," & TempID5 & "," & PRODUCT_TYPE_ID & _
                     ") "
         PART_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_ITEM SET " & _
                     "   PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "', " & _
                     "   UNIT_COUNT = " & TempID2 & ", " & _
                     "   MINIMUM_ALLOW = '" & MINIMUM_ALLOW & "', " & _
                     "   MAXIMUM_ALLOW = '" & MAXIMUM_ALLOW & "', " & _
                     "   PART_TYPE = " & TempID1 & ", " & _
                     "   PIG_FLAG = '" & ChangeQuote(Trim(PIG_FLAG)) & "', " & _
                     "   PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "', " & _
                     "   PIG_TYPE = '" & ChangeQuote(Trim(PIG_TYPE)) & "', " & _
                     "   UNIT_WEIGHT = " & UNIT_WEIGHT & ", " & _
                     "   BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "', " & _
                     "   BILL_DESC = '" & ChangeQuote(Trim(BILL_DESC)) & "', " & _
                     "   WEIGHT_PER_PACK = " & WEIGHT_PER_PACK & ", " & _
                     "   PARCEL_TYPE = " & PARCEL_TYPE & ", " & _
                     "   OLD_PART_ID = " & OLD_PART_ID & ", " & _
                     "   CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "', " & _
                     "   DEFAULT_LOCATION = " & TempID4 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & ", " & _
                    "   NUMBER_PLC_ID = " & TempID3 & "," & _
                    "   NUMBER_LAB_ID = " & TempID5 & "," & _
                    "   PART_MASTER_ID = " & TempID6 & "," & _
                    "   PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID & _
                     " WHERE PART_ITEM_ID = " & PART_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PART_LOCATION WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)

End Function

Private Sub Class_Terminate()
   Set PartLocations = Nothing
   Set Suppliers = Nothing
   Set HumidRates = Nothing
   Set Pictures = Nothing
End Sub
Public Sub GeneratePIHeader(FileID As Long)
Dim TempStr As String
   
   TempStr = "PI;"
   TempStr = TempStr & PART_NO & ";"
   TempStr = TempStr & UNIT_COUNT & ";"
   TempStr = TempStr & MINIMUM_ALLOW & ";"
   
   TempStr = TempStr & MAXIMUM_ALLOW & ";"
   TempStr = TempStr & PART_TYPE & ";"
   TempStr = TempStr & PIG_FLAG & ";"
   TempStr = TempStr & PART_DESC & ";"
   
   TempStr = TempStr & UNIT_WEIGHT & ";"
   TempStr = TempStr & BARCODE_NO & ";"
   TempStr = TempStr & BILL_DESC & ";"
   TempStr = TempStr & WEIGHT_PER_PACK & ";"
   TempStr = TempStr & PARCEL_TYPE & ";"
   
   Print #FileID, TempStr
End Sub

Public Function UpdatePlcPartNo()
Dim SQL1 As String
      
      SQL1 = "UPDATE PART_ITEM SET NUMBER_PLC_ID =  '" & ChangeQuote(Trim(NUMBER_PLC_ID)) & "' WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)
      
End Function
Public Function UpdatePlcPartLocation()
Dim SQL1 As String
      
      SQL1 = "UPDATE PART_ITEM SET DEFAULT_LOCATION =  " & DEFAULT_LOCATION & " WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)
      
End Function
Public Function UpdateLabPartNo()
Dim SQL1 As String
      
      SQL1 = "UPDATE PART_ITEM SET NUMBER_LAB_ID =  '" & ChangeQuote(Trim(NUMBER_LAB_ID)) & "' WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)
      
End Function
Public Function UpdateCancelFlag()
Dim SQL1 As String
      
      SQL1 = "UPDATE PART_ITEM SET CANCEL_FLAG =  '" & ChangeQuote(Trim(CANCEL_FLAG)) & "' WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)
      
End Function

