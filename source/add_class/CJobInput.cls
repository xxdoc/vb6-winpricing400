VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobInput"
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
Public JOB_INOUT_ID As Long
Public JOB_ID As Long
Public PART_ITEM_ID As Long
Public PRODUCT_TYPE_ID As Long
Public LOT_ID As String
Public BATCH_NO As String
Public LOT_NO As String
Public BIN_NO As Long
Public PALLET_NO As Long
Public LOCK_NO As Long
Public HEAD_PACK_NO As Long
Public MIX_DATE As Date
Public GOOD_AMOUNT As Double
Public LOSE_AMOUNT As Double
Public WEIGHT_PER_PACK As Double
Public PACK_AMOUNT As Double
Public REST_AMOUNT As Double
Public PACK_DATE As Date
Public TIME_PACK_BEGIN As Date
Public TIME_PACK_END As Date
Public TX_AMOUNT As Double
Public NOTE As String
Public LOCATION_ID As Long
Public LINK_ID As Long
Public TX_TYPE As String
Public SERIAL_NUMBER  As String
Public INOUT_REF As String
Public FROM_FORMULA As Long
Public AVG_PRICE As Double
Public GROUP_NO As Long
Public STD_AMOUNT As Double
Public PARAM_ID As Long

Public JOB_DOC_TYPE As Long
Public FROM_DATE As Date
Public TO_DATE As Date
Public PART_NO  As String
Public PART_DESC  As String
Public LOCATION_NO  As String
Public LOCATION_NAME  As String
Public PART_GROUP_ID As Long
Public PART_TYPE_ID As Long
Public PART_TYPE_NAME As String
Public FROM_FORMULA_NAME As String
Public UNIT_NAME As String
Public INCLUDE_UNIT_PRICE As Double
Public JOB_PART_ITEM_ID As Long
Public TOTAL_INCLUDE_PRICE As Double
Public PROCESS_ID As Long
Public PART_ITEM_SET As String
Public RAW_COST As Double

Public FULL_PALLET_FROM As Long
Public FULL_PALLET_TO As Long
Public FULL_UNIT_PER_PALLET As Long
Public SCRAP_PALLET As Long
Public SCRAP_UNIT_PER_PALLET As Long

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As String
Public ProcessSet As String
Public RX_AMOUNT As Double

'Public Pallet As Collection

Public PARCEL_TYPE As Long
Public YYYYMM As String
Public C_InventoryWhDoc As Collection
'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
'   Set Pallet = New Collection
   Set C_InventoryWhDoc = New Collection
   FROM_DATE = -1
   TO_DATE = -1
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
      SelectStr1 = "SELECT UN.UNIT_NAME, FM.FORMULA_NO FROM_FORMULA_NAME, INP.*,PI.PART_NO,PI.PART_DESC,LT.LOCATION_NO,LT.LOCATION_NAME,PT.PART_TYPE_NAME,PT.PART_TYPE_ID, LI.INCLUDE_UNIT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LT ON (LT.LOCATION_ID = INP.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PT.PART_TYPE_ID = PI.PART_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN FORMULA FM ON (INP.FROM_FORMULA = FM.FORMULA_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, INP.TX_TYPE, SUM(INP.TX_AMOUNT) TX_AMOUNT, SUM(LI.RAW_COST) RAW_COST "
      SelectStr1 = SelectStr1 & ",PI.PARCEL_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""
      
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT JB.PART_ITEM_ID JOB_PART_ITEM_ID,JP.PART_NO,PI.PART_NO, INP.PART_ITEM_ID PART_ITEM_ID, INP.TX_TYPE, PI.PART_NO, PI.PART_DESC, SUM(INP.TX_AMOUNT) TX_AMOUNT, SUM(INP.STD_AMOUNT) STD_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""

      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM JP ON (JP.PART_ITEM_ID = JB.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT JB.PART_ITEM_ID JOB_PART_ITEM_ID, INP.TX_TYPE, SUM(INP.TX_AMOUNT) TX_AMOUNT, SUM(INP.STD_AMOUNT) STD_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE,SUM(LI.RAW_COST) RAW_COST "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""

      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT PI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""

      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT JB.PART_ITEM_ID JOB_PART_ITEM_ID, INP.PART_ITEM_ID PART_ITEM_ID, SUM(INP.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""

      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT JB.PART_ITEM_ID JOB_PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(INP.STD_AMOUNT) STD_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
   
      WhereStr = ""

      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = JB.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(INP.STD_AMOUNT) STD_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT JB.PART_ITEM_ID JOB_PART_ITEM_ID,INP.PARAM_ID,SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
      
      WhereStr = "WHERE (INP.PARAM_ID IS NOT NULL) "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(INP.STD_AMOUNT) STD_AMOUNT "
      SelectStr1 = SelectStr1 & ",JB.YYYYMM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON ((INP.LINK_ID = LI.LINK_ID) AND (LI.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID))"
   ElseIf Ind = 11 Then          'LoadDistinctJobInputByPartItemYYYYMM
      SelectStr1 = "SELECT DISTINCT INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
    ElseIf Ind = 12 Then
      SelectStr1 = "SELECT INP.JOB_ID,INP.PART_ITEM_ID,INP.TX_TYPE,PI.PART_DESC,PI.PART_NO,PT.PART_TYPE_ID,PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      FromStr = "FROM JOB_INOUT INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (INP.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
   End If
   
   If JOB_INOUT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_INOUT_ID = " & JOB_INOUT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_INOUT_ID = " & JOB_INOUT_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INP.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INP.JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If JOB_DOC_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (jb.JOB_DOC_TYPE = " & JOB_DOC_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (jb.JOB_DOC_TYPE = " & JOB_DOC_TYPE & ") "
      End If
   End If
   
   If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (jb.PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (jb.PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
   
   If JOB_PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JB.PART_ITEM_ID = " & JOB_PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JB.PART_ITEM_ID = " & JOB_PART_ITEM_ID & ") "
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   
   If PART_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_TYPE = " & PART_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_TYPE = " & PART_TYPE_ID & ") "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JB.JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JB.JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JB.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JB.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
         
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INP.TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (INP.TX_TYPE = '" & TX_TYPE & "') "
      End If
   End If

   If Len(ProcessSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PROCESS_ID IN " & ProcessSet & ") "
      Else
         WhereStr = WhereStr & "AND (PROCESS_ID IN " & ProcessSet & ") "
      End If
   End If
   
   If Len(PART_ITEM_SET) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INP.PART_ITEM_ID IN " & PART_ITEM_SET & ") "
      Else
         WhereStr = WhereStr & "AND (INP.PART_ITEM_ID IN " & PART_ITEM_SET & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, INP.TX_TYPE,PI.PARCEL_TYPE "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY JB.PART_ITEM_ID, INP.PART_ITEM_ID,JP.PART_NO,PI.PART_NO, INP.TX_TYPE, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY JB.PART_ITEM_ID, INP.TX_TYPE "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY JB.PART_ITEM_ID, INP.PART_ITEM_ID "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY JB.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY JB.PART_ITEM_ID , INP.PARAM_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY INP.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,JB.YYYYMM "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SERIAL_NUMBER "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SERIAL_NUMBER "
      Else
         OrderByStr = "ORDER BY GROUP_NO ASC, JOB_INOUT_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PART_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY JP.PART_NO ASC, PI.PART_NO "
   ElseIf Ind = 4 Then
      'OrderByStr = "ORDER BY PI.PART_ITEM_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 6 Then
      'OrderByStr = "ORDER BY PI.PART_ITEM_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY JB.JOB_ID ASC,INP.TX_TYPE DESC  "
   End If
   
   If Ind = 12 Then
      OrderTypeStr = ""
   Else
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   
   If OrderByStr = "" Then
      OrderTypeStr = ""
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
         JOB_INOUT_ID = NVLI(Rs("JOB_INOUT_ID"), -1)
         JOB_ID = NVLI(Rs("JOB_ID"), -1)
         PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
         TX_AMOUNT = NVLS(Rs("TX_AMOUNT"), "")
         LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
         LINK_ID = NVLI(Rs("LINK_ID"), 0)
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         SERIAL_NUMBER = NVLS(Rs("SERIAL_NUMBER"), "")
         INOUT_REF = NVLS(Rs("INOUT_REF"), "")
         FROM_FORMULA = NVLI(Rs("FROM_FORMULA"), -1)
         FROM_FORMULA_NAME = NVLS(Rs("FROM_FORMULA_NAME"), "")
         AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
         GROUP_NO = NVLI(Rs("GROUP_NO"), -1)
         MIX_DATE = InternalDateToDate(NVLS(Rs("MIX_DATE"), ""))
         STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
         PARAM_ID = NVLI(Rs("PARAM_ID"), 0)
         INCLUDE_UNIT_PRICE = NVLD(Rs("INCLUDE_UNIT_PRICE"), 0)
         UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
         
         PART_NO = NVLS(Rs("PART_NO"), "")
         PART_DESC = NVLS(Rs("PART_DESC"), "")
         LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
         LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
         
         PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
         PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
         
         WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
         PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
        
'         FULL_PALLET_FROM = NVLI(Rs("FULL_PALLET_FROM"), -1)
'         FULL_PALLET_TO = NVLI(Rs("FULL_PALLET_TO"), -1)
'         FULL_UNIT_PER_PALLET = NVLI(Rs("FULL_UNIT_PER_PALLET"), -1)
'         SCRAP_PALLET = NVLI(Rs("SCRAP_PALLET"), -1)
'         SCRAP_UNIT_PER_PALLET = NVLI(Rs("SCRAP_UNIT_PER_PALLET"), -1)
         
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
       PART_NO = NVLS(Rs("PART_NO"), "")
        PART_DESC = NVLS(Rs("PART_DESC"), "")
        RAW_COST = NVLD(Rs("RAW_COST"), 0)
        
        PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), -1)
        
   ElseIf Ind = 3 Then
      JOB_PART_ITEM_ID = NVLI(Rs("JOB_PART_ITEM_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
       PART_NO = NVLS(Rs("PART_NO"), "")
        PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 4 Then
      JOB_PART_ITEM_ID = NVLI(Rs("JOB_PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      RAW_COST = NVLD(Rs("RAW_COST"), 0)
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
       PART_NO = NVLS(Rs("PART_NO"), "")
        PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 6 Then
      JOB_PART_ITEM_ID = NVLI(Rs("JOB_PART_ITEM_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 7 Then
      JOB_PART_ITEM_ID = NVLI(Rs("JOB_PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
   ElseIf Ind = 8 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
   ElseIf Ind = 9 Then
      JOB_PART_ITEM_ID = NVLI(Rs("JOB_PART_ITEM_ID"), -1)
      PARAM_ID = NVLD(Rs("PARAM_ID"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 10 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 11 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
    ElseIf Ind = 12 Then
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
       PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
       PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = FROM_FORMULA
   If FROM_FORMULA <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PARAM_ID
   If PARAM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_INOUT_SEQ", id, glbErrorLog)
      SQL1 = "INSERT INTO JOB_INOUT " & _
                     "( " & _
                     "   JOB_INOUT_ID, " & _
                     "   JOB_ID, " & _
                     "   PART_ITEM_ID," & _
                     "   TX_AMOUNT," & _
                     "   LOCATION_ID, " & _
                     "   LINK_ID,TX_TYPE, " & _
                     "   SERIAL_NUMBER, " & _
                      "   INOUT_REF, " & _
                      "   FROM_FORMULA, " & _
                      "   AVG_PRICE, " & _
                      "   GROUP_NO, " & _
                      "   MIX_DATE," & _
                      "   STD_AMOUNT, " & _
                      "   WEIGHT_PER_PACK, " & _
                      "   PACK_AMOUNT, " & _
                      "  PARAM_ID," & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     id & "," & _
                     "'" & JOB_ID & "'," & _
                     "'" & PART_ITEM_ID & "'," & _
                     "'" & TX_AMOUNT & "'," & _
                     "'" & LOCATION_ID & "'," & _
                     "'" & LINK_ID & "','" & ChangeQuote(Trim(TX_TYPE)) & "'," & _
                     "'" & ChangeQuote(Trim(SERIAL_NUMBER)) & "'," & _
                     "'" & ChangeQuote(Trim(INOUT_REF)) & "'," & _
                     TempID1 & "," & _
                     AVG_PRICE & "," & _
                     GROUP_NO & "," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(MIX_DATE))) & "'," & _
                     STD_AMOUNT & "," & _
                     WEIGHT_PER_PACK & "," & _
                     PACK_AMOUNT & "," & _
                     TempID2 & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
                     JOB_INOUT_ID = id
   Else 'Show edit TIME_PACK_BEGIN
                     SQL1 = "UPDATE JOB_INOUT SET " & _
                     "  PART_ITEM_ID = '" & PART_ITEM_ID & "', " & _
                     "  TX_AMOUNT= '" & TX_AMOUNT & "',LOCATION_ID = '" & LOCATION_ID & "', " & _
                     "  LINK_ID = '" & LINK_ID & "', " & _
                     "  TX_TYPE= '" & ChangeQuote(Trim(TX_TYPE)) & "', " & _
                     "  SERIAL_NUMBER= '" & ChangeQuote(Trim(SERIAL_NUMBER)) & "', " & _
                     "  INOUT_REF= '" & ChangeQuote(Trim(INOUT_REF)) & "', " & _
                     "  FROM_FORMULA = " & TempID1 & ", " & _
                     "  AVG_PRICE = " & AVG_PRICE & ", " & _
                     "  GROUP_NO = " & GROUP_NO & ", " & _
                     "  MIX_DATE = '" & ChangeQuote(Trim(DateToStringInt(MIX_DATE))) & "', " & _
                     "  STD_AMOUNT = " & STD_AMOUNT & ", " & _
                     "  WEIGHT_PER_PACK = " & WEIGHT_PER_PACK & ", " & _
                     "  PACK_AMOUNT = " & PACK_AMOUNT & ", " & _
                     "  PARAM_ID = " & TempID2 & ", " & _
                     "  MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "  MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "  WHERE JOB_INOUT_ID = " & JOB_INOUT_ID
   End If
   
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM JOB_INOUT WHERE JOB_INOUT_ID = " & JOB_INOUT_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
 Set C_InventoryWhDoc = Nothing
'Set Pallet = Nothing
End Sub
