VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPalletDoc"
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
Public PALLET_DOC_ID As Long
Public LOT_ITEM_WH_ID As Long
Public PALLET_DOC_NO As String
Public PALLET_DOC_NO_OLD As String
Public PART_ITEM_ID As Long
Public CAPACITY_AMOUNT As Double
Public BALANCE_FLAG As String
Public BALANCE_FLAG_F_I As String
Public MODIFY_DATE As Date
Public TRUCK_NO As String
Public EMP_ID As Long
Public LOT_DOC_ID As Long
Public LOT_DOC_ID_REF As Long
Public LOT_ID As Long
Public LOT_NO As String
Public LOCK_NO As Long
Public LOCK_NAME As String
Public LOCATION_ID As Long
Public NOT_PALLET_DOC_ID As String
Public INVENTORY_WH_DOC_ID As Long
Public START_DATE As Date
Public TX_TYPE As String
Public JOB_INOUT_ID As Long
Public PALLET_CAP_LAST As Double
Public TEMP_PALLET_CAP_LAST As Double
Public HEAD_PACK_NO As Long
Public BIN_NO As Long
Public DOCUMENT_NO  As String
Public DOCUMENT_TYPE As Long
Public DOCUMENT_TYPE_SET As String
'Public BALANCE_DOC_TYPE As Long
Public FULL_PALLET_FROM As Long
Public FULL_PALLET_TO As Long
Public FULL_UNIT_PER_PALLET As Long
Public SCRAP_PALLET As Long
Public SCRAP_UNIT_PER_PALLET As Long
'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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

   If Ind = 1 Then 'ใช้แล้ว
      SelectStr1 = "SELECT PLD.* "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
    ElseIf Ind = 2 Then 'ใช้แล้ว
      SelectStr1 = "SELECT DISTINCT(PLD.PALLET_DOC_ID),PLD.PALLET_DOC_NO,PLD.CAPACITY_AMOUNT,PLD.TRUCK_NO,PLD.EMP_ID,PLD.BALANCE_FLAG,PLD.LOT_DOC_ID,PLD.TX_TYPE,LT.*,LIW.*,PLD2.BALANCE_FLAG BALANCE_FLAG_F_I "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD2 ON (LTD.LOT_DOC_ID_REF = LTD2.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PALLET_DOC PLD2 ON (LTD2.LOT_DOC_ID = PLD2.LOT_DOC_ID) "
'      SelectStr1 = "SELECT PLD.*,LT.*,LIW.* "
'      SelectStr2 = "SELECT COUNT(*)  "
'
'      WhereStr = ""
'
'      FromStr = "FROM PALLET_DOC PLD "
'      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 3 Then 'ใช้แล้ว
       SelectStr1 = "SELECT PLD.* "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""

      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (LIW.INVENTORY_WH_DOC_ID =JB.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB_INOUT JIO ON (JB.JOB_ID =JIO.JOB_ID) "
   ElseIf Ind = 4 Then 'ใช้แล้ว
       SelectStr1 = "SELECT PLD.* "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC  IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
    ElseIf Ind = 5 Then 'ใช้แล้ว
      SelectStr1 = "SELECT PLD.PALLET_DOC_NO,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LT.LOT_ID,LT.LOT_NO,LIW.HEAD_PACK_NO,LIW.PART_ITEM_ID,PLD.TX_TYPE,LIW.BIN_NO,IWD.DOCUMENT_NO,IWD.DOCUMENT_TYPE,SUM(PLD.CAPACITY_AMOUNT) CAPACITY_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC  IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 6 Then 'ใช้แล้ว
       SelectStr1 = "SELECT PLD.TX_TYPE,LIW.HEAD_PACK_NO,MAX(PLD.PALLET_DOC_NO) AS PALLET_DOC_NO " ',LIW.FULL_PALLET_FROM,LIW.FULL_PALLET_TO,LIW.FULL_UNIT_PER_PALLET,LIW.SCRAP_PALLET,LIW.SCRAP_UNIT_PER_PALLET
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID = LIW.LOT_ITEM_WH_ID) "
   ElseIf Ind = 7 Then 'ใช้แล้ว
       SelectStr1 = "SELECT PLD.*,LIW.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID = LIW.LOT_ITEM_WH_ID) "
   ElseIf Ind = 8 Then  'ใช้แล้ว
       SelectStr1 = "SELECT IWD.INVENTORY_WH_DOC_ID,IWD.MODIFY_DATE  "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 9 Then 'ใช้แล้ว
      SelectStr1 = "SELECT PLD.PALLET_DOC_ID,PLD.PALLET_DOC_NO,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LTD.LOT_ITEM_WH_ID,LT.LOT_ID,LT.LOT_NO,LIW.HEAD_PACK_NO,LIW.PART_ITEM_ID,PLD.TX_TYPE,LTD.BIN_NO,IWD.DOCUMENT_NO,IWD.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID =LTD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC  IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 10 Then  'ใช้แล้ว
       SelectStr1 = "SELECT PLD.PALLET_DOC_NO,PLD.CAPACITY_AMOUNT,LT.LOT_NO,LT.LOT_ID,LIW.LOCATION_ID,LIW.HEAD_PACK_NO  "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      
      FromStr = "FROM PALLET_DOC PLD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (PLD.LOT_DOC_ID = LTD.LOT_DOC_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "

   End If

   If PALLET_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PLD.PALLET_DOC_ID = " & PALLET_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PLD.PALLET_DOC_ID = " & PALLET_DOC_ID & ") "
      End If
   End If
   
   If LOT_ITEM_WH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.LOT_ITEM_WH_ID = " & LOT_ITEM_WH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.LOT_ITEM_WH_ID = " & LOT_ITEM_WH_ID & ") "
      End If
   End If
   
   If Len(PALLET_DOC_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PLD.PALLET_DOC_NO = '" & PALLET_DOC_NO & "') "
      Else
         WhereStr = WhereStr & "AND (PLD.PALLET_DOC_NO ='" & PALLET_DOC_NO & "') "
      End If
   End If
   
    If LOT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.LOT_ID = " & LOT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LTD.LOT_ID = " & LOT_ID & ") "
      End If
   End If
   
   If LOT_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PLD.LOT_DOC_ID = " & LOT_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PLD.LOT_DOC_ID = " & LOT_DOC_ID & ") "
      End If
   End If
   
  If LOT_DOC_ID_REF > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.LOT_DOC_ID_REF = " & LOT_DOC_ID_REF & ") "
      Else
         WhereStr = WhereStr & "AND (LTD.LOT_DOC_ID_REF = " & LOT_DOC_ID_REF & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PLD.TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PLD.TX_TYPE = '" & TX_TYPE & "') "
      End If
   End If
   
   If Len(NOT_PALLET_DOC_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PLD.PALLET_DOC_ID NOT IN  (" & NOT_PALLET_DOC_ID & ")) "
      Else
         WhereStr = WhereStr & "AND (PLD.PALLET_DOC_ID NOT IN (" & NOT_PALLET_DOC_ID & ")) "
      End If
   End If
   
   If Len(BALANCE_FLAG) > 0 Then
       If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.BALANCE_FLAG = '" & BALANCE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (LTD.BALANCE_FLAG = '" & BALANCE_FLAG & "') "
      End If
   End If

   If JOB_INOUT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JIO.JOB_INOUT_ID = (" & JOB_INOUT_ID & ")) "
      Else
         WhereStr = WhereStr & "AND (JIO.JOB_INOUT_ID =(" & JOB_INOUT_ID & ")) "
      End If
   End If

   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.PART_ITEM_ID =" & PART_ITEM_ID & ") "
      End If
   End If
   
   If INVENTORY_WH_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (IWD.INVENTORY_WH_DOC_ID =" & INVENTORY_WH_DOC_ID & ") "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_TYPE =" & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(DOCUMENT_TYPE_SET) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      Else
         WhereStr = WhereStr & "AND (IWD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      End If
   End If
   
   If HEAD_PACK_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.HEAD_PACK_NO = " & HEAD_PACK_NO & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.HEAD_PACK_NO =" & HEAD_PACK_NO & ") "
      End If
   End If

   GroupByStr = ""
   If Ind = 5 Then
      GroupByStr = "GROUP BY PLD.PALLET_DOC_NO,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LT.LOT_ID,LT.LOT_NO,LIW.HEAD_PACK_NO,LIW.PART_ITEM_ID,PLD.TX_TYPE,LIW.BIN_NO,IWD.DOCUMENT_NO,IWD.DOCUMENT_TYPE "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY PLD.TX_TYPE,LIW.HEAD_PACK_NO  " ',LIW.FULL_PALLET_FROM,LIW.FULL_PALLET_TO,LIW.FULL_UNIT_PER_PALLET,LIW.SCRAP_PALLET,LIW.SCRAP_UNIT_PER_PALLET
   End If
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Or Ind = 2 Or Ind = 3 Or Ind = 4 Or Ind = 5 Or Ind = 9 Or Ind = 10 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PLD.PALLET_DOC_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PLD.PALLET_DOC_NO "
      Else
         OrderByStr = "ORDER BY PLD.PALLET_DOC_ID "
      End If
   ElseIf Ind <> 6 Then
      OrderByStr = "ORDER BY PLD.TX_TYPE "
   End If
   
   If Ind = 1 Or Ind = 2 Or Ind = 3 Or Ind = 4 Or Ind = 5 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   ElseIf Ind <> 6 Then
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
         PALLET_DOC_ID = NVLI(Rs("PALLET_DOC_ID"), -1)
'         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
         TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
         EMP_ID = NVLI(Rs("EMP_ID"), -1)
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      ElseIf Ind = 2 Then
         PALLET_DOC_ID = NVLI(Rs("PALLET_DOC_ID"), -1)
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
         PALLET_CAP_LAST = CAPACITY_AMOUNT
         TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
         EMP_ID = NVLI(Rs("EMP_ID"), -1)
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
         BALANCE_FLAG = NVLS(Rs("BALANCE_FLAG"), "")
         BALANCE_FLAG_F_I = NVLS(Rs("BALANCE_FLAG_F_I"), "")
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
         PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
       ElseIf Ind = 3 Or Ind = 4 Then
         PALLET_DOC_ID = NVLI(Rs("PALLET_DOC_ID"), -1)
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
         TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
         EMP_ID = NVLI(Rs("EMP_ID"), -1)
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      ElseIf Ind = 5 Then
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
         PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
       ElseIf Ind = 6 Then
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
       ElseIf Ind = 7 Then
         PALLET_DOC_ID = NVLI(Rs("PALLET_DOC_ID"), -1)
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ElseIf Ind = 8 Then
         INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
         MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
      ElseIf Ind = 9 Then
         PALLET_DOC_ID = NVLI(Rs("PALLET_DOC_ID"), -1)
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
         PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
      ElseIf Ind = 10 Then
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
         LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
         PALLET_DOC_NO_OLD = NVLS(Rs("PALLET_DOC_NO"), "")
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
'         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   End If
End Function
Public Function CopyObject(Ind As Long, TempObj As CPalletDoc)
   PALLET_DOC_ID = TempObj.PALLET_DOC_ID                                  '1
   PALLET_DOC_NO = TempObj.PALLET_DOC_NO                                                             '2
   CAPACITY_AMOUNT = TempObj.CAPACITY_AMOUNT                          '3
   PALLET_CAP_LAST = TempObj.PALLET_CAP_LAST
   TRUCK_NO = TempObj.TRUCK_NO                          '4
   EMP_ID = TempObj.EMP_ID                          '5
   LOT_DOC_ID = TempObj.LOT_DOC_ID                          '6
   TX_TYPE = TempObj.TX_TYPE                          '7
   LOT_ID = TempObj.LOT_ID                          '8
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = LOT_DOC_ID
   If LOT_DOC_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   If BALANCE_FLAG = "" Then
      BALANCE_FLAG = "N"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PALLET_DOC_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PALLET_DOC " & _
                     "( " & _
                     "PALLET_DOC_ID, " & _
                     "PALLET_DOC_NO, " & _
                      "CAPACITY_AMOUNT, " & _
                      "TRUCK_NO, " & _
                      "EMP_ID, " & _
                      "LOT_DOC_ID, " & _
                      "TX_TYPE, " & _
                      "BALANCE_FLAG, " & _
                     "CREATE_DATE, " & _
                    "CREATE_BY, " & _
                    "MODIFY_DATE, " & _
                    "MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     "'" & PALLET_DOC_NO & "'," & _
                     CAPACITY_AMOUNT & "," & _
                     "'" & TRUCK_NO & "'," & _
                     EMP_ID & "," & _
                     TempID1 & "," & _
                     "'" & TX_TYPE & "'," & _
                     "'" & BALANCE_FLAG & "'," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
                     PALLET_DOC_ID = ID
   Else 'Show edit
                     SQL1 = "UPDATE PALLET_DOC SET " & _
                     "  PALLET_DOC_NO = '" & PALLET_DOC_NO & "', " & _
                     "  CAPACITY_AMOUNT = " & CAPACITY_AMOUNT & ", " & _
                     "  TRUCK_NO = '" & TRUCK_NO & "', " & _
                     "  EMP_ID = " & EMP_ID & ", " & _
                     "  LOT_DOC_ID = " & TempID1 & ", " & _
                     "  BALANCE_FLAG = '" & BALANCE_FLAG & "', " & _
                     "  MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "  MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "  WHERE PALLET_DOC_ID = " & PALLET_DOC_ID
   End If
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function AddEditBalance()
Dim SQL1 As String
Dim ServerDateTime As String
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE PALLET_DOC SET " & _
   "  BALANCE_FLAG = '" & BALANCE_FLAG & "', " & _
   "  MODIFY_DATE = '" & ServerDateTime & "', " & _
   "  MODIFY_BY = " & glbUser.USER_ID & " " & _
   "  WHERE PALLET_DOC_ID = " & PALLET_DOC_ID

   If (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PALLET_DOC WHERE PALLET_DOC_ID = " & PALLET_DOC_ID
      m_Conn.Execute (SQL1)
End Function
Public Function AddEditPalletDocNo()
Dim SQL1 As String
Dim ServerDateTime As String
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE PALLET_DOC SET " & _
   "  PALLET_DOC_NO = '" & PALLET_DOC_NO & "', " & _
   "  MODIFY_DATE = '" & ServerDateTime & "', " & _
   "  MODIFY_BY = " & glbUser.USER_ID & " " & _
   "  WHERE PALLET_DOC_ID = " & PALLET_DOC_ID

   If (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
