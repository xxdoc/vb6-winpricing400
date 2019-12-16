VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLotDoc"
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
Public LOT_DOC_ID As Long
Public LOT_DOC_ID_REF As Long
Public LOT_ID As Long
Public LOT_DATE As Date
Public LOT_NO As String
Public BIN_NO As Long
Public BIN_NAME As String
Public LOCK_NO As Long
Public LOCK_NAME As String
Public HEAD_PACK_NO As Long
Public HEAD_PACK_NAME As String
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public LOCATION_NAME As String
Public PART_NO As String
Public LOT_ITEM_WH_ID As Long
Public FROM_DATE As Date
Public TO_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_TYPE As Long
Public DOCUMENT_TYPE_SET As String
Public PRODUCT_TYPE_ID As Long
Public START_DATE As Date
Public BL_START_DATE As Date
Public PACK_DATE As Date
Public TIME_PACK_BEGIN As Date
Public TIME_PACK_END As Date
Public LOT_AMOUNT As Double
Public LOT_PAYOFF As Double
Public LOT_BAL As Double
Public TX_TYPE As String
Public PALLET_DOC_NO As String
Public CAPACITY_AMOUNT As Double
Public C_PalletDoc  As Collection
Public BALANCE_FLAG As String
Public OUT_STOCK_FLAG As String
Public IS_BALANCE As Long
Public IS_EDIT_LOT As Boolean
Public INVENTORY_WH_DOC_ID As Long
Public VERIFY_FLAG As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set C_PalletDoc = New Collection
   DOCUMENT_DATE = -1
   START_DATE = -1
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

   If Ind = 1 Then 'ใช้แล้ว   แก้อีกแล้ว
      SelectStr1 = "SELECT LTD.*,LT.LOT_NO,LIW.START_DATE  " '
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
    ElseIf Ind = 2 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LT.LOT_NO,LT.LOT_DATE,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LTD.LOT_ID,LTD.BALANCE_FLAG,LIW.BIN_NO,LIW.LOCK_NO,LIW.HEAD_PACK_NO,LIW.BL_START_DATE,LIW.TX_TYPE,LIW.PACK_DATE,LIW.TIME_PACK_BEGIN,LIW.TIME_PACK_END,LIW.LOT_ITEM_WH_ID,IWD.START_DATE,IWD.DOCUMENT_TYPE,IWD.INVENTORY_WH_DOC_ID,LC.LOCATION_NAME,LC2.LOCATION_NAME BIN_NAME,LC3.LOCATION_NAME LOCK_NAME,LC4.LOCATION_NAME HEAD_PACK_NAME, "
        SelectStr1 = SelectStr1 & "LIW.LOCK_NO,LIW.PRODUCT_TYPE_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID = IWD.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC  ON (LIW.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LIW.BIN_NO = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC3 ON (LIW.LOCK_NO = LC3.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC4 ON (LIW.HEAD_PACK_NO = LC4.LOCATION_ID) "

    ElseIf Ind = 3 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LT.LOT_NO,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LTD.LOT_ID,LTD.LOT_ITEM_WH_ID,LIW.PART_ITEM_ID,LIW.LOCATION_ID,LIW.TX_TYPE,IWD.START_DATE,IWD.INVENTORY_WH_DOC_ID,Sum (PD.CAPACITY_AMOUNT)  LOT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN PALLET_DOC PD ON (LTD.LOT_DOC_ID = PD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 4 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LT.LOT_NO,LT.LOT_DATE,LTD.LOT_DOC_ID,LTD.BALANCE_FLAG,LTD.LOT_DOC_ID_REF,LTD2.LOT_ID,LTD2.LOT_ITEM_WH_ID,LIW2.BIN_NO,LIW2.LOCK_NO,LIW2.HEAD_PACK_NO,LIW2.TX_TYPE,LIW2.PACK_DATE,LIW2.TIME_PACK_BEGIN,LIW2.TIME_PACK_END,LIW2.LOT_ITEM_WH_ID,LIW2.BL_START_DATE,IWD.START_DATE,IWD.DOCUMENT_TYPE,LC.LOCATION_NAME,LC2.LOCATION_NAME BIN_NAME,LC3.LOCATION_NAME LOCK_NAME,LC4.LOCATION_NAME HEAD_PACK_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD2 ON (LTD.LOT_DOC_ID_REF = LTD2.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD2.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW2 ON (LTD2.LOT_ITEM_WH_ID =LIW2.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW2.INVENTORY_WH_DOC_ID = IWD.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC  ON (LIW2.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LIW2.BIN_NO = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC3 ON (LIW2.LOCK_NO = LC3.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC4 ON (LIW2.HEAD_PACK_NO = LC4.LOCATION_ID) "
   ElseIf Ind = 5 Then 'ใช้แล้ว
      SelectStr1 = "SELECT DISTINCT LTD.LOT_ID,LT.LOT_NO "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LIW.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID = IWD.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LIW.BIN_NO = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC3 ON (LIW.LOCK_NO = LC3.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC4 ON (LIW.HEAD_PACK_NO = LC4.LOCATION_ID) "
    ElseIf Ind = 6 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LTD.*,LT.LOT_NO,IWD.DOCUMENT_NO,IWD.DOCUMENT_TYPE,IWD.DOCUMENT_DATE,IWD.START_DATE,LIW.BL_START_DATE,LIW.PACK_DATE,LIW.TIME_PACK_BEGIN,LIW.TIME_PACK_END,LIW.BIN_NO,LIW.LOCK_NO,LIW.HEAD_PACK_NO,LC2.LOCATION_NAME BIN_NAME,LC3.LOCATION_NAME LOCK_NAME,LC4.LOCATION_NAME HEAD_PACK_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID = IWD.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LIW.BIN_NO = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC3 ON (LIW.LOCK_NO = LC3.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC4 ON (LIW.HEAD_PACK_NO = LC4.LOCATION_ID) "
   ElseIf Ind = 7 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LTD.* "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
    ElseIf Ind = 8 Then  'ใช้แล้ว
      SelectStr1 = "SELECT LTD.LOT_DOC_ID "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID = IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 9 Then  'ใช้แล้ว
      SelectStr1 = "SELECT LTD.LOT_DOC_ID,LT.LOT_NO,LTD.LOT_DOC_ID_REF,LC3.LOCATION_NAME LOCK_NAME,IWD2.START_DATE,LIW2.BL_START_DATE,IWD2.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD2 ON (LTD.LOT_DOC_ID_REF = LTD2.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW2 ON (LTD2.LOT_ITEM_WH_ID = LIW2.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC3 ON (LIW2.LOCK_NO = LC3.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD2 ON (LIW2.INVENTORY_WH_DOC_ID =IWD2.INVENTORY_WH_DOC_ID)"
   ElseIf Ind = 10 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LT.LOT_NO,LTD.LOT_DOC_ID_REF,LTD.LOT_ID,LIW.TX_TYPE,Sum (PD.CAPACITY_AMOUNT)  LOT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN PALLET_DOC PD ON (LTD.LOT_DOC_ID = PD.LOT_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
   ElseIf Ind = 11 Then 'ใช้แล้ว
      SelectStr1 = "SELECT LT.LOT_NO,PD.PALLET_DOC_NO,PD.TX_TYPE,SUM(PD.CAPACITY_AMOUNT) CAPACITY_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PALLET_DOC PD ON (LTD.LOT_DOC_ID = PD.LOT_DOC_ID) "

      WhereStr = "WHERE PD.PALLET_DOC_ID IS NOT NULL "
   ElseIf Ind = 12 Then 'loadfifo
      SelectStr1 = "SELECT LT.LOT_NO,LT.LOT_DATE,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LTD.LOT_ID,LTD.BALANCE_FLAG,LIW.BIN_NO,LIW.LOCK_NO,LIW.HEAD_PACK_NO,LIW.BL_START_DATE,LIW.TX_TYPE,LIW.PACK_DATE,LIW.TIME_PACK_BEGIN,LIW.TIME_PACK_END,LIW.LOT_ITEM_WH_ID,IWD.START_DATE,IWD.DOCUMENT_TYPE,IWD.INVENTORY_WH_DOC_ID,LC.LOCATION_NAME,LC2.LOCATION_NAME BIN_NAME,LC3.LOCATION_NAME LOCK_NAME,LC4.LOCATION_NAME HEAD_PACK_NAME, "
      SelectStr1 = SelectStr1 & "LIW.LOCK_NO,LIW.PRODUCT_TYPE_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      WhereStr = ""
      FromStr = "FROM LOT_DOC LTD "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID =LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LIW ON (LTD.LOT_ITEM_WH_ID =LIW.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN (SELECT DISTINCT IVD2.INVENTORY_WH_DOC_ID,IVD2.START_DATE,IVD2.DOCUMENT_TYPE,IVD2.DOCUMENT_DATE,J.VERIFY_FLAG FROM INVENTORY_WH_DOC IVD2 "
      FromStr = FromStr & "INNER JOIN JOB J ON (IVD2.INVENTORY_WH_DOC_ID =J.INVENTORY_WH_DOC_ID)) AS IWD ON (LIW.INVENTORY_WH_DOC_ID =IWD.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC  ON (LIW.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LIW.BIN_NO = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC3 ON (LIW.LOCK_NO = LC3.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC4 ON (LIW.HEAD_PACK_NO = LC4.LOCATION_ID) "
   End If
   
   If LOT_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.LOT_DOC_ID = " & LOT_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LTD.LOT_DOC_ID = " & LOT_DOC_ID & ") "
      End If
   End If
   
    
   
   If IS_BALANCE = 1 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.LOT_DOC_ID <> " & LOT_DOC_ID_REF & ") "
      Else
         WhereStr = WhereStr & "AND (LTD.LOT_DOC_ID <> " & LOT_DOC_ID_REF & ") "
      End If
   Else
      If LOT_DOC_ID_REF > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (LTD.LOT_DOC_ID_REF = " & LOT_DOC_ID_REF & ") "
         Else
            WhereStr = WhereStr & "AND (LTD.LOT_DOC_ID_REF = " & LOT_DOC_ID_REF & ") "
         End If
      End If
   End If
   
   If LOT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.LOT_ID = " & LOT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LTD.LOT_ID = " & LOT_ID & ") "
      End If
   End If

   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LIW.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Len(OUT_STOCK_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.OUT_STOCK_FLAG = '" & OUT_STOCK_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (LIW.OUT_STOCK_FLAG = '" & OUT_STOCK_FLAG & "') "
      End If
   End If
   
   If Len(VERIFY_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IWD.VERIFY_FLAG = '" & VERIFY_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (IWD.VERIFY_FLAG = '" & VERIFY_FLAG & "') "
      End If
   End If

   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_NO = '" & PART_NO & "') "
      Else
         WhereStr = WhereStr & "AND (PI.PART_NO = '" & PART_NO & "') "
      End If
   End If
   
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LIW.TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (LIW.TX_TYPE = '" & TX_TYPE & "') "
      End If
   End If
   
   If LOT_ITEM_WH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTD.LOT_ITEM_WH_ID = " & LOT_ITEM_WH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LTD.LOT_ITEM_WH_ID = " & LOT_ITEM_WH_ID & ") "
      End If
   End If
   
    If Len(BALANCE_FLAG) > 0 Then
       If Len(WhereStr) <= 0 Then
          WhereStr = "WHERE (LTD.BALANCE_FLAG = '" & BALANCE_FLAG & "') "
       Else
          WhereStr = WhereStr & "AND (LTD.BALANCE_FLAG = '" & BALANCE_FLAG & "') "
       End If
    End If
   
  If Ind = 8 Then
      If DOCUMENT_DATE > -1 Then
         TempDate = DateToStringInt(Trim(DOCUMENT_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (IWD.DOCUMENT_DATE < '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (IWD.DOCUMENT_DATE  < '" & ChangeQuote(TempDate) & "')"
         End If
      End If
      
   End If
   
   If TX_TYPE = "I" Then
      If FROM_DATE > -1 Then
         TempDate = DateToStringIntLow(Trim(FROM_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (LIW.PACK_DATE >= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (LIW.PACK_DATE  >= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
   
      If TO_DATE > -1 Then
         TempDate = DateToStringIntHi(Trim(TO_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (LIW.PACK_DATE   <= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (LIW.PACK_DATE  <= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
ElseIf TX_TYPE = "E" Then
       If FROM_DATE > -1 Then
         TempDate = DateToStringIntLow(Trim(FROM_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (IWD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (IWD.DOCUMENT_DATE  >= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
   
      If TO_DATE > -1 Then
         TempDate = DateToStringIntHi(Trim(TO_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (IWD.DOCUMENT_DATE   <= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (IWD.DOCUMENT_DATE  <= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
Else
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LTD.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LTD.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LTD.CREATE_DATE  <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LTD.CREATE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
End If
   
'   If FROM_DATE > -1 Then
'      TempDate = DateToStringIntLow(Trim(FROM_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (LTD.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (LTD.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
'
'   If TO_DATE > -1 Then
'      TempDate = DateToStringIntHi(Trim(TO_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (LTD.CREATE_DATE  <= '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (LTD.CREATE_DATE <= '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
   
   GroupByStr = ""
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 3 Then
      GroupByStr = "GROUP BY LT.LOT_NO,LTD.LOT_DOC_ID,LTD.LOT_DOC_ID_REF,LTD.LOT_ID,LTD.LOT_ITEM_WH_ID,LIW.PART_ITEM_ID,LIW.LOCATION_ID,LIW.TX_TYPE,IWD.START_DATE,IWD.INVENTORY_WH_DOC_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY LT.LOT_NO,LTD.LOT_DOC_ID_REF,LTD.LOT_ID,LIW.TX_TYPE  "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY LT.LOT_NO,PD.PALLET_DOC_NO,PD.TX_TYPE  "
   End If

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LTD.LOT_DOC_ID "
      Else
         OrderByStr = "ORDER BY LTD.LOT_DOC_ID "
      End If
   ElseIf Ind = 2 Or Ind = 5 Or Ind = 12 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LTD.LOT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LIW.PACK_DATE ASC,TIME_PACK_BEGIN "
      End If
   ElseIf Ind = 10 Then
      OrderByStr = ""
   ElseIf Ind = 11 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PD.PALLET_DOC_NO "
      End If
   Else
      OrderByStr = "ORDER BY LTD.LOT_DOC_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Or Ind = 5 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
  If Ind = 10 Then
      OrderTypeStr = ""
 End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
         PACK_DATE = InternalDateToDate(NVLS(Rs("PACK_DATE"), ""))
         TIME_PACK_BEGIN = InternalTimeToTime(NVLS(Rs("TIME_PACK_BEGIN"), ""))
         TIME_PACK_END = InternalTimeToTime(NVLS(Rs("TIME_PACK_END"), ""))
    ElseIf Ind = 2 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         BALANCE_FLAG = NVLS(Rs("BALANCE_FLAG"), "")
         LOCK_NO = NVLI(Rs("LOCK_NO"), -1)
         LOCK_NAME = NVLS(Rs("LOCK_NAME"), "")
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
         BIN_NAME = NVLS(Rs("BIN_NAME"), "")
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
         HEAD_PACK_NAME = NVLS(Rs("HEAD_PACK_NAME"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
         BL_START_DATE = InternalDateToDate(NVLS(Rs("BL_START_DATE"), ""))
         PACK_DATE = InternalDateToDate(NVLS(Rs("PACK_DATE"), ""))
         TIME_PACK_BEGIN = InternalTimeToTime(NVLS(Rs("TIME_PACK_BEGIN"), ""))
         TIME_PACK_END = InternalTimeToTime(NVLS(Rs("TIME_PACK_END"), ""))
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
         PRODUCT_TYPE_ID = NVLI(Rs("PRODUCT_TYPE_ID"), -1)
         LOCK_NO = NVLI(Rs("LOCK_NO"), -1)
         LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
         INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
   ElseIf Ind = 3 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
         LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
         LOT_AMOUNT = NVLD(Rs("LOT_AMOUNT"), 0)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
   ElseIf Ind = 4 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_DATE = InternalDateToDate(NVLS(Rs("LOT_DATE"), ""))
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         BALANCE_FLAG = NVLS(Rs("BALANCE_FLAG"), "")
         LOCK_NO = NVLI(Rs("LOCK_NO"), -1)
         LOCK_NAME = NVLS(Rs("LOCK_NAME"), "")
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
         BIN_NAME = NVLS(Rs("BIN_NAME"), "")
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
         HEAD_PACK_NAME = NVLS(Rs("HEAD_PACK_NAME"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
         PACK_DATE = InternalDateToDate(NVLS(Rs("PACK_DATE"), ""))
         TIME_PACK_BEGIN = InternalTimeToTime(NVLS(Rs("TIME_PACK_BEGIN"), ""))
         TIME_PACK_END = InternalTimeToTime(NVLS(Rs("TIME_PACK_END"), ""))
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         BL_START_DATE = InternalDateToDate(NVLS(Rs("BL_START_DATE"), ""))
         DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
         LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      ElseIf Ind = 5 Then
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
       ElseIf Ind = 6 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
         BL_START_DATE = InternalDateToDate(NVLS(Rs("BL_START_DATE"), ""))
         PACK_DATE = InternalDateToDate(NVLS(Rs("PACK_DATE"), ""))
         TIME_PACK_BEGIN = InternalTimeToTime(NVLS(Rs("TIME_PACK_BEGIN"), ""))
         TIME_PACK_END = InternalTimeToTime(NVLS(Rs("TIME_PACK_END"), ""))
         LOCK_NO = NVLI(Rs("LOCK_NO"), -1)
         LOCK_NAME = NVLS(Rs("LOCK_NAME"), "")
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
         BIN_NAME = NVLS(Rs("BIN_NAME"), "")
         HEAD_PACK_NO = NVLI(Rs("HEAD_PACK_NO"), -1)
         HEAD_PACK_NAME = NVLS(Rs("HEAD_PACK_NAME"), "")
         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
         DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
         DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      ElseIf Ind = 7 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_ITEM_WH_ID = NVLI(Rs("LOT_ITEM_WH_ID"), -1)
         BIN_NO = NVLI(Rs("BIN_NO"), -1)
      ElseIf Ind = 8 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
      ElseIf Ind = 9 Then
         LOT_DOC_ID = NVLI(Rs("LOT_DOC_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
         LOCK_NAME = NVLS(Rs("LOCK_NAME"), "")
         START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
         BL_START_DATE = InternalDateToDate(NVLS(Rs("BL_START_DATE"), ""))
         DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      ElseIf Ind = 10 Then
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_AMOUNT = NVLD(Rs("LOT_AMOUNT"), 0)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
         LOT_DOC_ID_REF = NVLI(Rs("LOT_DOC_ID_REF"), -1)
     ElseIf Ind = 11 Then
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         PALLET_DOC_NO = NVLS(Rs("PALLET_DOC_NO"), "")
         CAPACITY_AMOUNT = NVLD(Rs("CAPACITY_AMOUNT"), 0)
         TX_TYPE = NVLS(Rs("TX_TYPE"), "")
'         DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
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

   TempID1 = LOT_DOC_ID_REF
   If LOT_DOC_ID_REF <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = BALANCE_FLAG
   If BALANCE_FLAG = "" Then
      TempID2 = "N"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LOT_DOC_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO LOT_DOC " & _
                     "( " & _
                     "LOT_DOC_ID, " & _
                     "LOT_ID, " & _
                     "LOT_ITEM_WH_ID, " & _
                     "LOT_DOC_ID_REF, " & _
                     "BIN_NO, " & _
                     "BALANCE_FLAG, " & _
                     "CREATE_DATE, " & _
                    "CREATE_BY, " & _
                    "MODIFY_DATE, " & _
                    "MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     LOT_ID & "," & _
                     LOT_ITEM_WH_ID & "," & _
                     TempID1 & "," & _
                     BIN_NO & "," & _
                     "'" & TempID2 & "'," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
                     LOT_DOC_ID = ID
   Else 'Show edit
                     SQL1 = "UPDATE LOT_DOC SET " & _
                     "  LOT_ID = " & LOT_ID & ", " & _
                     "  LOT_ITEM_WH_ID = " & LOT_ITEM_WH_ID & ", " & _
                     "  LOT_DOC_ID_REF = " & TempID1 & ", " & _
                     "  BIN_NO = " & BIN_NO & ", " & _
                     "  MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "  MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "  WHERE LOT_DOC_ID = " & LOT_DOC_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function AddEditBalance()
Dim SQL1 As String
Dim ServerDateTime As String
Dim TempID1 As String
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

TempID1 = LOT_DOC_ID_REF
If LOT_DOC_ID_REF <= 0 Then
   TempID1 = "NULL "
End If
   
   SQL1 = "UPDATE LOT_DOC SET " & _
   "  LOT_DOC_ID_REF = " & TempID1 & ", " & _
   "  BALANCE_FLAG = '" & BALANCE_FLAG & "', " & _
   "  MODIFY_DATE = '" & ServerDateTime & "', " & _
   "  MODIFY_BY = " & glbUser.USER_ID & " " & _
   "  WHERE LOT_DOC_ID = " & LOT_DOC_ID

   If (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function AddEditLot()
Dim SQL1 As String
Dim ServerDateTime As String
Dim TempID1 As String
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE LOT_DOC SET " & _
   "  LOT_ID = " & LOT_ID & ", " & _
   "  MODIFY_DATE = '" & ServerDateTime & "', " & _
   "  MODIFY_BY = " & glbUser.USER_ID & " " & _
   "  WHERE LOT_DOC_ID_REF = " & LOT_DOC_ID

   If (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function AddEditBalanceReFlag()
Dim SQL1 As String
Dim ServerDateTime As String
Dim TempID1 As String
Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

TempID1 = "NULL "
   
   SQL1 = "UPDATE LOT_DOC SET " & _
   "  LOT_DOC_ID_REF = " & TempID1 & ", " & _
   "  BALANCE_FLAG = '" & BALANCE_FLAG & "', " & _
   "  MODIFY_DATE = '" & ServerDateTime & "', " & _
   "  MODIFY_BY = " & glbUser.USER_ID & " " & _
   "  WHERE LOT_DOC_ID_REF = " & LOT_DOC_ID & _
   " AND BALANCE_FLAG = 'Y'"

   If (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PALLET_DOC WHERE LOT_DOC_ID = " & LOT_DOC_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_DOC WHERE LOT_DOC_ID = " & LOT_DOC_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set C_PalletDoc = Nothing
End Sub


