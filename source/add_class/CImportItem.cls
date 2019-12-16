VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImportItem"
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

Public IMPORT_ITEM_ID As Long
Public INVENTORY_DOC_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public ACTUAL_UNIT_PRICE As Double
Public INCLUDE_UNIT_PRICE As Double
Public PREVIOUS_PRICE As Double
Public PREVIOUS_AMOUNT As Double
Public NEW_PRICE As Double
Public TX_AMOUNT As Double
Public NEW_AMOUNT As Double
Public TRANSACTION_SEQ As Long
Public GUI_ID As Long
Public CALCULATE_FLAG As String
Public FATHER_NO As String
Public MOTHER_NO As String
Public TOTAL_WEIGHT As Double
Public TOTAL_ACTUAL_PRICE As Double
Public TOTAL_INCLUDE_PRICE As Double

Public TX_TYPE As String
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long
Public PART_TYPE_NAME As String
Public PART_TYPE_NO As String
Public LOCATION_NO As String
Public LOCATION_NAME As String
Public OLD_TX_AMOUNT As Double
Public OLD_LOCATION_ID As Long
Public PIG_TYPE As String
Public UNIT_NAME As String
Public DOCUMENT_DATE As Date
Public COMMIT_FLAG As String
Public DOCUMENT_TYPE As Long
Public PIG_FLAG As String
Public SUM_ALL_IMPORT_AMT As Double
Public SUM_ALL_IMPORT As Double
Public SALE_FLAG As String
Public DOCUMENT_NO As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DATE1 As Date

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   TX_TYPE = "I"
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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, PT.*, LC.*, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM IMPORT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (UG.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PTT.PART_TYPE_NO, PTT.PART_TYPE_NAME, PI.PART_NO, PI.PART_DESC, IVD.DOCUMENT_NO, IVD.DOCUMENT_DATE, II.TX_AMOUNT, UN.UNIT_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II, INVENTORY_DOC IVD, PART_TYPE PTT, PART_ITEM PI, UNIT UN "
      
      WhereStr = "WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      WhereStr = WhereStr & "AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      WhereStr = WhereStr & "AND (PI.UNIT_COUNT = UN.UNIT_ID) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 1) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT PI.PART_NO, PI.PIG_TYPE, II.LOCATION_ID, SUM(II.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT PI.PART_NO, PI.PIG_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT II.PART_ITEM_ID, SUM(II.TOTAL_INCLUDE_PRICE) SUM_ALL_IMPORT_AMT, SUM(II.TX_AMOUNT) SUM_ALL_IMPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE >= '" & DateToStringIntLow(FROM_DATE) & "')"
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE < '" & DateToStringIntLow(TO_DATE) & "')"
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT II.PART_ITEM_ID, IVD.DOCUMENT_TYPE, IVD.SALE_FLAG, SUM(II.TOTAL_INCLUDE_PRICE) SUM_IMPORT_AMT, SUM(II.TX_AMOUNT) SUM_IMPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
   ElseIf (Ind = 7) Or (Ind = 10) Then
      SelectStr1 = "SELECT UN.UNIT_NAME, II.*, PT.PART_TYPE_NO, PT.PART_TYPE_NAME, IVD.DOCUMENT_TYPE, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, PI.PART_NO, PI.PART_DESC, "
      SelectStr1 = SelectStr1 & "PI.PIG_TYPE, PI.PART_TYPE, II.TX_AMOUNT TX_AMOUNT, II.TOTAL_INCLUDE_PRICE TX_TOTAL_PRICE, "
      SelectStr1 = SelectStr1 & "II.TRANSACTION_SEQ "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE, II.TRANSACTION_SEQ, II.LOCATION_ID, II.PART_ITEM_ID, II.NEW_AMOUNT, II.NEW_PRICE , 'I' "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE II.TRANSACTION_SEQ IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX(II1.TRANSACTION_SEQ) "
      WhereStr = WhereStr & "   FROM IMPORT_ITEM II1 "
      WhereStr = WhereStr & "   LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (II1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "   WHERE (IVD1.DOCUMENT_DATE = IVD.DOCUMENT_DATE) "
      WhereStr = WhereStr & "   AND (II1.PART_ITEM_ID = II.PART_ITEM_ID) "
      WhereStr = WhereStr & "   AND (II1.LOCATION_ID = II.LOCATION_ID) "
      WhereStr = WhereStr & ") "
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT DISTINCT II.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
'      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
      WhereStr = ""
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT II.LOCATION_ID, IVD.DOCUMENT_TYPE, SUM(II.TOTAL_INCLUDE_PRICE) SUM_IMPORT_AMT, SUM(II.TX_AMOUNT) SUM_IMPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT DISTINCT II.LOCATION_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (II.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
      WhereStr = "WHERE (LC.LOCATION_TYPE = 1) "
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT II.PART_ITEM_ID, IVD.DOCUMENT_TYPE, SUM(II.TOTAL_INCLUDE_PRICE) SUM_IMPORT_AMT, SUM(II.TX_AMOUNT) SUM_IMPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT II.PART_ITEM_ID, II.LOCATION_ID, SUM(II.TX_AMOUNT) TX_AMOUNT, SUM(II.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 5) "
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT II.PART_ITEM_ID, SUM(II.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM IMPORT_ITEM II "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (II.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   End If
   
   If IMPORT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IMPORT_ITEM_ID = " & IMPORT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (IMPORT_ITEM_ID = " & IMPORT_ITEM_ID & ") "
      End If
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
'   If Ind = 2 Then
      If PART_TYPE > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
         Else
            WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
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
'   End If
   
   If Trim(PART_NO) <> "" Then
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
   
   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = "GROUP BY PI.PART_NO, II.LOCATION_ID, PI.PIG_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY II.PART_ITEM_ID "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY II.PART_ITEM_ID, IVD.DOCUMENT_TYPE, IVD.SALE_FLAG "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY II.LOCATION_ID, IVD.DOCUMENT_TYPE "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY II.PART_ITEM_ID, IVD.DOCUMENT_TYPE "
   ElseIf Ind = 14 Then
      GroupByStr = "GROUP BY II.PART_ITEM_ID,  II.LOCATION_ID "
   ElseIf Ind = 15 Then
      GroupByStr = "GROUP BY II.PART_ITEM_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IMPORT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IMPORT_ITEM_ID "
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
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, UG.TRANSACTION_SEQ  "
      Else
         OrderByStr = "ORDER BY IMPORT_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY IMPORT_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PI.PART_NO, PI.PIG_TYPE "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 7 Then
      If OrderType = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, II.TRANSACTION_SEQ ASC "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO DESC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, II.TRANSACTION_SEQ ASC "
      End If
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY II.PART_ITEM_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY PART_NO "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY LOCATION_ID "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY LOCATION_ID "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 14 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 15 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   End If
   
   If Ind <> 7 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      IMPORT_ITEM_ID = NVLI(Rs("IMPORT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      INCLUDE_UNIT_PRICE = NVLD(Rs("INCLUDE_UNIT_PRICE"), 0)
      PREVIOUS_PRICE = NVLD(Rs("PREVIOUS_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      FATHER_NO = NVLS(Rs("FATHER_NO"), "")
      MOTHER_NO = NVLS(Rs("MOTHER_NO"), "")
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      
      OLD_TX_AMOUNT = TX_AMOUNT
      OLD_LOCATION_ID = LOCATION_ID
      
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      TX_TYPE = "I" 'Import
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 3 Then
      PART_NO = NVLS(Rs("PART_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      SUM_ALL_IMPORT_AMT = NVLD(Rs("SUM_ALL_IMPORT_AMT"), 0)
      SUM_ALL_IMPORT = NVLD(Rs("SUM_ALL_IMPORT"), 0)
   ElseIf Ind = 6 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
      SUM_ALL_IMPORT_AMT = NVLD(Rs("SUM_IMPORT_AMT"), 0)
      SUM_ALL_IMPORT = NVLD(Rs("SUM_IMPORT"), 0)
   ElseIf (Ind = 7) Or (Ind = 10) Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TX_TOTAL_PRICE"), 0)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
   ElseIf Ind = 8 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 9 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   ElseIf Ind = 11 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      SUM_ALL_IMPORT_AMT = NVLD(Rs("SUM_IMPORT_AMT"), 0)
      SUM_ALL_IMPORT = NVLD(Rs("SUM_IMPORT"), 0)
   ElseIf Ind = 12 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
   ElseIf Ind = 13 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      SUM_ALL_IMPORT_AMT = NVLD(Rs("SUM_IMPORT_AMT"), 0)
      SUM_ALL_IMPORT = NVLD(Rs("SUM_IMPORT"), 0)
   ElseIf Ind = 14 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
   ElseIf Ind = 15 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("IMPORT_ITEM_SEQ", ID, glbErrorLog)
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "INSERT INTO IMPORT_ITEM " & _
                     "( " & _
                     "   IMPORT_ITEM_ID, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   ACTUAL_UNIT_PRICE, " & _
                     "   INCLUDE_UNIT_PRICE, " & _
                     "   PREVIOUS_PRICE, " & _
                     "   PREVIOUS_AMOUNT, " & _
                     "   NEW_PRICE, " & _
                     "   TX_AMOUNT, " & _
                     "   NEW_AMOUNT, "
            SQL1 = SQL1 & _
                     "   TRANSACTION_SEQ, " & _
                     "   GUI_ID, " & _
                     "   CALCULATE_FLAG, " & _
                     "   FATHER_NO, " & _
                     "   MOTHER_NO, " & _
                     "   TOTAL_WEIGHT, " & _
                     "   TOTAL_ACTUAL_PRICE, " & _
                     "   TOTAL_INCLUDE_PRICE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & INVENTORY_DOC_ID & "," & _
                         "" & PART_ITEM_ID & "," & _
                         "" & LOCATION_ID & "," & _
                         "" & ACTUAL_UNIT_PRICE & "," & _
                         "" & INCLUDE_UNIT_PRICE & "," & _
                         "" & PREVIOUS_PRICE & "," & _
                         "" & PREVIOUS_AMOUNT & "," & _
                         "" & NEW_PRICE & ","
               SQL1 = SQL1 & _
                         "" & TX_AMOUNT & "," & _
                         "" & NEW_AMOUNT & "," & _
                         "" & ID2 & "," & _
                         "" & GUI_ID & "," & _
                         "'" & CALCULATE_FLAG & "'," & _
                         "'" & FATHER_NO & "'," & _
                         "'" & MOTHER_NO & "'," & _
                         "" & TOTAL_WEIGHT & "," & _
                         "" & TOTAL_ACTUAL_PRICE & "," & _
                         "" & TOTAL_INCLUDE_PRICE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         IMPORT_ITEM_ID = ID
   Else 'Show edit
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "UPDATE IMPORT_ITEM SET " & _
                     "   INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   LOCATION_ID = " & LOCATION_ID & ", " & _
                     "   ACTUAL_UNIT_PRICE = " & ACTUAL_UNIT_PRICE & ", " & _
                     "   INCLUDE_UNIT_PRICE = " & INCLUDE_UNIT_PRICE & ", " & _
                     "   PREVIOUS_PRICE = " & PREVIOUS_PRICE & ", " & _
                     "   PREVIOUS_AMOUNT = " & PREVIOUS_AMOUNT & ", " & _
                     "   NEW_PRICE = " & NEW_PRICE & ", " & _
                     "   TX_AMOUNT = " & TX_AMOUNT & ", " & _
                     "   NEW_AMOUNT = " & NEW_AMOUNT & ", " & _
                     "   TRANSACTION_SEQ = " & ID2 & ", " & _
                     "   GUI_ID = " & GUI_ID & ", " & _
                     "   CALCULATE_FLAG = '" & CALCULATE_FLAG & "', " & _
                     "   FATHER_NO = '" & FATHER_NO & "', " & _
                     "   MOTHER_NO = '" & MOTHER_NO & "', " & _
                     "   TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   TOTAL_ACTUAL_PRICE = " & TOTAL_ACTUAL_PRICE & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE IMPORT_ITEM_ID = " & IMPORT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM IMPORT_ITEM WHERE IMPORT_ITEM_ID = " & IMPORT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

