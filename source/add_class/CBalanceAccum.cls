VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalanceAccum"
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

Public BALANCE_ACCUM_ID As Long
Public DOCUMENT_DATE As Date
Public PART_ITEM_ID As Long
Public IMPORT_AMOUNT As Double
Public EXPORT_AMOUNT As Double
Public BALANCE_AMOUNT As Double
Public TOTAL_INCLUDE_PRICE As Double
Public LOCATION_ID As Long
Public AVG_PRICE As Double

Public FROM_DATE As Date
Public TO_DATE As Date
Public TO_DATE1 As Date
Public MASTER_FLAG As String
Public PIG_FLAG As String
Public LOCATION_TYPE As Long
Public PART_TYPE As Long
 Public LOCATION_NO As String
 Public LOCATION_NAME As String
 Public PART_NO As String
 Public PART_DESC As String
 Public PART_TYPE_ID As Long
 Public PART_GROUP_ID As Long
  Public NEW_ACCOUNT_ID As Long
  
  Public PART_TYPE_NAME As String
  Public PART_TYPE_NO As String
  
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
Dim GroupByStr As String
Dim MustCount As Boolean

   If Ind = 1 Then
      SelectStr1 = "SELECT CM.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM CM "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SUM(IMPORT_AMOUNT - EXPORT_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM CM "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT PART_ITEM_ID, SUM(IMPORT_AMOUNT-EXPORT_AMOUNT) BALANCE_AMOUNT, SUM(TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM CM "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT PART_ITEM_ID, LOCATION_ID, SUM(IMPORT_AMOUNT - EXPORT_AMOUNT) BALANCE_AMOUNT  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM CM "
      
      WhereStr = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(AVG_PRICE) AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
     WhereStr = "WHERE BALANCE_ACCUM_ID IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX(BALANCE_ACCUM_ID) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & "   AND (BA2.BALANCE_AMOUNT >= 0) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT BA1.BALANCE_ACCUM_ID, BA1.DOCUMENT_DATE, BA1.LOCATION_ID, BA1.PART_ITEM_ID, BA1.BALANCE_AMOUNT, BA1.TOTAL_INCLUDE_PRICE, BA1.AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      
     WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & ") "
     
     MustCount = False
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT BA1.DOCUMENT_DATE, BA1.LOCATION_ID, BA1.PART_ITEM_ID, BA1.AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "

     WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT DISTINCT CM.LOCATION_ID, CM.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PT.PART_TYPE_ID = PI.PART_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.LOCATION_ID = LC.LOCATION_ID) "
      
     WhereStr = ""
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT BA1.LOCATION_ID, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(AVG_PRICE) AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
     WhereStr = "WHERE BALANCE_ACCUM_ID IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX(BALANCE_ACCUM_ID) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & "   AND (BA2.BALANCE_AMOUNT >= 0) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(AVG_PRICE) AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
     WhereStr = "WHERE BALANCE_ACCUM_ID IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX(BALANCE_ACCUM_ID) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & "   AND (BA2.BALANCE_AMOUNT >= 0) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT DISTINCT CM.LOCATION_ID, LC.LOCATION_NO, LC.LOCATION_NAME, CM.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",PT.PART_TYPE_NO,PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.LOCATION_ID = LC.LOCATION_ID) "
      
     WhereStr = ""
   End If
   
   If BALANCE_ACCUM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID & ") "
      End If
   End If
      
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   
   If LOCATION_ID > 0 Then
      If Ind = 1 Or Ind = 2 Or Ind = 3 Or Ind = 4 Or Ind = 8 Or Ind = 11 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (CM.LOCATION_ID = " & LOCATION_ID & ") "
         Else
            WhereStr = WhereStr & "AND (CM.LOCATION_ID = " & LOCATION_ID & ") "
         End If
      ElseIf Ind = 5 Or Ind = 6 Or Ind = 7 Or Ind = 9 Or Ind = 10 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE BA1.LOCATION_ID = " & LOCATION_ID & ") "
         Else
            WhereStr = WhereStr & "AND (BA1.LOCATION_ID = " & LOCATION_ID & ") "
         End If
      Else
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
         Else
            WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
         End If
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
         WhereStr = "WHERE (PART_TYPE_ID = " & PART_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE_ID = " & PART_TYPE_ID & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
      
   If LOCATION_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_TYPE = " & LOCATION_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_TYPE = " & LOCATION_TYPE & ") "
      End If
   End If
      
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If Ind = 3 Then
      GroupByStr = "GROUP BY CM.PART_ITEM_ID "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY PART_ITEM_ID, LOCATION_ID "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY PART_ITEM_ID "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY LOCATION_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY PART_ITEM_ID "
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY BALANCE_ACCUM_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY BALANCE_ACCUM_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 6 Then
      'OrderByStr = "ORDER BY BALANCE_ACCUM_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY BALANCE_ACCUM_ID "
   ElseIf Ind = 8 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOCATION_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LOCATION_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PART_DESC "
      Else
         OrderByStr = "ORDER BY LOCATION_NO "
      End If
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY LOCATION_ID "
   ElseIf Ind = 10 Then
      'OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   End If
   
   If Ind <> 6 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   If OrderByStr = "" Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BALANCE_ACCUM_ID = NVLI(Rs("BALANCE_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 2 Then
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 6 Then
      BALANCE_ACCUM_ID = NVLI(Rs("BALANCE_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
'      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "")
   ElseIf Ind = 7 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 8 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
   ElseIf Ind = 9 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 10 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 11 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   End If
End Function

Public Function AddEditData()
On Error Resume Next
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

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BALANCE_ACCUM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BALANCE_ACCUM " & _
                     "( " & _
                     "   BALANCE_ACCUM_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   IMPORT_AMOUNT, " & _
                     "   EXPORT_AMOUNT, " & _
                     "   BALANCE_AMOUNT, " & _
                     "   TOTAL_INCLUDE_PRICE, " & _
                     "   LOCATION_ID, " & _
                     "   AVG_PRICE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "'," & _
                        PART_ITEM_ID & "," & _
                        IMPORT_AMOUNT & "," & _
                        EXPORT_AMOUNT & "," & _
                        BALANCE_AMOUNT & "," & _
                        TOTAL_INCLUDE_PRICE & "," & _
                        LOCATION_ID & "," & _
                        AVG_PRICE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BALANCE_ACCUM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_ACCUM SET " & _
                     "    DOCUMENT_DATE = '" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "', " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    IMPORT_AMOUNT = " & IMPORT_AMOUNT & ", " & _
                     "    EXPORT_AMOUNT = " & EXPORT_AMOUNT & ", " & _
                     "    BALANCE_AMOUNT = " & BALANCE_AMOUNT & ", " & _
                     "    TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "    LOCATION_ID = " & LOCATION_ID & ", " & _
                     "    AVG_PRICE = " & AVG_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function UpdatePriceAmount()
Dim ServerDateTime As String
Dim SQL1 As String
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_ACCUM SET " & _
                     "    BALANCE_AMOUNT = " & BALANCE_AMOUNT & ", " & _
                     "    TOTAL_INCLUDE_PRICE = " & BALANCE_AMOUNT * AVG_PRICE & ", " & _
                     "    AVG_PRICE = " & AVG_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM BALANCE_ACCUM WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String

   SQL1 = "DELETE FROM BALANCE_ACCUM "
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   SQL1 = SQL1 & WhereStr
   m_Conn.Execute (SQL1)
End Function


