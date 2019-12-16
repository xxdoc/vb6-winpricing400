VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobExt"
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

Public JOB_ID As Long
Public JOB_NO As String
Public JOB_DATE As Date
Public JOB_DESC As String
Public PROCESS_ID As Long
Public BATCH_NO As String
Public APPROVED_BY As Long
Public RESPONSE_BY As Long
Public START_DATE  As Date
Public FINISH_DATE As Date
Public COMMIT_FLAG As String
Public INVENTORY_DOC_ID As Long
Public JOB_DOC_TYPE As Long
Public PART_ITEM_ID As Long
Public STD_AMOUNT As Double
Public ACTUAL_AMOUNT As Double
Public FORMULA_ID As Long

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As Long

Public PARCEL_TYPE As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public LONG_NAMEA As String
Public LAST_NAMEA As String
Public LONG_NAMER As String
Public LAST_NAMER As String
Public DOC_NO As String
Public PROCESS_NAME As String
Public OLD_COMMIT_FLAG As String
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long
Public ProcessSet As String

Public Inputs As Collection
Public Outputs As Collection
Public Peoples As Collection
Public Machines As Collection
Public Parameters As Collection
Public Verifies As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   START_DATE = -1
   FINISH_DATE = -1
   COMMIT_FLAG = ""
    JOB_DATE = -1
    
   FROM_DATE = -1
   TO_DATE = -1
   
   Set Inputs = New Collection
   Set Outputs = New Collection
   Set Peoples = New Collection
   Set Machines = New Collection
   Set Parameters = New Collection
   Set Verifies = New Collection
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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT J.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB_EXT J "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(J.ACTUAL_AMOUNT) ACTUAL_AMOUNT, SUM(J.STD_AMOUNT) STD_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (J.PART_ITEM_ID = PI.PART_ITEM_ID) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(J.ACTUAL_AMOUNT) ACTUAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",PI.PARCEL_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (J.PART_ITEM_ID = PI.PART_ITEM_ID) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (J.PART_ITEM_ID = PI.PART_ITEM_ID) "
   End If
   
   If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_ID = " & JOB_ID & ") "
      End If
   End If
 
    If Trim(JOB_NO) <> "" Then
      If WildCard(JOB_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         End If
      End If
   End If
 
   If Trim(JOB_DESC) <> "" Then
      If WildCard(JOB_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(JOB_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(JOB_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "')"
         End If
      End If
   End If
     
   If JOB_DATE > -1 Then
   TempDate = JOB_DATE
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
 If Trim(BATCH_NO) <> "" Then
      If WildCard(BATCH_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(BATCH_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(BATCH_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BATCH_NO = '" & ChangeQuote(Trim(BATCH_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (BATCH_NO = '" & ChangeQuote(Trim(BATCH_NO)) & "')"
         End If
      End If
   End If
   
   If APPROVED_BY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APPROVED_BY = " & APPROVED_BY & ") "
      Else
         WhereStr = WhereStr & "AND (APPROVED_BY = " & APPROVED_BY & ") "
      End If
   End If
   
   If RESPONSE_BY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RESPONSE_BY = " & RESPONSE_BY & ") "
      Else
         WhereStr = WhereStr & "AND (RESPONSE_BY = " & RESPONSE_BY & ") "
      End If
   End If
   
  If START_DATE > -1 Then
      TempDate = START_DATE
      If WhereStr = "" Then
         WhereStr = "WHERE (START_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (START_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
  If FINISH_DATE > -1 Then
      TempDate = FINISH_DATE
      If WhereStr = "" Then
         WhereStr = "WHERE (FINISH_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (FINISH_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUCCESS_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SUCCESS_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If JOB_DOC_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_DOC_TYPE = " & JOB_DOC_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_DOC_TYPE = " & JOB_DOC_TYPE & ") "
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If Len(ProcessSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PROCESS_ID IN " & ProcessSet & ") "
      Else
         WhereStr = WhereStr & "AND (PROCESS_ID IN " & ProcessSet & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY JOB_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY JOB_DESC "
         ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY JOB_DATE "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY BATCH_NO "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY START_DATE "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY FINISH_DATE "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY APPROVED_BY "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY RESPONSE_BY "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY PROCESS_NO "
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY DOC_NO "
      Else
         OrderByStr = "ORDER BY JOB_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PART_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PART_NO "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      JOB_DESC = NVLS(Rs("JOB_DESC"), "")
      PROCESS_ID = NVLI(Rs("PROCESS_ID"), -1)
      APPROVED_BY = NVLI(Rs("APPROVED_BY"), -1)
      RESPONSE_BY = NVLI(Rs("RESPONSE_BY"), -1)
      START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
      FINISH_DATE = InternalDateToDate(NVLS(Rs("FINISH_DATE"), ""))
      COMMIT_FLAG = NVLS(Rs("SUCCESS_FLAG"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      BATCH_NO = NVLS(Rs("BATCH_NO"), "")
'      LONG_NAMEA = NVLS(Rs("LONGA"), "")
'      LAST_NAMEA = NVLS(Rs("LASTA"), "")
'      LONG_NAMER = NVLS(Rs("LONGR"), "")
'      LAST_NAMER = NVLS(Rs("LASTR"), "")
'      DOC_NO = NVLS(Rs("DOCUMENT_NO"), "")
'      PROCESS_NAME = NVLS(Rs("PROCESS_NAME"), "")
      JOB_DOC_TYPE = NVLI(Rs("JOB_DOC_TYPE"), -1)
      OLD_COMMIT_FLAG = COMMIT_FLAG
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
      FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), -1)
      
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
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
Dim TempID5 As String
Dim TempID6 As String

   TempID1 = PROCESS_ID
   If PROCESS_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = APPROVED_BY
   If APPROVED_BY <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = RESPONSE_BY
   If RESPONSE_BY <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID4 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID4 = "NULL "
   End If
   
   TempID5 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID5 = "NULL "
   End If

   TempID6 = FORMULA_ID
   If FORMULA_ID <= 0 Then
      TempID6 = "NULL "
   End If

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO JOB " & _
                     "( " & _
                     "   JOB_ID, " & _
                     "   JOB_NO, " & _
                     "   JOB_DATE, " & _
                     "   JOB_DESC, " & _
                     "   BATCH_NO, " & _
                     "   PROCESS_ID, " & _
                     "   APPROVED_BY, " & _
                     "   RESPONSE_BY, " & _
                    "   START_DATE, " & _
                     "   FINISH_DATE, " & _
                     "   SUCCESS_FLAG, " & _
                        "   INVENTORY_DOC_ID, " & _
                        "   JOB_DOC_TYPE, " & _
                        "   PART_ITEM_ID, "
            SQL1 = SQL1 & _
                     "   ACTUAL_AMOUNT, " & _
                     "   STD_AMOUNT, " & _
                     "   FORMULA_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                         "'" & ChangeQuote(Trim(JOB_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(JOB_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(BATCH_NO)) & "'," & _
                        TempID1 & "," & _
                         TempID2 & "," & _
                          TempID3 & "," & _
                          "'" & ChangeQuote(Trim(DateToStringInt(START_DATE))) & "'," & _
                          "'" & ChangeQuote(Trim(DateToStringInt(FINISH_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                        TempID4 & "," & _
                        JOB_DOC_TYPE & "," & _
                        TempID5 & ","
               SQL1 = SQL1 & _
                        ACTUAL_AMOUNT & "," & _
                        STD_AMOUNT & "," & _
                        TempID6 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         JOB_ID = id
   Else 'Show edit
      SQL1 = "UPDATE JOB SET " & _
                     "   JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "', " & _
                     "   JOB_DATE= '" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "', " & _
                      "  JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "', " & _
                     "   BATCH_NO= '" & ChangeQuote(Trim(BATCH_NO)) & "', " & _
                    "   PROCESS_ID = " & TempID1 & "," & _
                     "   APPROVED_BY = " & TempID2 & "," & _
                      "   RESPONSE_BY = " & TempID3 & "," & _
                    "   START_DATE = '" & ChangeQuote(Trim(DateToStringInt(START_DATE))) & "', " & _
                    "   FINISH_DATE = '" & ChangeQuote(Trim(DateToStringInt(FINISH_DATE))) & "', " & _
                      "  SUCCESS_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   INVENTORY_DOC_ID = " & TempID4 & "," & _
                     "   JOB_DOC_TYPE = " & JOB_DOC_TYPE & "," & _
                     "   PART_ITEM_ID = " & TempID5 & "," & _
                     "   ACTUAL_AMOUNT = " & ACTUAL_AMOUNT & "," & _
                     "   STD_AMOUNT = " & STD_AMOUNT & "," & _
                     "   FORMULA_ID = " & TempID6 & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_ID = " & JOB_ID
   End If
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM JOB_PARAMETER WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM JOB_INOUT WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)

     SQL1 = "DELETE FROM JOB_RESOURCE WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
     SQL1 = "DELETE FROM JOB_VERIFY WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM JOB WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Parameters = Nothing
   Set Verifies = Nothing
End Sub
Public Function DeleteDataParameter()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM JOB_PARAMETER JP  WHERE JP.JOB_ID IN (SELECT UG.JOB_ID FROM JOB UG " & WhereStr & ")"
   m_Conn.Execute (SQL1)
      
End Function



