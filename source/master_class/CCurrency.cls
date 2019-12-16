VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurrency"
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

Public CURRENCY_EXC_ID As Long
Public CURRENCY_EXC_NO As String
Public EXCHANGE_DATE As Date
Public MONEY_FAMILY_ID1 As Long
Public C1 As String
Public MONEY_FAMILY_ID2 As Long
Public C2 As String
Public COEFFICIENT As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   EXCHANGE_DATE = -1
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

   SelectStr1 = "SELECT CU.*,PR1.MONEY_FAMILY_NAME C1,PR2.MONEY_FAMILY_NAME C2 "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM CURRENCY CU "
     FromStr = FromStr & "LEFT OUTER JOIN MONEY_FAMILY PR1 ON (CU.MONEY_FAMILY_ID1 = PR1.MONEY_FAMILY_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN MONEY_FAMILY PR2 ON (CU.MONEY_FAMILY_ID2 = PR2.MONEY_FAMILY_ID) "

   If CURRENCY_EXC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CURRENCY_EXC_ID = " & CURRENCY_EXC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CURRENCY_EXC_ID = " & CURRENCY_EXC_ID & ") "
      End If
   End If
   
    If Trim(CURRENCY_EXC_NO) <> "" Then
      If WildCard(CURRENCY_EXC_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CURRENCY_EXC_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CURRENCY_EXC_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CURRENCY_EXC_NO = '" & ChangeQuote(Trim(CURRENCY_EXC_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CURRENCY_EXC_NO = '" & ChangeQuote(Trim(CURRENCY_EXC_NO)) & "')"
         End If
      End If
   End If
 
If Trim(COEFFICIENT) <> "" Then
      If WildCard(COEFFICIENT, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(COEFFICIENT,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(COEFFICIENT,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (COEFFICIENT = '" & ChangeQuote(Trim(COEFFICIENT)) & "')"
         Else
            WhereStr = WhereStr & "AND (COEFFICIENT = '" & ChangeQuote(Trim(COEFFICIENT)) & "')"
         End If
      End If
   End If
 
     
   If EXCHANGE_DATE > -1 Then
      TempDate = DateToStringIntLow(EXCHANGE_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE (EXCHANGE_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EXCHANGE_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If MONEY_FAMILY_ID1 > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONEY_FAMILY_ID1 = " & MONEY_FAMILY_ID1 & ") "
      Else
         WhereStr = WhereStr & "AND (MONEY_FAMILY_ID1 = " & MONEY_FAMILY_ID1 & ") "
      End If
   End If
   
   If MONEY_FAMILY_ID2 > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONEY_FAMILY_ID2 = " & MONEY_FAMILY_ID2 & ") "
      Else
         WhereStr = WhereStr & "AND (MONEY_FAMILY_ID2 = " & MONEY_FAMILY_ID2 & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CURRENCY_EXC_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EXCHANGE_DATE "
      ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY COEFFICIENT "
ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY MONEY_FAMILY_ID1 "
ElseIf OrderBy = 5 Then
      OrderByStr = "ORDER BY MONEY_FAMILY_ID2 "
      Else
      OrderByStr = "ORDER BY CURRENCY_EXC_ID "
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
   CURRENCY_EXC_ID = NVLI(Rs("CURRENCY_EXC_ID"), -1)
   CURRENCY_EXC_NO = NVLS(Rs("CURRENCY_EXC_NO"), "")
  C1 = NVLS(Rs("C1"), "")
   C2 = NVLS(Rs("C2"), "")
   MONEY_FAMILY_ID1 = NVLI(Rs("MONEY_FAMILY_ID1"), -1)
   MONEY_FAMILY_ID2 = NVLI(Rs("MONEY_FAMILY_ID2"), -1)
   EXCHANGE_DATE = InternalDateToDate(NVLS(Rs("EXCHANGE_DATE"), ""))
   COEFFICIENT = NVLD(Rs("COEFFICIENT"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = MONEY_FAMILY_ID1
   If MONEY_FAMILY_ID1 <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = MONEY_FAMILY_ID2
   If MONEY_FAMILY_ID2 <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CURRENCY_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO CURRENCY " & _
                     "( " & _
                     "   CURRENCY_EXC_ID, " & _
                     "   CURRENCY_EXC_NO, " & _
                     "   EXCHANGE_DATE, " & _
                     "   MONEY_FAMILY_ID1, " & _
                     "   MONEY_FAMILY_ID2, " & _
                     "   COEFFICIENT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(CURRENCY_EXC_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EXCHANGE_DATE))) & "'," & _
                        TempID1 & "," & _
                         TempID2 & "," & _
                           "'" & ChangeQuote(Trim(COEFFICIENT)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CURRENCY_EXC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CURRENCY SET " & _
                     "   CURRENCY_EXC_NO = '" & ChangeQuote(Trim(CURRENCY_EXC_NO)) & "', " & _
                     "   EXCHANGE_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXCHANGE_DATE))) & "', " & _
                    "   MONEY_FAMILY_ID1 = " & TempID1 & "," & _
                     "   MONEY_FAMILY_ID2 = " & TempID2 & "," & _
                    "  COEFFICIENT = '" & ChangeQuote(Trim(COEFFICIENT)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CURRENCY_EXC_ID = " & CURRENCY_EXC_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM CURRENCY WHERE CURRENCY_EXC_ID = " & CURRENCY_EXC_ID
      m_Conn.Execute (SQL1)
End Function


