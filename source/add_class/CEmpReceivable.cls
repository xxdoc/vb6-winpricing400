VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpReceivable"
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

Public EMP_RECEIVABLE_ID As Long
Public EMP_ID As Long
Public BORROW_NO As String
Public BORROW_AMOUNT As Double
Public CLOSED_FLAG As String
Public PAID_AMOUNT As Double
Public BORROW_DESC As String
Public BORROW_DATE As Date
Public EFFECTIVE_DATE As Date
Public JOURNAL_ID As Long
Public LONG_NAME As String
Public LAST_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public EMP_NAME As String
Public OLD_AMOUNT As Double


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT JT.*, NM.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMP_RECEIVABLE JT "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (JT.EMP_ID = EN.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "

   WhereStr = ""
   
   If EMP_RECEIVABLE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_RECEIVABLE_ID = " & EMP_RECEIVABLE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_RECEIVABLE_ID = " & EMP_RECEIVABLE_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   If Trim(BORROW_NO) <> "" Then
      If WildCard(BORROW_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(BORROW_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(BORROW_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BORROW_NO = '" & ChangeQuote(Trim(BORROW_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (BORROW_NO = '" & ChangeQuote(Trim(BORROW_NO)) & "')"
         End If
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BORROW_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BORROW_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BORROW_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BORROW_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(EMP_NAME) <> "" Then
      If WildCard(EMP_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
   OrderByStr = "ORDER BY BORROW_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY BORROW_DATE "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY LONG_NAME "
ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY BORROW_AMOUNT "
   Else
      OrderByStr = "ORDER BY EMP_RECEIVABLE_ID  "
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
   ''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   EMP_RECEIVABLE_ID = NVLI(Rs("EMP_RECEIVABLE_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   BORROW_NO = NVLS(Rs("BORROW_NO"), "")
   BORROW_AMOUNT = NVLD(Rs("BORROW_AMOUNT"), 0)
   OLD_AMOUNT = NVLD(Rs("BORROW_AMOUNT"), 0)
   CLOSED_FLAG = NVLS(Rs("CLOSED_FLAG"), "")
   PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), -1)
   BORROW_DESC = NVLS(Rs("BORROW_DESC"), "")
   EFFECTIVE_DATE = InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), ""))
   BORROW_DATE = InternalDateToDate(NVLS(Rs("BORROW_DATE"), ""))
   JOURNAL_ID = NVLI(Rs("JOURNAL_ID"), -1)
   LONG_NAME = NVLS(Rs("LONG_NAME"), "")
   LAST_NAME = NVLS(Rs("LAST_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   Diff = BORROW_AMOUNT - OLD_AMOUNT
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMP_RECEIVABLE_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO EMP_RECEIVABLE " & _
                     "( " & _
                     "   EMP_RECEIVABLE_ID, " & _
                     "   EMP_ID, " & _
                     "   BORROW_NO, " & _
                     "   BORROW_AMOUNT, " & _
                     "   CLOSED_FLAG, " & _
                     "   PAID_AMOUNT, " & _
                     "   BORROW_DESC, " & _
                     "   EFFECTIVE_DATE, " & _
                     "   BORROW_DATE, " & _
                     "   JOURNAL_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         id & "," & _
                         "" & EMP_ID & "," & _
                         "'" & BORROW_NO & "'," & _
                         "" & BORROW_AMOUNT & "," & _
                         "'" & CLOSED_FLAG & "'," & _
                         "" & PAID_AMOUNT & "," & _
                         "'" & BORROW_DESC & "'," & _
                         "'" & DateToStringInt(Trim(EFFECTIVE_DATE)) & "'," & _
                         "'" & DateToStringInt(Trim(BORROW_DATE)) & "'," & _
                         "" & JOURNAL_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMP_RECEIVABLE_ID = id
   Else 'Show edit
      SQL1 = "UPDATE EMP_RECEIVABLE SET " & _
                     "   EMP_ID = " & EMP_ID & ", " & _
                     "   BORROW_NO = '" & BORROW_NO & "', " & _
                     "   BORROW_AMOUNT = " & BORROW_AMOUNT & ", " & _
                     "   CLOSED_FLAG = '" & CLOSED_FLAG & "', " & _
                     "   PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "   BORROW_DESC = '" & BORROW_DESC & "', " & _
                     "    EFFECTIVE_DATE = '" & DateToStringInt(EFFECTIVE_DATE) & "', " & _
                     "    BORROW_DATE = '" & DateToStringInt(BORROW_DATE) & "', " & _
                     "   JOURNAL_ID = " & JOURNAL_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_RECEIVABLE_ID = " & EMP_RECEIVABLE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
      
      SQL1 = "UPDATE EMPLOYEE SET CREDIT_AMOUNT = CREDIT_AMOUNT + " & Diff & " WHERE EMP_ID = " & EMP_ID & " "
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
            
      SQL1 = "UPDATE EMPLOYEE SET CREDIT_AMOUNT = CREDIT_AMOUNT - " & BORROW_AMOUNT & " WHERE EMP_ID = " & EMP_ID & " "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EMP_RECEIVABLE WHERE EMP_RECEIVABLE_ID = " & EMP_RECEIVABLE_ID
      m_Conn.Execute (SQL1)
End Function


Public Function AddEditDataFlag()
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
      SQL1 = "UPDATE EMP_RECEIVABLE SET " & _
                     "   CLOSED_FLAG = '" & CLOSED_FLAG & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_ID = " & EMP_ID
   ''Debug.Print (SQL1)
      m_Conn.Execute (SQL1)
      
End Function

