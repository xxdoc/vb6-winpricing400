VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSliptSalary"
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
Public SliptAdd As Collection
Public SliptSub As Collection
' End of generic data member

Public SLIPT_SALARY_ID As Long
Public EMP_ID As Long
Public MONTH_NO As Long
Public YEAR_NO As Long
Public SALARY As Double
Public SUM_SLIPT_ADD As Double
Public SUM_SLIPT_SUB As Double
Public SUM_BORROW As Double
Public LOCK_FLAG As String
Public FROM_MONTH As Long
Public FROM_YEAR As Long
Public TO_MONTH As Long
Public TO_YEAR As Long
Public MONTHLY_ADD_ID As Long
Public MONTHLY_SUB_ID As Long
Public SUM_ADD As Double
Public SUM_SUB As Double
Public SA As CSliptAdd
Public SB As CSliptSub


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set SliptAdd = New Collection
   Set SliptSub = New Collection
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

   SelectStr1 = "SELECT SS.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SLIPT_SALARY SS "
   
   WhereStr = ""
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If

   If MONTH_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO = " & MONTH_NO & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO = " & MONTH_NO & ") "
      End If
   End If
   
   If YEAR_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO = " & YEAR_NO & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO = " & YEAR_NO & ") "
      End If
   End If
   
   If FROM_MONTH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO >= " & FROM_MONTH & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO >= " & FROM_MONTH & ") "
      End If
   End If
   
   If FROM_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO >= " & FROM_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO >= " & FROM_YEAR & ") "
      End If
   End If
   If TO_MONTH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO <= " & TO_MONTH & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO <= " & TO_MONTH & ") "
      End If
   End If
   
   If TO_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO <= " & TO_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO <= " & TO_YEAR & ") "
      End If
   End If
   
      
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = " ORDER BY MONTH_NO , YEAR_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EMP_ID "
   Else
      OrderByStr = "ORDER BY SLIPT_SALARY_ID "
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
'In this class we will not handle any error or any transaction
Public Function QueryData1(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT SUM(SA.MONTHLY_AMOUNT) SUM_ADD  "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SLIPT_SALARY SS "
   FromStr = FromStr & "LEFT OUTER JOIN SLIPT_ADD SA ON (SS.SLIPT_SALARY_ID = SA.SLIPT_SALARY_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN MONTHLY_ADD MA ON (SA.MONTHLY_ADD = MA.MONTHLY_ADD_ID) "
   
   WhereStr = ""
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If

   If MONTH_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO = " & MONTH_NO & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO = " & MONTH_NO & ") "
      End If
   End If
   
   If YEAR_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO = " & YEAR_NO & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO = " & YEAR_NO & ") "
      End If
   End If
   
   If FROM_MONTH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO >= " & FROM_MONTH & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO >= " & FROM_MONTH & ") "
      End If
   End If
   
   If FROM_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO >= " & FROM_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO >= " & FROM_YEAR & ") "
      End If
   End If
   If TO_MONTH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO <= " & TO_MONTH & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO <= " & TO_MONTH & ") "
      End If
   End If
   
   If TO_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO <= " & TO_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO <= " & TO_YEAR & ") "
      End If
   End If
   
If MONTHLY_ADD_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTHLY_ADD_ID = " & MONTHLY_ADD_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MONTHLY_ADD_ID = " & MONTHLY_ADD_ID & ") "
      End If
   End If
      
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
'''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function
Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT SUM(SB.MONTHLY_AMOUNT) SUM_SUB  "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SLIPT_SALARY SS "
   FromStr = FromStr & "LEFT OUTER JOIN SLIPT_SUB SB ON (SS.SLIPT_SALARY_ID = SB.SLIPT_SALARY_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN MONTHLY_SUB MB ON (SB.MONTHLY_SUB = MB.MONTHLY_SUB_ID) "
   
   WhereStr = ""
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If

   If MONTH_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO = " & MONTH_NO & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO = " & MONTH_NO & ") "
      End If
   End If
   
   If YEAR_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO = " & YEAR_NO & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO = " & YEAR_NO & ") "
      End If
   End If
   
   If FROM_MONTH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO >= " & FROM_MONTH & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO >= " & FROM_MONTH & ") "
      End If
   End If
   
   If FROM_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO >= " & FROM_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO >= " & FROM_YEAR & ") "
      End If
   End If
   If TO_MONTH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTH_NO <= " & TO_MONTH & ") "
      Else
         WhereStr = WhereStr & "AND (MONTH_NO <= " & TO_MONTH & ") "
      End If
   End If
   
   If TO_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO <= " & TO_YEAR & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO <= " & TO_YEAR & ") "
      End If
   End If
   
If MONTHLY_SUB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTHLY_SUB_ID = " & MONTHLY_SUB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MONTHLY_SUB_ID = " & MONTHLY_SUB_ID & ") "
      End If
   End If
      
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)

   SLIPT_SALARY_ID = NVLI(Rs("SLIPT_SALARY_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   MONTH_NO = NVLI(Rs("MONTH_NO"), -1)
   YEAR_NO = NVLI(Rs("YEAR_NO"), -1)
   SALARY = NVLD(Rs("SALARY"), -1)
   SUM_SLIPT_ADD = NVLD(Rs("SUM_SLIPT_ADD"), -1)
   SUM_SLIPT_SUB = NVLD(Rs("SUM_SLIPT_SUB"), -1)
   SUM_BORROW = NVLD(Rs("SUM_BORROW"), -1)
   LOCK_FLAG = NVLS(Rs("LOCK_FLAG"), "")
End Function
Public Function PopulateFromRS1(Rs As ADODB.Recordset)
SUM_ADD = NVLD(Rs("SUM_ADD"), 0)
End Function
Public Function PopulateFromRS2(Rs As ADODB.Recordset)
SUM_SUB = NVLD(Rs("SUM_SUB"), 0)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SLIPT_SALARY_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO SLIPT_SALARY " & _
                     "( " & _
                     "   SLIPT_SALARY_ID, " & _
                     "   EMP_ID, " & _
                     "   MONTH_NO, " & _
                    "   YEAR_NO, " & _
                     "   SALARY, " & _
                     "   SUM_SLIPT_ADD, " & _
                        "   SUM_SLIPT_SUB, " & _
                     "   SUM_BORROW, " & _
                     "   LOCK_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         id & "," & _
                         "'" & ChangeQuote(Trim(EMP_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(MONTH_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(YEAR_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(SALARY)) & "'," & _
                        "'" & ChangeQuote(Trim(SUM_SLIPT_ADD)) & "'," & _
                        "'" & ChangeQuote(Trim(SUM_SLIPT_SUB)) & "'," & _
                        "'" & ChangeQuote(Trim(SUM_BORROW)) & "'," & _
                        "'" & ChangeQuote(Trim(LOCK_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SLIPT_SALARY_ID = id
   Else 'Show edit
      SQL1 = "UPDATE SLIPT_SALARY SET " & _
                     "   SALARY = '" & ChangeQuote(Trim(SALARY)) & "', " & _
                     "   SUM_SLIPT_ADD = '" & ChangeQuote(Trim(SUM_SLIPT_ADD)) & "', " & _
                     "   SUM_SLIPT_SUB = '" & ChangeQuote(Trim(SUM_SLIPT_SUB)) & "', " & _
                     "   SUM_BORROW = '" & ChangeQuote(Trim(SUM_BORROW)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SLIPT_SALARY_ID = " & SLIPT_SALARY_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
    
      SQL1 = "DELETE FROM SLIPT_ADD WHERE SLIPT_SALARY_ID = " & SLIPT_SALARY_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM SLIPT_SUB WHERE SLIPT_SALARY_ID = " & SLIPT_SALARY_ID
      m_Conn.Execute (SQL1)
      SQL1 = "DELETE FROM SLIPT_SALARY WHERE SLIPT_SALARY_ID = " & SLIPT_SALARY_ID
      m_Conn.Execute (SQL1)
End Function


