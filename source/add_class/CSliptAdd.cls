VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSliptAdd"
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
Public SLIPT_ADD_ID As Long
Public SLIPT_SALARY_ID As Long
Public MONTHLY_ADD As Long
Public MONTHLY_AMOUNT As Double
Public MONTHLY_NAME As String


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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

   SelectStr1 = "SELECT SA.*,MA.MONTHLY_ADD_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SLIPT_ADD SA "
   FromStr = FromStr & "LEFT OUTER JOIN MONTHLY_ADD MA ON (SA.MONTHLY_ADD = MA.MONTHLY_ADD_ID) "
   WhereStr = ""
   If SLIPT_SALARY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SLIPT_SALARY_ID = " & SLIPT_SALARY_ID & ") "
      Else
         WhereStr = WhereStr & " AND (SLIPT_SALARY_ID = " & SLIPT_SALARY_ID & ") "
      End If
   End If
   If MONTHLY_ADD > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTHLY_ADD = " & MONTHLY_ADD & ") "
      Else
         WhereStr = WhereStr & "AND (MONTHLY_ADD = " & MONTHLY_ADD & ") "
      End If
   End If

   If SLIPT_ADD_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SLIPT_ADD_ID = " & SLIPT_ADD_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SLIPT_ADD_ID = " & SLIPT_ADD_ID & ") "
      End If
   End If
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SLIPT_SALARY_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EMP_ID "
   Else
      OrderByStr = "ORDER BY MONTHLY_ADD "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   SLIPT_SALARY_ID = NVLI(Rs("SLIPT_SALARY_ID"), -1)
   SLIPT_ADD_ID = NVLI(Rs("SLIPT_ADD_ID"), -1)
   MONTHLY_ADD = NVLI(Rs("MONTHLY_ADD"), -1)
   MONTHLY_AMOUNT = NVLD(Rs("MONTHLY_AMOUNT"), -1)
   MONTHLY_NAME = NVLS(Rs("MONTHLY_ADD_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'NAME_ID & "," & _
" NAME_ID = " & NAME_ID & ", " & _
'
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SLIPT_ADD_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO SLIPT_ADD " & _
                     "( " & _
                     "   SLIPT_ADD_ID, " & _
                     "   SLIPT_SALARY_ID, " & _
                     "   MONTHLY_ADD, " & _
                     "   MONTHLY_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         id & "," & _
                         "'" & ChangeQuote(Trim(SLIPT_SALARY_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(MONTHLY_ADD)) & "'," & _
                        "'" & ChangeQuote(Trim(MONTHLY_AMOUNT)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SLIPT_ADD_ID = id
   Else 'Show edit
      SQL1 = "UPDATE SLIPT_ADD SET " & _
                     "   MONTHLY_ADD = '" & ChangeQuote(Trim(MONTHLY_ADD)) & "', " & _
                     "   MONTHLY_AMOUNT = '" & ChangeQuote(Trim(MONTHLY_AMOUNT)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SLIPT_ADD_ID = " & SLIPT_ADD_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM SLIPT_ADD WHERE SLIPT_ADD_ID = " & SLIPT_ADD_ID
      m_Conn.Execute (SQL1)
End Function


