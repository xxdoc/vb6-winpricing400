VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMonthlyAdd"
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

Public MONTHLY_ADD_ID    As Long
Public MONTHLY_ADD_NAME    As String
Public MONTHLY_ADD_NO    As String


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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT ACT.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MONTHLY_ADD ACT "
   
   WhereStr = ""
   If MONTHLY_ADD_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTHLY_ADD_ID = " & MONTHLY_ADD_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MONTHLY_ADD_ID = " & MONTHLY_ADD_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY MONTHLY_ADD_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY MONTHLY_ADD_NAME "
      Else
         OrderByStr = "ORDER BY MONTHLY_ADD_NO "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   MONTHLY_ADD_ID = NVLI(Rs("MONTHLY_ADD_ID"), -1)
   MONTHLY_ADD_NAME = NVLS(Rs("MONTHLY_ADD_NAME"), "")
   MONTHLY_ADD_NO = NVLS(Rs("MONTHLY_ADD_NO"), "")
   
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MONTHLY_ADD_SEQ", id, glbErrorLog)
      SQL1 = "INSERT INTO MONTHLY_ADD " & _
                     "( " & _
                     "   MONTHLY_ADD_ID, " & _
                    "   MONTHLY_ADD_NO, " & _
                     "   MONTHLY_ADD_NAME, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         id & "," & _
                          "'" & ChangeQuote(Trim(MONTHLY_ADD_NO)) & "', " & _
                         "'" & ChangeQuote(Trim(MONTHLY_ADD_NAME)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE MONTHLY_ADD SET " & _
                    "   MONTHLY_ADD_NO = '" & ChangeQuote(Trim(MONTHLY_ADD_NO)) & "', " & _
                     "   MONTHLY_ADD_NAME = '" & ChangeQuote(Trim(MONTHLY_ADD_NAME)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & _
                     "WHERE MONTHLY_ADD_ID = " & MONTHLY_ADD_ID
   End If
   ''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM MONTHLY_ADD WHERE MONTHLY_ADD_ID = " & MONTHLY_ADD_ID
      m_Conn.Execute (SQL1)
End Function

