VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSystemParam"
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

Public PARAM_ID    As Long
Public PARAM_NAME    As String
Public PARAM_TYPE As String
Public PARAM_VALUE    As String
Public PARAM_DEFAULT    As String
Public PARAM_DESC As String

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

   SelectStr1 = "SELECT SP.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SYSTEM_PARAM SP "
   
   WhereStr = ""
   If PARAM_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PARAM_ID = " & PARAM_ID & ") "
      Else
         WhereStr = "AND (PARAM_ID = " & PARAM_ID & ") "
      End If
   End If
   
   If Trim(PARAM_NAME) <> "" Then
      If WildCard(PARAM_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PARAM_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PARAM_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PARAM_NAME = '" & ChangeQuote(Trim(PARAM_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (PARAM_NAME = '" & ChangeQuote(Trim(PARAM_NAME)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PARAM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PARAM_NAME "
      Else
         OrderByStr = "ORDER BY PARAM_NAME "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   PARAM_ID = NVLI(Rs("PARAM_ID"), -1)
   PARAM_NAME = NVLS(Rs("PARAM_NAME"), "")
   PARAM_TYPE = NVLS(Rs("PARAM_TYPE"), "")
   PARAM_VALUE = NVLS(Rs("PARAM_VALUE"), "")
   PARAM_DEFAULT = NVLS(Rs("PARAM_DEFAULT"), "")
   PARAM_DESC = NVLS(Rs("PARAM_DESC"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SYSTEM_PARAM_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO SYSTEM_PARAM " & _
                     "( " & _
                     "   PARAM_ID, " & _
                     "   PARAM_NAME, " & _
                     "   PARAM_TYPE, " & _
                     "   PARAM_VALUE, " & _
                     "   PARAM_DEFAULT, " & _
                     "   PARAM_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PARAM_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(PARAM_TYPE)) & "'," & _
                         "'" & ChangeQuote(Trim(PARAM_VALUE)) & "'," & _
                         "'" & ChangeQuote(Trim(PARAM_DEFAULT)) & "'," & _
                         "'" & ChangeQuote(Trim(PARAM_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE SYSTEM_PARAM SET " & _
                     "   PARAM_NAME = '" & ChangeQuote(Trim(PARAM_NAME)) & "', " & _
                     "   PARAM_TYPE = '" & ChangeQuote(Trim(PARAM_TYPE)) & "', " & _
                     "   PARAM_VALUE = '" & ChangeQuote(Trim(PARAM_VALUE)) & "', " & _
                     "   PARAM_DEFAULT = '" & ChangeQuote(Trim(PARAM_DEFAULT)) & "', " & _
                     "   PARAM_DESC = '" & ChangeQuote(Trim(PARAM_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PARAM_ID = " & PARAM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SYSTEM_PARAM WHERE PARAM_ID = " & PARAM_ID
      m_Conn.Execute (SQL1)
End Function



