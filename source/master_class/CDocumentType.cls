VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentType"
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

Public DOCTYPE_ID    As Long
Public DOCTYPE_NAME    As String
Public DOCTYPE_NO    As String
Public CARD_ID_FLAG As String

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
   
   FromStr = "FROM DOCUMENT_TYPE ACT "
   
   WhereStr = ""
   If DOCTYPE_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCTYPE_ID = " & DOCTYPE_ID & ") "
      Else
         WhereStr = "AND (DOCTYPE_ID = " & DOCTYPE_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCTYPE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY DOCTYPE_NAME "
      Else
         OrderByStr = "ORDER BY DOCUMENT_NO "
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
   DOCTYPE_ID = NVLI(Rs("DOCTYPE_ID"), -1)
   DOCTYPE_NAME = NVLS(Rs("DOCTYPE_NAME"), "")
   CARD_ID_FLAG = NVLS(Rs("CARD_ID_FLAG"), "N")
   DOCTYPE_NO = NVLS(Rs("DOCUMENT_NO"), "")
   
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   CARD_ID_FLAG = "N"
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DOCUMENT_TYPE_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO DOCUMENT_TYPE " & _
                     "( " & _
                     "   DOCTYPE_ID, " & _
                     "   DOCTYPE_NAME, " & _
                     "   CARD_ID_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY, " & _
                    "   DOCUMENT_NO " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCTYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(CARD_ID_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ChangeQuote(Trim(DOCTYPE_NO)) & "' " & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE DOCUMENT_TYPE SET " & _
                     "   DOCTYPE_NAME = '" & ChangeQuote(Trim(DOCTYPE_NAME)) & "', " & _
                     "   CARD_ID_FLAG = '" & ChangeQuote(Trim(CARD_ID_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & ", " & _
                    "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCTYPE_NO)) & "'  " & _
                     "WHERE DOCTYPE_ID = " & DOCTYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM DOCUMENT_TYPE WHERE DOCTYPE_ID = " & DOCTYPE_ID
      m_Conn.Execute (SQL1)
End Function


