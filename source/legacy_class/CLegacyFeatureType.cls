VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLegacyFeatureType"
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

Public FEATYPE_ID    As Long
Public FEATYPE_NAME    As String
Public FEATYPE_CODE As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
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
   
   FromStr = "FROM FEATURE_TYPE ACT "
   
   WhereStr = ""
   If FEATYPE_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEATYPE_ID = " & FEATYPE_ID & ") "
      Else
         WhereStr = "AND (FEATYPE_ID = " & FEATYPE_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FEATYPE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY FEATYPE_NAME "
      Else
         OrderByStr = "ORDER BY FEATYPE_ID "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   FEATYPE_ID = NVLI(Rs("FEATYPE_ID"), -1)
   FEATYPE_NAME = NVLS(Rs("FEATYPE_NAME"), "")
   FEATYPE_CODE = NVLS(Rs("FEATYPE_CODE"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FEATURE_TYPE_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO FEATURE_TYPE " & _
                     "( " & _
                     "   FEATYPE_ID, " & _
                     "   FEATYPE_NAME, " & _
                     "   FEATYPE_CODE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(FEATYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(FEATYPE_CODE)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE FEATURE_TYPE SET " & _
                     "   FEATYPE_NAME = '" & ChangeQuote(Trim(FEATYPE_NAME)) & "', " & _
                     "   FEATYPE_CODE = '" & ChangeQuote(Trim(FEATYPE_CODE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FEATYPE_ID = " & FEATYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM FEATURE_TYPE WHERE FEATYPE_ID = " & FEATYPE_ID
      m_Conn.Execute (SQL1)
End Function



