VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFeatureType"
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

Public FEATURE_TYPE_ID As Long
Public FEATURE_TYPE_NO As String
Public FEATURE_TYPE_NAME As String
Public LOGISTIC_FLAG As String
Public FEATURE_GROUP_ID As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FEATURE_TYPE UG "
   
   WhereStr = ""
   If FEATURE_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEATURE_TYPE_ID = " & FEATURE_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FEATURE_TYPE_ID = " & FEATURE_TYPE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FEATURE_TYPE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY FEATURE_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY FEATURE_TYPE_ID "
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
   FEATURE_TYPE_ID = NVLI(Rs("FEATURE_TYPE_ID"), -1)
   FEATURE_TYPE_NAME = NVLS(Rs("FEATURE_TYPE_NAME"), "")
   FEATURE_TYPE_NO = NVLS(Rs("FEATURE_TYPE_NO"), "")
   LOGISTIC_FLAG = NVLS(Rs("LOGISTIC_FLAG"), "")
   FEATURE_GROUP_ID = NVLI(Rs("FEATURE_GROUP_ID"), -1)
   
   KEY_ID = FEATURE_TYPE_ID
   KEY_LOOKUP = FEATURE_TYPE_NO
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = FEATURE_GROUP_ID
   If FEATURE_GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FEATURE_TYPE_SEQ", ID, glbErrorLog)
      
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("FEATURE_TYPE_SEQ", ID, glbErrorLog)
      Else
         ID = FEATURE_TYPE_ID
      End If
      
      SQL1 = "INSERT INTO FEATURE_TYPE " & _
                     "( " & _
                     "   FEATURE_TYPE_ID, " & _
                     "   FEATURE_TYPE_NAME, " & _
                     "   FEATURE_TYPE_NO, " & _
                     "   LOGISTIC_FLAG, " & _
                     "   FEATURE_GROUP_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(FEATURE_TYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(FEATURE_TYPE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(LOGISTIC_FLAG)) & "'," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         FEATURE_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE FEATURE_TYPE SET " & _
                     "   FEATURE_TYPE_NO = '" & ChangeQuote(Trim(FEATURE_TYPE_NO)) & "', " & _
                     "   FEATURE_TYPE_NAME = '" & ChangeQuote(Trim(FEATURE_TYPE_NAME)) & "', " & _
                     "   LOGISTIC_FLAG = '" & ChangeQuote(Trim(LOGISTIC_FLAG)) & "', " & _
                     "   FEATURE_GROUP_ID = " & TempID1 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FEATURE_TYPE_ID = " & FEATURE_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM FEATURE_TYPE WHERE FEATURE_TYPE_ID = " & FEATURE_TYPE_ID
      m_Conn.Execute (SQL1)
End Function




