VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSellType"
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

Public SELL_TYPE_ID As Long
Public SELL_TYPE_NO As String
Public SELL_TYPE_NAME As String

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
   
   FromStr = "FROM SELL_TYPE UG "
   
   WhereStr = ""
   If SELL_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SELL_TYPE_ID = " & SELL_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SELL_TYPE_ID = " & SELL_TYPE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SELL_TYPE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SELL_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY SELL_TYPE_ID "
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
   SELL_TYPE_ID = NVLI(Rs("SELL_TYPE_ID"), -1)
   SELL_TYPE_NAME = NVLS(Rs("SELL_TYPE_NAME"), "")
   SELL_TYPE_NO = NVLS(Rs("SELL_TYPE_NO"), "")
   
   KEY_ID = SELL_TYPE_ID
   KEY_LOOKUP = SELL_TYPE_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SELL_TYPE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SELL_TYPE " & _
                     "( " & _
                     "   SELL_TYPE_ID, " & _
                     "   SELL_TYPE_NAME, " & _
                     "   SELL_TYPE_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(SELL_TYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(SELL_TYPE_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SELL_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SELL_TYPE SET " & _
                     "   SELL_TYPE_NO = '" & ChangeQuote(Trim(SELL_TYPE_NO)) & "', " & _
                     "   SELL_TYPE_NAME = '" & ChangeQuote(Trim(SELL_TYPE_NAME)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SELL_TYPE_ID = " & SELL_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM SELL_TYPE WHERE SELL_TYPE_ID = " & SELL_TYPE_ID
      m_Conn.Execute (SQL1)
End Function


