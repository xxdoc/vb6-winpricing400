VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParameterItem"
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

Public PARAMETER_ITEM_ID As Long
Public PARAMETER_ID As Long
Public PROCESS_ID As Long
Public SELECT_FLAG As String

Public HGI_PARAMETER_ID As Long
Public PARAMETER_PROCESS_NO As String
Public PARAMETER_PROCESS_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT HG.PARAMETER_PROCESS_NO, HG.PARAMETER_PROCESS_NAME, HG.PARAMETER_PROCESS_ID HG_PARAMETER_PROCESS_ID,HGI.PARAMETER_ID HGI_PARAMETER_ID, HGI.* "
   
   FromStr = "FROM PARAMETER_ITEM HGI "
   FromStr = FromStr & "LEFT OUTER JOIN PARAMETER_PROCESS HG ON (HGI.PARAMETER_ID = HG.PARAMETER_PROCESS_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PROCESS LC ON (HGI.PROCESS_ID = LC.PROCESS_ID) "

   SelectStr2 = "SELECT COUNT(*)  "
      
   If PARAMETER_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PARAMETER_ITEM_ID = " & PARAMETER_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PARAMETER_ITEM_ID = " & PARAMETER_ITEM_ID & ") "
      End If
   End If
   
   If PARAMETER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PARAMETER_ID = " & PARAMETER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PARAMETER_ID = " & PARAMETER_ID & ") "
      End If
   End If
  If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (HGI.PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (HGI.PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
   
   
   If Len(SELECT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SELETE_FLAG = '" & SELECT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SELECT_FLAG = '" & SELECT_FLAG & "') "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PARAMETER_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PARAMETER_ID "
   Else
      OrderByStr = "ORDER BY PARAMETER_ID "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   PARAMETER_ITEM_ID = NVLI(Rs("PARAMETER_ITEM_ID"), -1)
   PARAMETER_ID = NVLI(Rs("PARAMETER_ID"), -1)
   PROCESS_ID = NVLI(Rs("PROCESS_ID"), -1)
   
   HGI_PARAMETER_ID = NVLI(Rs("HGI_PARAMETER_ID"), -1)
   PARAMETER_PROCESS_NO = NVLS(Rs("PARAMETER_PROCESS_NO"), "")
   PARAMETER_PROCESS_NAME = NVLS(Rs("PARAMETER_PROCESS_NAME"), "")
   
   If PARAMETER_ID <> HGI_PARAMETER_ID Then
      SELECT_FLAG = "N"
   Else
      SELECT_FLAG = NVLS(Rs("SELECT_FLAG"), "")
   End If
   
    KEY_ID = PARAMETER_ID
   KEY_LOOKUP = PARAMETER_PROCESS_NO

End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = PARAMETER_ITEM_ID
   If PARAMETER_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PARAMETER_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PARAMETER_ITEM " & _
                     "( " & _
                     "   PARAMETER_ITEM_ID, " & _
                     "   PARAMETER_ID, " & _
                     "   PROCESS_ID, " & _
                     "   SELECT_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PARAMETER_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(PROCESS_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(SELECT_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PARAMETER_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PARAMETER_ITEM SET " & _
                     "   PARAMETER_ID = '" & ChangeQuote(Trim(PARAMETER_ID)) & "', " & _
                     "   PROCESS_ID = '" & ChangeQuote(Trim(PROCESS_ID)) & "', " & _
                     "   SELECT_FLAG = '" & ChangeQuote(Trim(SELECT_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PARAMETER_ITEM_ID = " & PARAMETER_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PARAMETER_ITEM WHERE PARAMETER_ITEM_ID = " & PARAMETER_ITEM_ID
      m_Conn.Execute (SQL1)
End Function


