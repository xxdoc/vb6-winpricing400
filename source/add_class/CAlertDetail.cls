VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAlertDetail"
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

Public ALERT_DETAIL_ID As Long
Public ALERT_BOX_ID As Long
Public USER_ID As Long
Public READ_FLAG As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public USER_NAME  As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT AD.*  "
      SelectStr1 = SelectStr1 & ",UA.USER_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM ALERT_DETAIL AD "
      FromStr = FromStr & "LEFT OUTER JOIN USER_ACCOUNT UA ON (UA.USER_ID = AD.USER_ID) "
   End If
   
   If ALERT_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AD.ALERT_DETAIL_ID = " & ALERT_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AD.ALERT_DETAIL_ID = " & ALERT_DETAIL_ID & ") "
      End If
   End If
   
   If ALERT_BOX_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AD.ALERT_BOX_ID = " & ALERT_BOX_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AD.ALERT_BOX_ID = " & ALERT_BOX_ID & ") "
      End If
   End If
   
   GroupByStr = ""
'   If Ind = 2 Then
   '   GroupByStr = "GROUP BY AD.USER_ACCOUNT_ID, PI.USER_NAME, PI.PART_DESC "
'   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY AD.USER_ACCOUNT_ID, PI.USER_NAME, PI.PART_DESC,PI.PARCEL_TYPE "
'   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      OrderByStr = "ORDER BY UA.USER_NAME "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      ALERT_DETAIL_ID = NVLI(Rs("ALERT_DETAIL_ID"), -1)
      ALERT_BOX_ID = NVLI(Rs("ALERT_BOX_ID"), -1)
      USER_ID = NVLI(Rs("USER_ID"), -1)
      
      READ_FLAG = NVLS(Rs("READ_FLAG"), "")
      
      USER_NAME = NVLS(Rs("USER_NAME"), "")
   ElseIf Ind = 2 Then
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
   
   TempID1 = USER_ID
   If USER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO ALERT_DETAIL " & _
                     "( " & _
                     "   ALERT_BOX_ID, " & _
                     "   USER_ID, " & _
                     "   READ_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        ALERT_BOX_ID & "," & _
                        TempID1 & "," & _
                        "'" & ChangeQuote(Trim(READ_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING ALERT_DETAIL_ID "
   Else 'Show edit
      SQL1 = "UPDATE ALERT_DETAIL SET " & _
                     "  READ_FLAG = '" & ChangeQuote(Trim(READ_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ALERT_DETAIL_ID = " & ALERT_DETAIL_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      ALERT_DETAIL_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ALERT_DETAIL WHERE ALERT_DETAIL_ID = " & ALERT_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function
