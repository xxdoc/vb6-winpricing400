VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLayout"
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

Public LAY_OUT_ID As Long
Public LAY_OUT_NO As String
Public LAY_OUT_NAME As String
Public LOCATION_ID As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, itemcount As Long)
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
   
   FromStr = "FROM LAY_OUT UG "
   
   WhereStr = ""
   If LAY_OUT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LAY_OUT_ID = " & LAY_OUT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LAY_OUT_ID = " & LAY_OUT_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY LAY_OUT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LAY_OUT_NAME "
   Else
      OrderByStr = "ORDER BY LAY_OUT_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   LAY_OUT_ID = NVLI(Rs("LAY_OUT_ID"), -1)
   LAY_OUT_NAME = NVLS(Rs("LAY_OUT_NAME"), "")
   LAY_OUT_NO = NVLS(Rs("LAY_OUT_NO"), "")
   LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   
   KEY_ID = LAY_OUT_ID
   KEY_LOOKUP = LAY_OUT_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = LOCATION_ID
   If LOCATION_ID < 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LAY_OUT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO LAY_OUT " & _
                     "( " & _
                     "   LAY_OUT_ID, " & _
                     "   LAY_OUT_NAME, " & _
                     "   LAY_OUT_NO, " & _
                     "   LOCATION_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(LAY_OUT_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(LAY_OUT_NO)) & "'," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LAY_OUT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LAY_OUT SET " & _
                     "   LAY_OUT_NO = '" & ChangeQuote(Trim(LAY_OUT_NO)) & "', " & _
                     "   LAY_OUT_NAME = '" & ChangeQuote(Trim(LAY_OUT_NAME)) & "', " & _
                     "   LOCATION_ID = " & TempID1 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LAY_OUT_ID = " & LAY_OUT_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LAY_OUT WHERE LAY_OUT_ID = " & LAY_OUT_ID
      m_Conn.Execute (SQL1)
End Function

