VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLocation"
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

Public LOCATION_ID As Long
Public LOCATION_ID_SET As String
Public LOCATION_NAME As String
Public LOCATION_NO As String
Public LOCATION_TYPE As Long
Public SALE_FLAG  As String
Public LOCATION_GROUP As Long
Public LOCATION_GROUP_NAME As String
Public MAX_AMOUNT As Double
Public RX_AMOUNT As Double
Public UNIT_ID As Long
Public UNIT_NAME As String
Public KEY_CODE As String

Public PART_GROUP_ID As Long
Public KEY_ID As Long
Public KEY_LOOKUP As String

Public RightItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
   SALE_FLAG = ""
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
Dim SetStr As String

   SelectStr1 = "SELECT UG.*,MR.KEY_NAME LOCATION_GROUP_NAME,UN.UNIT_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM LOCATION UG "
   FromStr = FromStr & "LEFT JOIN MASTER_REF MR ON (UG.LOCATION_GROUP=MR.KEY_ID)  "
   FromStr = FromStr & "LEFT JOIN UNIT UN ON (UG.UNIT_ID=UN.UNIT_ID)  "
   WhereStr = ""
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Len(LOCATION_ID_SET) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (" & LOCATION_ID_SET & ") "
      Else
         WhereStr = WhereStr & "AND (& LOCATION_ID_SET &) "
      End If
   End If
   'LOCATION_ID_SET
   If LOCATION_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_TYPE = " & LOCATION_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_TYPE = " & LOCATION_TYPE & ") "
      End If
   End If
   
   If Len(SALE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SALE_FLAG = '" & SALE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SALE_FLAG = '" & SALE_FLAG & "') "
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      SetStr = "(SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (PART_GROUP_ID = " & PART_GROUP_ID & ") AND (SELECT_FLAG = 'Y'))"
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID IN " & SetStr & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID IN " & SetStr & ") "
      End If
   End If
   
    If LOCATION_GROUP > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_GROUP = " & LOCATION_GROUP & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_GROUP = " & LOCATION_GROUP & ") "
      End If
   End If
   
   If Len(KEY_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (KEY_CODE = '" & KEY_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (KEY_CODE = '" & KEY_CODE & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY LOCATION_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LOCATION_NAME "
   Else
      OrderByStr = "ORDER BY LOCATION_ID "
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
   LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
   LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
   LOCATION_TYPE = NVLI(Rs("LOCATION_TYPE"), 0)
   SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
   PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
   LOCATION_GROUP = NVLI(Rs("LOCATION_GROUP"), 0)
   LOCATION_GROUP_NAME = NVLS(Rs("LOCATION_GROUP_NAME"), "")
   MAX_AMOUNT = NVLD(Rs("MAX_AMOUNT"), 0)
   RX_AMOUNT = NVLD(Rs("RX_AMOUNT"), 0)
   UNIT_ID = NVLI(Rs("UNIT_ID"), 0)
   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   
   KEY_ID = LOCATION_ID
   KEY_LOOKUP = LOCATION_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String

   TempID1 = PART_GROUP_ID
   If TempID1 <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LOCATION_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO LOCATION " & _
                     "( " & _
                     "   LOCATION_ID, " & _
                     "   LOCATION_NAME, " & _
                     "   LOCATION_NO, " & _
                     "   LOCATION_TYPE, " & _
                     "   SALE_FLAG, " & _
                     "   PART_GROUP_ID, " & _
                     "   LOCATION_GROUP, " & _
                     "   MAX_AMOUNT, " & _
                     "   UNIT_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         id & "," & _
                         "'" & ChangeQuote(Trim(LOCATION_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(LOCATION_NO)) & "'," & _
                         LOCATION_TYPE & "," & _
                         "'" & ChangeQuote(Trim(SALE_FLAG)) & "',"
            SQL1 = SQL1 & _
                         TempID1 & "," & _
                         "" & LOCATION_GROUP & "," & _
                         "" & MAX_AMOUNT & "," & _
                         "" & UNIT_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LOCATION_ID = id
   Else 'Show edit
      SQL1 = "UPDATE LOCATION SET " & _
                     "   LOCATION_NAME = '" & ChangeQuote(Trim(LOCATION_NAME)) & "', " & _
                     "   LOCATION_NO = '" & ChangeQuote(Trim(LOCATION_NO)) & "', " & _
                     "   LOCATION_TYPE = " & LOCATION_TYPE & ", " & _
                     "   SALE_FLAG = '" & ChangeQuote(Trim(SALE_FLAG)) & "', " & _
                     "   PART_GROUP_ID = " & TempID1 & ", " & _
                     "   LOCATION_GROUP = " & LOCATION_GROUP & ", " & _
                     "   MAX_AMOUNT = " & MAX_AMOUNT & ", " & _
                     "   RX_AMOUNT = " & RX_AMOUNT & ", " & _
                     "   UNIT_ID = " & UNIT_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOCATION_ID = " & LOCATION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LOCATION WHERE LOCATION_ID = " & LOCATION_ID
      m_Conn.Execute (SQL1)
End Function


