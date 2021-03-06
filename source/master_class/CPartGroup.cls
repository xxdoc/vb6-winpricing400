VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartGroup"
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

Public PART_GROUP_ID As Long
Public PART_GROUP_NAME As String
Public PART_GROUP_NO As String
Public SET_PRODUCT_ID As Long

Public SELECT_FLAG As String
Public KEY_ID As Long
Public KEY_LOOKUP As String

Public PartGroupItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartGroupItems = New Collection
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
   
   FromStr = "FROM PART_GROUP UG "
   
   WhereStr = ""
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = "AND (PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PART_GROUP_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PART_GROUP_NAME "
   Else
      OrderByStr = "ORDER BY PART_GROUP_ID "
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
   PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
   PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
   PART_GROUP_NO = NVLS(Rs("PART_GROUP_NO"), "")
   SET_PRODUCT_ID = NVLI(Rs("SET_PRODUCT_ID"), -1)
   
   KEY_ID = PART_GROUP_ID
   KEY_LOOKUP = PART_GROUP_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = PART_GROUP_ID
   If PART_GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = SET_PRODUCT_ID
   If SET_PRODUCT_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_GROUP_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PART_GROUP " & _
                     "( " & _
                     "   PART_GROUP_ID, " & _
                     "   PART_GROUP_NAME, " & _
                     "   PART_GROUP_NO, " & _
                     "   SET_PRODUCT_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PART_GROUP_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_GROUP_NO)) & "'," & _
                         TempID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_GROUP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_GROUP SET " & _
                     "   PART_GROUP_NAME = '" & ChangeQuote(Trim(PART_GROUP_NAME)) & "', " & _
                     "   PART_GROUP_NO = '" & ChangeQuote(Trim(PART_GROUP_NO)) & "', " & _
                     "   SET_PRODUCT_ID = " & TempID2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_GROUP_ID = " & PART_GROUP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PART_GROUP WHERE PART_GROUP_ID = " & PART_GROUP_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM HGROUP_ITEM WHERE PART_GROUP_ID = " & PART_GROUP_ID
      m_Conn.Execute (SQL1)
End Function




