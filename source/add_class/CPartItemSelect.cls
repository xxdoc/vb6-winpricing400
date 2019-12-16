VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartItemSelect"
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

Public PART_ITEM_SELECT_ID As Long
Public PART_ITEM_SELECT_NO As String
Public PART_ITEM_ID As Long

Public PART_NO  As String
Public PART_GROUP_ID As Long
Public Part_Sel_Coll As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Part_Sel_Coll = New Collection
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

   SelectStr1 = "SELECT PIS.PART_ITEM_SELECT_ID,PIS.PART_ITEM_SELECT_NO,PI.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM_SELECT PIS "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM  PI ON (PIS.PART_ITEM_ID = PI.PART_ITEM_ID) "
    FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
    
   WhereStr = ""
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_NO = " & PART_NO & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_NO = " & PART_NO & ") "
      End If
   End If
   
   If Len(PART_ITEM_SELECT_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIS.PART_ITEM_SELECT_NO = '" & PART_ITEM_SELECT_NO & "') "
      Else
         WhereStr = WhereStr & "AND (PIS.PART_ITEM_SELECT_NO = '" & PART_ITEM_SELECT_NO & "') "
      End If
   End If
   
   If PART_GROUP_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI.PART_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      Else
         OrderByStr = "ORDER BY PI.PART_NO "
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
   PART_ITEM_SELECT_ID = NVLI(Rs("PART_ITEM_SELECT_ID"), -1)
   PART_ITEM_SELECT_NO = NVLS(Rs("PART_ITEM_SELECT_NO"), "")
   PART_ITEM_ID = NVLI(Rs("PREFPART_ITEM_IDIX_ID"), -1)
   PART_NO = NVLS(Rs("PART_NO"), "")
   PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("NAME_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PART_ITEM_SELECT " & _
                     "( " & _
                     "   PART_ITEM_SELECT_ID, " & _
                     "   PART_ITEM_SELECT_NO, " & _
                     "   PART_ITEM_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & PART_ITEM_SELECT_NO & "'," & _
                         "" & PART_ITEM_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_ITEM_SELECT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE NAME SET " & _
                     "   PART_ITEM_SELECT_NO = '" & PART_ITEM_SELECT_NO & "', " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_ITEM_SELECT_ID = " & ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
'      SQL1 = "DELETE FROM PART_ITEM_SELECT WHERE PART_ITEM_SELECT_NO = '" & PART_ITEM_SELECT_NO & "'"
     SQL1 = "DELETE FROM PART_ITEM_SELECT "
      m_Conn.Execute (SQL1)
End Function


Private Sub Class_Terminate()
   Set Part_Sel_Coll = Nothing
End Sub
