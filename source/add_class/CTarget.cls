VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTarget"
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

Public TARGET_ID As Long
Public YEAR_NO As Long
Public TARGET_DESC As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public CollTargerDetail As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
    
   Set CollTargerDetail = New Collection
End Sub

Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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
      SelectStr1 = "SELECT TG.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM TARGET TG "
   End If
   
   If TARGET_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TG.TARGET_ID = " & TARGET_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TG.TARGET_ID = " & TARGET_ID & ") "
      End If
   End If
   
   If YEAR_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TG.YEAR_NO = " & YEAR_NO & ") "
      Else
         WhereStr = WhereStr & "AND (TG.YEAR_NO = " & YEAR_NO & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY TG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY TG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      OrderByStr = "ORDER BY TG.YEAR_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TARGET_ID = NVLI(Rs("TARGET_ID"), -1)
      
      YEAR_NO = NVLI(Rs("YEAR_NO"), -1)
      TARGET_DESC = NVLS(Rs("TARGET_DESC"), "")
      
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

'   TempID1 = PROCESS_ID
'   If PROCESS_ID <= 0 Then
'      TempID1 = "NULL "
'   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO TARGET " & _
                     "( " & _
                     "   YEAR_NO, " & _
                     "   TARGET_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_NAME " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        YEAR_NO & "," & _
                         "'" & ChangeQuote(Trim(TARGET_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         "'" & ChangeQuote(Trim(glbUser.USER_NAME)) & "'" & _
                     ")  RETURNING TARGET_ID "
   Else 'Show edit
      SQL1 = "UPDATE TARGET SET " & _
                    "    YEAR_NO = " & YEAR_NO & ", " & _
                    "  TARGET_DESC = '" & ChangeQuote(Trim(TARGET_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_NAME = '" & glbUser.USER_NAME & "'" & _
                     "WHERE TARGET_ID = " & TARGET_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      TARGET_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM TARGET_DETAIL WHERE TARGET_ID = " & TARGET_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM TARGET WHERE TARGET_ID = " & TARGET_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CollTargerDetail = New Collection
End Sub
