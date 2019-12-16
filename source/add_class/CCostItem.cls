VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostItem"
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

Public COST_ITEM_ID As Long
Public COSTPRD_ITEM_ID As Long
Public PARAM_PROCESS_ID As Long
Public ITEM_COST As Double

Public PARAMETER_PROCESS_NAME  As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As String

Public CostItemRaws As Collection

'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CostItemRaws = New Collection
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

   SelectStr1 = "SELECT INP.*, PI.PARAMETER_PROCESS_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM COST_ITEM INP "
   FromStr = FromStr & "LEFT OUTER JOIN PARAMETER_PROCESS PI ON (PI.PARAMETER_PROCESS_ID= INP.PARAM_PROCESS_ID ) "
   
   If COST_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COST_ITEM_ID = " & COST_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COST_ITEM_ID = " & COST_ITEM_ID & ") "
      End If
   End If
   
   If COSTPRD_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY COST_ITEM_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY COST_ITEM_ID "
   Else
      OrderByStr = "ORDER BY COST_ITEM_ID "
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
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   COST_ITEM_ID = NVLI(Rs("COST_ITEM_ID"), -1)
   COSTPRD_ITEM_ID = NVLI(Rs("COSTPRD_ITEM_ID"), -1)
   PARAM_PROCESS_ID = NVLI(Rs("PARAM_PROCESS_ID"), -1)
   ITEM_COST = NVLD(Rs("ITEM_COST"), 0)
   
   PARAMETER_PROCESS_NAME = NVLS(Rs("PARAMETER_PROCESS_NAME"), "")
  End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = PARAM_PROCESS_ID
   If PARAM_PROCESS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("COST_ITEM_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO COST_ITEM " & _
                     "( " & _
                     "   COST_ITEM_ID, " & _
                     "   COSTPRD_ITEM_ID, " & _
                     "   PARAM_PROCESS_ID, " & _
                     "   ITEM_COST, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                             "'" & COSTPRD_ITEM_ID & "'," & _
                             TempID1 & "," & _
                        ITEM_COST & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         COST_ITEM_ID = id
   Else 'Show edit
                        SQL1 = "UPDATE COST_ITEM SET " & _
                        "  PARAM_PROCESS_ID = " & TempID1 & ", " & _
                     "   ITEM_COST= " & ITEM_COST & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE COST_ITEM_ID = " & COST_ITEM_ID
   End If
   
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM COST_ITEM WHERE COST_ITEM_ID = " & COST_ITEM_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM CSTITEM_RAW WHERE COST_ITEM_ID = " & COST_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CostItemRaws = Nothing
End Sub
