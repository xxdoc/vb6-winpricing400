VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostPrdItem"
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

Public COSTPRD_ITEM_ID As Long
Public COST_PRODUCTION_ID As Long
Public PART_ITEM_ID As Long
Public COST_AMOUNT As Double 'ปริมาณที่ใช้
Public RAW_AMOUNT As Double 'มูลค่าวัตถุดิบที่ใช้

Public PARAMETER_PROCESS_NAME  As String
Public PART_NO As String
Public PART_DESC As String
Public EXPENSE_AMOUNT As Double

Public PARCEL_TYPE As PARCEL_TYPE
Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As String

Public CostItems As Collection

'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CostItems = New Collection
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
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT INP.*, PIT.PART_NO, PIT.PART_DESC, PIT.PARCEL_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM COSTPRD_ITEM INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PIT ON (PIT.PART_ITEM_ID= INP.PART_ITEM_ID ) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT INP.PART_ITEM_ID, SUM(INP.COST_AMOUNT) COST_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM COSTPRD_ITEM INP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PIT ON (PIT.PART_ITEM_ID= INP.PART_ITEM_ID ) "
   End If
   
   If COSTPRD_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID & ") "
      End If
   End If
   
   If COST_PRODUCTION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY INP.PART_ITEM_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY COSTPRD_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY COSTPRD_ITEM_ID "
      Else
         OrderByStr = "ORDER BY COSTPRD_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PART_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   COSTPRD_ITEM_ID = NVLI(Rs("COSTPRD_ITEM_ID"), -1)
   COST_PRODUCTION_ID = NVLI(Rs("COST_PRODUCTION_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   COST_AMOUNT = NVLD(Rs("COST_AMOUNT"), 0)
   EXPENSE_AMOUNT = NVLD(Rs("EXPENSE_AMOUNT"), 0)
   RAW_AMOUNT = NVLD(Rs("RAW_AMOUNT"), 0)
   
   PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), -1)
   PART_NO = NVLS(Rs("PART_NO"), "")
   PART_DESC = NVLS(Rs("PART_DESC"), "")
  End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("COSTPRD_ITEM_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO COSTPRD_ITEM " & _
                     "( " & _
                     "   COSTPRD_ITEM_ID, " & _
                     "   COST_PRODUCTION_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   COST_AMOUNT, " & _
                     "   EXPENSE_AMOUNT, " & _
                     "   RAW_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                             "'" & COST_PRODUCTION_ID & "'," & _
                             TempID1 & "," & _
                        COST_AMOUNT & "," & _
                        EXPENSE_AMOUNT & "," & _
                        RAW_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         COSTPRD_ITEM_ID = id
   Else 'Show edit
                        SQL1 = "UPDATE COSTPRD_ITEM SET " & _
                        "  PART_ITEM_ID = " & TempID1 & ", " & _
                     "   COST_AMOUNT= " & COST_AMOUNT & ", " & _
                     "   EXPENSE_AMOUNT= " & EXPENSE_AMOUNT & ", " & _
                     "   RAW_AMOUNT= " & RAW_AMOUNT & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID
   End If
   
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM CSTITEM_RAW WHERE COST_ITEM_ID IN (SELECT COST_ITEM_ID FROM COST_ITEM WHERE COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID & ") "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM COST_ITEM WHERE COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM COSTPRD_ITEM WHERE COSTPRD_ITEM_ID = " & COSTPRD_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CostItems = Nothing
End Sub
