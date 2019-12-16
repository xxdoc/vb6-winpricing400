VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPackProductionItem"
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

Public PACK_PRODUCTION_ITEM_ID As Long
Public PACK_PRODUCTION_ID As Long
Public PART_ITEM_ID As Long
Public PART_DESC As String
Public TX_AMOUNT As Double
Public WEIGHT_PER_PACK As Long
Public PACK_AMOUNT As Double
Public PALLET_LABEL_YELLOW As String
Public PALLET_LABEL_GREEN As String
Public SEWING_THREAD As Long
Public NOTE As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public KEY_ID As Long
Public KEY_LOOKUP As String

'Public CostItems As Collection
'Public CostItemsEx As Collection
'Public ExpenseItem As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
'   Set CostItems = New Collection
'   Set ExpenseItem = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String

If Ind = 1 Then
   SelectStr1 = "SELECT PPI.*,PI.PART_DESC  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PACK_PRODUCTION_ITEM PPI "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PPI.PART_ITEM_ID = PI.PART_ITEM_ID) "
   WhereStr = ""
End If
   
   If PACK_PRODUCTION_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PACK_PRODUCTION_ITEM_ID  = " & PACK_PRODUCTION_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PACK_PRODUCTION_ITEM_ID  = " & PACK_PRODUCTION_ITEM_ID & ") "
      End If
   End If
   
   If PACK_PRODUCTION_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PACK_PRODUCTION_ITEM_ID  = " & PACK_PRODUCTION_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PACK_PRODUCTION_ITEM_ID  = " & PACK_PRODUCTION_ITEM_ID & ") "
      End If
   End If
   
  If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID  = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID  = " & PART_ITEM_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PACK_PRODUCTION_ITEM_ID  "
   Else
      OrderByStr = "ORDER BY PACK_PRODUCTION_ITEM_ID  "
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
   PACK_PRODUCTION_ITEM_ID = NVLI(Rs("PACK_PRODUCTION_ITEM_ID"), -1)
   PACK_PRODUCTION_ID = NVLI(Rs("PACK_PRODUCTION_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   WEIGHT_PER_PACK = NVLI(Rs("WEIGHT_PER_PACK"), -1)
   PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
   PALLET_LABEL_YELLOW = NVLS(Rs("PALLET_LABEL_YELLOW"), "")
   PALLET_LABEL_GREEN = NVLS(Rs("PALLET_LABEL_GREEN"), "")
   SEWING_THREAD = NVLI(Rs("SEWING_THREAD"), -1)
   NOTE = NVLS(Rs("NOTE"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PACK_PRODUCTION_ITEM_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PACK_PRODUCTION_ITEM " & _
                     "(  PACK_PRODUCTION_ITEM_ID, " & _
                     "   PACK_PRODUCTION_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   TX_AMOUNT, " & _
                     "   WEIGHT_PER_PACK, " & _
                     "   PACK_AMOUNT, " & _
                     "   PALLET_LABEL_YELLOW, " & _
                     "   PALLET_LABEL_GREEN, " & _
                     "   SEWING_THREAD, " & _
                     "   NOTE, " & _
                     "   CREATE_DATE,CREATE_BY,MODIFY_DATE,MODIFY_BY )"
      SQL1 = SQL1 & "VALUES " & _
                     "( " & ID & "," & _
                     "" & PACK_PRODUCTION_ID & "," & _
                     "" & PART_ITEM_ID & "," & _
                     "" & TX_AMOUNT & "," & _
                     "" & WEIGHT_PER_PACK & "," & _
                     "" & PACK_AMOUNT & "," & _
                     "'" & ChangeQuote(Trim(PALLET_LABEL_YELLOW)) & "'," & _
                     "'" & ChangeQuote(Trim(PALLET_LABEL_GREEN)) & "'," & _
                     "" & SEWING_THREAD & "," & _
                     "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                     "'" & ServerDateTime & "'," & glbUser.USER_ID & ",'" & ServerDateTime & "'," & glbUser.USER_ID & ") "
                         
         PACK_PRODUCTION_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PACK_PRODUCTION_ITEM SET " & _
                     "   PACK_PRODUCTION_ID= " & PACK_PRODUCTION_ID & ", " & _
                     "   PART_ITEM_ID= " & PART_ITEM_ID & ", " & _
                     "   TX_AMOUNT= " & TX_AMOUNT & ", " & _
                     "   WEIGHT_PER_PACK= " & WEIGHT_PER_PACK & ", " & _
                     "   PACK_AMOUNT= " & PACK_AMOUNT & ", " & _
                     "   PALLET_LABEL_YELLOW = '" & ChangeQuote(Trim(PALLET_LABEL_YELLOW)) & "', " & _
                     "   PALLET_LABEL_GREEN = '" & ChangeQuote(Trim(PALLET_LABEL_GREEN)) & "', " & _
                     "   SEWING_THREAD= " & SEWING_THREAD & ", " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PACK_PRODUCTION_ITEM_ID = " & PACK_PRODUCTION_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PACK_PRODUCTION_ITEM WHERE PACK_PRODUCTION_ITEM_ID = " & PACK_PRODUCTION_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
'   Set ExpenseItem = Nothing
End Sub
