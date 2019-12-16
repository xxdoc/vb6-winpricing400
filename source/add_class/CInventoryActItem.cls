VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryActItem"
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

Public INVENTORY_ACT_ITEM_ID As Long
Public INVENTORY_ACT_ID As Long
Public PART_ITEM_ID As Long
Public INVENTORY_ACT_ITEM_TYPE As Long
Public INVENTORY_ACT_AMOUNT As Double
Public NOTE As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date

Public PART_NO As String
Public PART_DESC As String

Public INVENTORY_ACT_AREA As Long
Public INVENTORY_ACT_DATE As Date
Public INVENTORY_ACT_DESC As String
Public UNIT_NAME As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
      SelectStr1 = "SELECT IAI.*,IA.* "
      SelectStr1 = SelectStr1 & ",PI.PART_NO,PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM INVENTORY_ACT_ITEM IAI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_ACT IA ON (IAI.INVENTORY_ACT_ID = IA.INVENTORY_ACT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (IAI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   End If
   
   If INVENTORY_ACT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IAI.INVENTORY_ACT_ITEM_ID = " & INVENTORY_ACT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (IAI.INVENTORY_ACT_ITEM_ID = " & INVENTORY_ACT_ITEM_ID & ") "
      End If
   End If
   
   If INVENTORY_ACT_ITEM_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IAI.INVENTORY_ACT_ITEM_TYPE = " & INVENTORY_ACT_ITEM_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (IAI.INVENTORY_ACT_TYPE = " & INVENTORY_ACT_ITEM_TYPE & ") "
      End If
   End If
   
   If INVENTORY_ACT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IAI.INVENTORY_ACT_ID = " & INVENTORY_ACT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (IAI.INVENTORY_ACT_ID = " & INVENTORY_ACT_ID & ") "
      End If
   End If
   
   If INVENTORY_ACT_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PN.INVENTORY_ACT_AREA = " & INVENTORY_ACT_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (PN.INVENTORY_ACT_AREA = " & INVENTORY_ACT_AREA & ") "
      End If
   End If
   
   If INVENTORY_ACT_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(INVENTORY_ACT_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IA.INVENTORY_ACT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IA.INVENTORY_ACT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If FROM_DATE > -1 Then
       TempDate = DateToStringIntLow(Trim(FROM_DATE))
       If WhereStr = "" Then
          WhereStr = "WHERE (IA.INVENTORY_ACT_DATE >= '" & ChangeQuote(TempDate) & "')"
       Else
          WhereStr = WhereStr & "AND (IA.INVENTORY_ACT_DATE >= '" & ChangeQuote(TempDate) & "')"
       End If
    End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IA.INVENTORY_ACT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IA.INVENTORY_ACT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY IAI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY IAI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 4 Then
        OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PI.PART_DESC "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY IAI.INVENTORY_ACT_AMOUNT "
      Else
         OrderByStr = "ORDER BY IAI.INVENTORY_ACT_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PI.PART_NO "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      INVENTORY_ACT_ITEM_ID = NVLI(Rs("INVENTORY_ACT_ITEM_ID"), -1)
      INVENTORY_ACT_ID = NVLI(Rs("INVENTORY_ACT_ID"), -1)
      INVENTORY_ACT_ITEM_TYPE = NVLI(Rs("INVENTORY_ACT_ITEM_TYPE"), -1)
      INVENTORY_ACT_AMOUNT = NVLD(Rs("INVENTORY_ACT_AMOUNT"), 0)
      INVENTORY_ACT_AMOUNT = NVLD(Rs("INVENTORY_ACT_AMOUNT"), 0)
      INVENTORY_ACT_AREA = NVLI(Rs("INVENTORY_ACT_AREA"), -1)
      INVENTORY_ACT_DATE = InternalDateToDate(NVLS(Rs("INVENTORY_ACT_DATE"), "")) 'NVLS(Rs("INVENTORY_ACT_DATE"), "")
      INVENTORY_ACT_DESC = NVLS(Rs("INVENTORY_ACT_DESC"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      PART_ITEM_ID = NVLS(Rs("PART_ITEM_ID"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
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

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   Call glbDatabaseMngr.GetSeqID("INVENTORY_ACT_ITEM_SEQ", ID, glbErrorLog)
      
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO INVENTORY_ACT_ITEM " & _
                     "( " & _
                     "   INVENTORY_ACT_ITEM_ID, " & _
                     "   INVENTORY_ACT_ID, " & _
                     "   INVENTORY_ACT_ITEM_TYPE, " & _
                     "   PART_ITEM_ID, " & _
                     "   INVENTORY_ACT_AMOUNT, " & _
                     "   NOTE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        ID & "," & _
                        INVENTORY_ACT_ID & "," & _
                        INVENTORY_ACT_ITEM_TYPE & "," & _
                        TempID1 & "," & _
                        INVENTORY_ACT_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING INVENTORY_ACT_ITEM_ID "
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_ACT_ITEM SET " & _
                     "  PART_ITEM_ID = " & TempID1 & ", " & _
                     "  INVENTORY_ACT_AMOUNT = " & INVENTORY_ACT_AMOUNT & ", " & _
                      "  NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_ACT_ITEM_ID = " & INVENTORY_ACT_ITEM_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      INVENTORY_ACT_ITEM_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM INVENTORY_ACT_ITEM WHERE INVENTORY_ACT_ITEM_ID = " & INVENTORY_ACT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

