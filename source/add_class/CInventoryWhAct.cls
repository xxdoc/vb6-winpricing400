VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryWhAct"
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

Public INVENTORY_WH_ACT_ID As Long
Public INVENTORY_WH_ACT_AREA As Long
Public INVENTORY_WH_ACT_DATE As Date

Public INVENTORY_WH_ACT_DESC As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long


Public FROM_DATE As Date
Public TO_DATE As Date

Public PART_NO As String
Public PRODUCT_NO As String

Public CollRawMaterials As Collection
Public CollPhamacyRoom As Collection
Public CollSilo As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   FROM_DATE = -1
   TO_DATE = -1
   Set CollRawMaterials = New Collection
   Set CollPhamacyRoom = New Collection
   Set CollSilo = New Collection
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
      SelectStr1 = "SELECT IA.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM INVENTORY_WH_ACT IA "
   End If
   
   If INVENTORY_WH_ACT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IA.INVENTORY_WH_ACT_ID = " & INVENTORY_WH_ACT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (IA.INVENTORY_WH_ACT_ID = " & INVENTORY_WH_ACT_ID & ") "
      End If
   End If
 
   If INVENTORY_WH_ACT_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IA.INVENTORY_WH_ACT_AREA = " & INVENTORY_WH_ACT_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (IA.INVENTORY_WH_ACT_AREA = " & INVENTORY_WH_ACT_AREA & ") "
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If Len(PRODUCT_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PRODUCT_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PRODUCT_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PRODUCT_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM INVENTORY_WH_ACT_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.INVENTORY_WH_ACT_ID = IA.INVENTORY_WH_ACT_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PRODUCT_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IA.INVENTORY_WH_ACT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IA.INVENTORY_WH_ACT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IA.INVENTORY_WH_ACT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IA.INVENTORY_WH_ACT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY IA.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY IA.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IA.INVENTORY_WH_ACT_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IA.INVENTORY_WH_ACT_FROM "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY IA.INVENTORY_WH_ACT_TO "
      Else
         OrderByStr = "ORDER BY IA.INVENTORY_WH_ACT_DATE "
      End If
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
      INVENTORY_WH_ACT_ID = NVLI(Rs("INVENTORY_WH_ACT_ID"), -1)
      INVENTORY_WH_ACT_AREA = NVLI(Rs("INVENTORY_WH_ACT_AREA"), -1)
      INVENTORY_WH_ACT_DATE = InternalDateToDate(NVLS(Rs("INVENTORY_WH_ACT_DATE"), ""))
      
      INVENTORY_WH_ACT_DESC = NVLS(Rs("INVENTORY_WH_ACT_DESC"), "")
      
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
   
      Call glbDatabaseMngr.GetSeqID("INVENTORY_WH_ACT_SEQ", ID, glbErrorLog) 'INVENTORY_WH_ACT_ID
      
      SQL1 = "INSERT INTO INVENTORY_WH_ACT " & _
                     "( " & _
                     "   INVENTORY_WH_ACT_ID, " & _
                     "   INVENTORY_WH_ACT_AREA, " & _
                     "   INVENTORY_WH_ACT_DATE, " & _
                     "   INVENTORY_WH_ACT_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        ID & "," & _
                        "" & INVENTORY_WH_ACT_AREA & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(INVENTORY_WH_ACT_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(INVENTORY_WH_ACT_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING INVENTORY_WH_ACT_ID "
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_WH_ACT SET " & _
                     "   INVENTORY_WH_ACT_DATE= '" & ChangeQuote(Trim(DateToStringInt(INVENTORY_WH_ACT_DATE))) & "', " & _
                      "  INVENTORY_WH_ACT_DESC = '" & ChangeQuote(Trim(INVENTORY_WH_ACT_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_WH_ACT_ID = " & INVENTORY_WH_ACT_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
     INVENTORY_WH_ACT_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM INVENTORY_WH_ACT_ITEM WHERE INVENTORY_WH_ACT_ID = " & INVENTORY_WH_ACT_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM INVENTORY_WH_ACT WHERE INVENTORY_WH_ACT_ID = " & INVENTORY_WH_ACT_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CollRawMaterials = New Collection
   Set CollPhamacyRoom = New Collection
   Set CollSilo = New Collection
End Sub
