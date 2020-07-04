VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanning"
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

Public PLANNING_ID As Long
Public PLANNING_AREA As Long
Public PLANNING_FROM As Date
Public PLANNING_TO As Date
Public PLANNING_DATE As Date
Public PLANNING_DESC As String
Public PLAN_VERSION As Long

Public PLANNING_ITEM_ID As Long
Public PLANNING_SUB_TYPE As Long
Public PART_ITEM_ID As Long
Public PART_TYPE_GROUP As String
Public PART_GROUP_ID As Long
Public PLAN_AMOUNT As Double

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As Long

Public FROM_DATE As Date
Public TO_DATE As Date

Public PART_NO As String
Public PRODUCT_NO As String

Public CollPartUse As Collection
Public CollProductGet As Collection
Public CollPartSup As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   FROM_DATE = -1
   TO_DATE = -1
   
    
   Set CollPartUse = New Collection
   Set CollProductGet = New Collection
   Set CollPartSup = New Collection
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
      SelectStr1 = "SELECT PN.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM PLANNING PN "
   ElseIf Ind = 2 Then
       SelectStr1 = "SELECT MAX(PN.PLAN_VERSION)  PLAN_VERSION "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM PLANNING PN "
   ElseIf Ind = 3 Then
       SelectStr1 = "SELECT PN.PLANNING_ID,PN.PLANNING_DATE,PN.PLANNING_FROM,PN.PLANNING_TO,PNI.PART_ITEM_ID,PNI.PLAN_AMOUNT,PN.PLAN_VERSION "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM PLANNING PN "
      FromStr = FromStr & "LEFT OUTER JOIN PLANNING_ITEM PNI ON (PN.PLANNING_ID = PNI.PLANNING_ID)  "
   End If
   
   If PLANNING_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PN.PLANNING_ID = " & PLANNING_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_ID = " & PLANNING_ID & ") "
      End If
   End If
 
   If PLANNING_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PN.PLANNING_AREA = " & PLANNING_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_AREA = " & PLANNING_AREA & ") "
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If Len(PRODUCT_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PRODUCT_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PRODUCT_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PRODUCT_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM PLANNING_ITEM PNI LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (PNI.PLANNING_ID = PN.PLANNING_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PRODUCT_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PN.PLANNING_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PN.PLANNING_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY PN.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
    ' GroupByStr = "GROUP BY PN.PLANNING_ID,PN.PLANNING_DATE,PN.PLANNING_FROM,PN.PLANNING_TO,PNI.PART_ITEM_ID,PNI.PLAN_AMOUNT "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PN.PLANNING_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PN.PLANNING_FROM "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PN.PLANNING_TO "
      Else
         OrderByStr = "ORDER BY PN.PLANNING_DATE DESC,PN.PLAN_VERSION "
      End If
   ElseIf Ind = 3 Then
       OrderByStr = "ORDER BY PN.PLANNING_DATE,PN.PLAN_VERSION "
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
      PLANNING_ID = NVLI(Rs("PLANNING_ID"), -1)
      PLANNING_AREA = NVLI(Rs("PLANNING_AREA"), -1)
      
      PLANNING_FROM = InternalDateToDate(NVLS(Rs("PLANNING_FROM"), ""))
      PLANNING_TO = InternalDateToDate(NVLS(Rs("PLANNING_TO"), ""))
      PLANNING_DATE = InternalDateToDate(NVLS(Rs("PLANNING_DATE"), ""))
      
      PLANNING_DESC = NVLS(Rs("PLANNING_DESC"), "")
      PLAN_VERSION = NVLI(Rs("PLAN_VERSION"), 0)
   ElseIf Ind = 2 Then
      PLAN_VERSION = NVLI(Rs("PLAN_VERSION"), 0)
   ElseIf Ind = 3 Then
      PLANNING_ID = NVLI(Rs("PLANNING_ID"), -1)
'      PLANNING_AREA = NVLI(Rs("PLANNING_AREA"), -1)
      
      PLANNING_FROM = InternalDateToDate(NVLS(Rs("PLANNING_FROM"), ""))
      PLANNING_TO = InternalDateToDate(NVLS(Rs("PLANNING_TO"), ""))
      PLANNING_DATE = InternalDateToDate(NVLS(Rs("PLANNING_DATE"), ""))
      
'      PLANNING_DESC = NVLS(Rs("PLANNING_DESC"), "")
      PLAN_VERSION = NVLI(Rs("PLAN_VERSION"), 0)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PLAN_AMOUNT = NVLD(Rs("PLAN_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
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
      SQL1 = "INSERT INTO PLANNING " & _
                     "( " & _
                     "   PLANNING_AREA, " & _
                     "   PLANNING_FROM, " & _
                     "   PLANNING_TO, " & _
                     "   PLANNING_DATE, " & _
                     "   PLANNING_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY ," & _
                    "    PLAN_VERSION " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        PLANNING_AREA & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PLANNING_FROM))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PLANNING_TO))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PLANNING_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(PLANNING_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "" & PLAN_VERSION & _
                     ")  RETURNING PLANNING_ID "
   Else 'Show edit
      SQL1 = "UPDATE PLANNING SET " & _
                     "   PLANNING_FROM= '" & ChangeQuote(Trim(DateToStringInt(PLANNING_FROM))) & "', " & _
                     "   PLANNING_TO= '" & ChangeQuote(Trim(DateToStringInt(PLANNING_TO))) & "', " & _
                     "   PLANNING_DATE= '" & ChangeQuote(Trim(DateToStringInt(PLANNING_DATE))) & "', " & _
                      "  PLANNING_DESC = '" & ChangeQuote(Trim(PLANNING_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PLANNING_ID = " & PLANNING_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      PLANNING_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PLANNING_ITEM WHERE PLANNING_ID = " & PLANNING_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM PLANNING WHERE PLANNING_ID = " & PLANNING_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CollPartUse = New Collection
   Set CollProductGet = New Collection
   Set CollPartSup = New Collection
End Sub
