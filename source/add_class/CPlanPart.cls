VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanPart"
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

Public PLAN_PART_ID As Long
Public PLAN_AREA As Long
Public PLAN_DATE As Date
Public PART_ITEM_ID As Long

Public PLAN_IN As Double
Public PLAN_OUT As Double

Public CANCEL_FLAG As String

Public MODIFY_DATE As Date

Public FROM_DATE As Date
Public TO_DATE As Date
 
 Public PART_NO As String
 Public PART_DESC As String
 
 Public NOTE As String
 
 Public FROM_PART_NO As String
 Public TO_PART_NO As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim WhereStr2 As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim GroupByStr As String
Dim MustCount As Boolean

   If Ind = 1 Then
      SelectStr1 = "SELECT PNP.*,PI.PART_NO,PI.PART_DESC "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PLAN_PART PNP "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PNP.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PNP.PART_ITEM_ID,PNP.PLAN_DATE,PNP.PLAN_AREA,PNP.PLAN_IN,PNP.PLAN_OUT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PLAN_PART PNP "
      FromStr = FromStr & " INNER JOIN "
      FromStr = FromStr & " ( "
      FromStr = FromStr & "      SELECT PNP2.PART_ITEM_ID,PNP2.PLAN_DATE,PNP2.PLAN_AREA,MAX(PNP2.MODIFY_DATE) MODIFY_DATE "
      FromStr = FromStr & "      FROM PLAN_PART PNP2 "
      
      If PLAN_AREA > 0 Then
         If Len(WhereStr2) <= 0 Then
            WhereStr2 = " WHERE (PNP2.PLAN_AREA = " & PLAN_AREA & ") "
         Else
            WhereStr2 = WhereStr2 & " AND (PNP2.PLAN_AREA = " & PLAN_AREA & ") "
         End If
      End If
   
      If FROM_DATE > 0 Then
         TempDate = DateToStringIntLow(Trim(FROM_DATE))
         If WhereStr2 = "" Then
            WhereStr2 = "WHERE (PNP2.PLAN_DATE >= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr2 = WhereStr2 & "AND (PNP2.PLAN_DATE >= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
      
      If TO_DATE > 0 Then
         TempDate = DateToStringIntHi(Trim(TO_DATE))
         If WhereStr2 = "" Then
            WhereStr2 = "WHERE (PNP2.PLAN_DATE <= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr2 = WhereStr2 & "AND (PNP2.PLAN_DATE <= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
      
      If Len(Trim(CANCEL_FLAG)) > 0 Then
         If WhereStr2 = "" Then
            WhereStr2 = "WHERE (PNP2.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
         Else
            WhereStr2 = WhereStr2 & "AND (PNP2.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
         End If
      End If
      
      FromStr = FromStr & "  " & WhereStr2
      FromStr = FromStr & "      GROUP BY PNP2.PART_ITEM_ID,PNP2.PLAN_DATE,PNP2.PLAN_AREA "
      FromStr = FromStr & " ) "
      FromStr = FromStr & " AS P ON (PNP.PLAN_AREA = P.PLAN_AREA AND PNP.PART_ITEM_ID = P.PART_ITEM_ID AND PNP.PLAN_DATE = P.PLAN_DATE AND PNP.MODIFY_DATE = P.MODIFY_DATE) "

   End If
   
   If PLAN_PART_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PNP.PLAN_PART_ID = " & PLAN_PART_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PNP.PLAN_PART_ID = " & PLAN_PART_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PNP.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PNP.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If PLAN_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PNP.PLAN_AREA = " & PLAN_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (PNP.PLAN_AREA = " & PLAN_AREA & ") "
      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PNP.PLAN_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PNP.PLAN_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PNP.PLAN_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PNP.PLAN_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   If Len(Trim(FROM_PART_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (PI.PART_NO >= '" & ChangeQuote(Trim(FROM_PART_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (PI.PART_NO >= '" & ChangeQuote(Trim(FROM_PART_NO)) & "')"
      End If
   End If
   
   If Len(Trim(TO_PART_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (PI.PART_NO <= '" & ChangeQuote(Trim(TO_PART_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (PI.PART_NO <= '" & ChangeQuote(Trim(TO_PART_NO)) & "')"
      End If
   End If
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (PNP.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (PNP.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      End If
   End If
   
   
'   If Ind = 2 Then
'      GroupByStr = "GROUP BY PNP.PART_ITEM_ID,PNP.PLAN_DATE "
'   ElseIf Ind = 4 Then
'      GroupByStr = "GROUP BY PART_ITEM_ID, PLAN_AREA "
'   ElseIf Ind = 5 Then
'      GroupByStr = "GROUP BY PART_ITEM_ID "
'   ElseIf Ind = 9 Then
'      GroupByStr = "GROUP BY PLAN_AREA "
'   ElseIf Ind = 10 Then
'      GroupByStr = "GROUP BY PART_ITEM_ID "
'   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PNP.PLAN_DATE "
      Else
         OrderByStr = "ORDER BY PNP.PLAN_DATE DESC,PNP.MODIFY_DATE DESC,PI.PART_NO "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 6 Then
      'OrderByStr = "ORDER BY PLAN_PART_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY PLAN_PART_ID "
   ElseIf Ind = 8 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOCATION_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LOCATION_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PART_DESC "
      Else
         OrderByStr = "ORDER BY LOCATION_NO "
      End If
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY PLAN_AREA "
   ElseIf Ind = 10 Then
      'OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   If OrderByStr = "" Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      PLAN_PART_ID = NVLI(Rs("PLAN_PART_ID"), 0)
      PLAN_DATE = InternalDateToDate(NVLS(Rs("PLAN_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PLAN_IN = NVLD(Rs("PLAN_IN"), 0)
      PLAN_OUT = NVLD(Rs("PLAN_OUT"), 0)
      PLAN_AREA = NVLI(Rs("PLAN_AREA"), 0)
      
      NOTE = NVLS(Rs("NOTE"), "")
      
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 2 Then
      PLAN_DATE = InternalDateToDate(NVLS(Rs("PLAN_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PLAN_AREA = NVLI(Rs("PLAN_AREA"), 0)
      
      PLAN_IN = NVLD(Rs("PLAN_IN"), 0)
      PLAN_OUT = NVLD(Rs("PLAN_OUT"), 0)
   End If
End Function
Public Function AddEditData()
On Error Resume Next
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String
Dim TempID8 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PLAN_PART_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PLAN_PART " & _
                     "( " & _
                     "   PLAN_PART_ID, " & _
                     "   PLAN_AREA, " & _
                     "   PLAN_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   PLAN_IN, " & _
                     "   PLAN_OUT, " & _
                     "   NOTE, " & _
                     "   CANCEL_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        PLAN_AREA & "," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(PLAN_DATE))) & "'," & _
                        PART_ITEM_ID & "," & _
                        PLAN_IN & "," & _
                        PLAN_OUT & "," & _
                        "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                        "'" & ChangeQuote(Trim(CANCEL_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PLAN_PART_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PLAN_PART SET " & _
                     "    PLAN_DATE = '" & ChangeQuote(DateToStringInt(Trim(PLAN_DATE))) & "', " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    PLAN_IN = " & PLAN_IN & ", " & _
                     "    PLAN_OUT = " & PLAN_OUT & ", " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "   CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PLAN_PART_ID = " & PLAN_PART_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PLAN_PART WHERE PLAN_PART_ID = " & PLAN_PART_ID
      m_Conn.Execute (SQL1)
End Function
Public Function UpdateCancelByDateArea()
Dim ServerDateTime As String
Dim SQL1 As String
            Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
            SQL1 = "UPDATE PLAN_PART SET " & _
                     "   CANCEL_FLAG = 'Y', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PLAN_DATE = '" & ChangeQuote(DateToStringInt(Trim(PLAN_DATE))) & "'" & " AND " & " PLAN_AREA = " & PLAN_AREA
                     
            m_Conn.Execute (SQL1)
End Function
