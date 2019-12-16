VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAlertBox"
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
   
Public ALERT_BOX_ID As Long
Public ALERT_BOX_FROM As Date
Public ALERT_BOX_TO As Date
Public ALERT_BOX_TYPE As Long
Public ALERT_BOX_DESC As String
Public ALERT_ALL_FLAG  As String
Public ALERT_CANCEL_FLAG As String

Public ALERT_DATE_SEARCH As Date

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public USER_NAME As String

Public CollAlertDetail As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   ALERT_DATE_SEARCH = -1
   
   Set CollAlertDetail = New Collection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long, Optional MustCount As Boolean = False)
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
      SelectStr1 = "SELECT AB.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM ALERT_BOX AB "
   End If
   
   If ALERT_BOX_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AB.ALERT_BOX_ID = " & ALERT_BOX_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AB.ALERT_BOX_ID = " & ALERT_BOX_ID & ") "
      End If
   End If
 
   If ALERT_BOX_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AB.ALERT_BOX_TYPE = " & ALERT_BOX_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (AB.ALERT_BOX_TYPE = " & ALERT_BOX_TYPE & ") "
      End If
   End If
   
   If Len(ALERT_CANCEL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AB.ALERT_CANCEL_FLAG = '" & ALERT_CANCEL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (AB.ALERT_CANCEL_FLAG = '" & ALERT_CANCEL_FLAG & "') "
      End If
   End If
   
   If Len(USER_NAME) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(USER_NAME, SubLen, NewStr) Then
            WhereStr = "WHERE ((AB.ALERT_ALL_FLAG='Y') OR ((SELECT COUNT(*) FROM ALERT_DETAIL ALI LEFT OUTER JOIN USER_ACCOUNT PI ON (ALI.USER_ID = PI.USER_ID) WHERE (ALI.ALERT_BOX_ID = AB.ALERT_BOX_ID) AND (SUBSTR(PI.USER_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)) "
         Else
            WhereStr = "WHERE ((AB.ALERT_ALL_FLAG='Y') OR ((SELECT COUNT(*) FROM ALERT_DETAIL ALI LEFT OUTER JOIN USER_ACCOUNT PI ON (ALI.USER_ID = PI.USER_ID) WHERE (ALI.ALERT_BOX_ID = AB.ALERT_BOX_ID) AND (PI.USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "')) > 0)) "
         End If
      Else
         If WildCard(USER_NAME, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((AB.ALERT_ALL_FLAG='Y') OR ((SELECT COUNT(*) FROM ALERT_DETAIL ALI LEFT OUTER JOIN USER_ACCOUNT PI ON (ALI.USER_ID = PI.USER_ID) WHERE (ALI.ALERT_BOX_ID = AB.ALERT_BOX_ID) AND (SUBSTR(PI.USER_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0)) "
         Else
            WhereStr = WhereStr & "AND ((AB.ALERT_ALL_FLAG='Y') OR ((SELECT COUNT(*) FROM ALERT_DETAIL ALI LEFT OUTER JOIN USER_ACCOUNT PI ON (ALI.USER_ID = PI.USER_ID) WHERE (ALI.ALERT_BOX_ID = AB.ALERT_BOX_ID) AND (PI.USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "')) > 0))"
         End If
      End If
   End If
   
   If ALERT_DATE_SEARCH > -1 Then
      TempDate = DateToStringIntLow(Trim(ALERT_DATE_SEARCH))
      If WhereStr = "" Then
         WhereStr = "WHERE (AB.ALERT_BOX_FROM <= '" & ChangeQuote(TempDate) & "' AND " & "AB.ALERT_BOX_TO >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (AB.ALERT_BOX_FROM <= '" & ChangeQuote(TempDate) & "' AND " & "AB.ALERT_BOX_TO >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY AB.USER_ACCOUNT_ID, PI.USER_NAME, PI.PART_DESC "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY AB.USER_ACCOUNT_ID, PI.USER_NAME, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY AB.ALERT_BOX_FROM "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY AB.ALERT_BOX_FROM "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY AB.ALERT_BOX_TO "
      Else
         OrderByStr = "ORDER BY AB.ALERT_BOX_ID "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      ALERT_BOX_ID = NVLI(Rs("ALERT_BOX_ID"), -1)
      ALERT_BOX_TYPE = NVLI(Rs("ALERT_BOX_TYPE"), -1)
      
      ALERT_BOX_FROM = InternalDateToDate(NVLS(Rs("ALERT_BOX_FROM"), ""))
      ALERT_BOX_TO = InternalDateToDate(NVLS(Rs("ALERT_BOX_TO"), ""))
      
      ALERT_BOX_DESC = NVLS(Rs("ALERT_BOX_DESC"), "")
      
      ALERT_ALL_FLAG = NVLS(Rs("ALERT_ALL_FLAG"), "")
      ALERT_CANCEL_FLAG = NVLS(Rs("ALERT_CANCEL_FLAG"), "")
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

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO ALERT_BOX " & _
                     "( " & _
                     "   ALERT_BOX_TYPE, " & _
                     "   ALERT_BOX_FROM, " & _
                     "   ALERT_BOX_TO, " & _
                     "   ALERT_BOX_DESC, " & _
                     "   ALERT_ALL_FLAG, " & _
                     "   ALERT_CANCEL_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        ALERT_BOX_TYPE & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ALERT_BOX_FROM))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ALERT_BOX_TO))) & "'," & _
                        "'" & ChangeQuote(Trim(ALERT_BOX_DESC)) & "'," & _
                        "'" & ChangeQuote(Trim(ALERT_ALL_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(ALERT_CANCEL_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING ALERT_BOX_ID "
   Else 'Show edit
      SQL1 = "UPDATE ALERT_BOX SET " & _
                     "   ALERT_BOX_FROM= '" & ChangeQuote(Trim(DateToStringInt(ALERT_BOX_FROM))) & "', " & _
                     "   ALERT_BOX_TO= '" & ChangeQuote(Trim(DateToStringInt(ALERT_BOX_TO))) & "', " & _
                     "  ALERT_BOX_DESC = '" & ChangeQuote(Trim(ALERT_BOX_DESC)) & "', " & _
                     "  ALERT_ALL_FLAG = '" & ChangeQuote(Trim(ALERT_ALL_FLAG)) & "', " & _
                     "  ALERT_CANCEL_FLAG = '" & ChangeQuote(Trim(ALERT_CANCEL_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ALERT_BOX_ID = " & ALERT_BOX_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      ALERT_BOX_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM ALERT_DETAIL WHERE ALERT_BOX_ID = " & ALERT_BOX_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM ALERT_BOX WHERE ALERT_BOX_ID = " & ALERT_BOX_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CollAlertDetail = New Collection
End Sub
