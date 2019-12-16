VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExWorksPriceItem"
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

Public EX_WORKS_PRICE_ITEM_ID As Long
Public EX_WORKS_PRICE_ID As Long
Public PART_ITEM_ID As Long
Public RATE_TYPE As Long
Public PACKAGE_RATE As Double
Public SocPartType As Long

Public EX_WORKS_PRICE_CODE As String
Public EX_WORKS_PRICE_DESC As String
Public EX_WORKS_PRICE_LEVEL As Long
Public PART_TYPE As Long
Public PART_NO As String
Public PART_DESC As String

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Public SUBSCRIBER_ID As Long
Public ACCOUNT_ID As Long
Public ENABLE_FLAG As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
'Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
'Dim SQL1 As String
'Dim SQL2 As String
'Dim SelectStr1 As String
'Dim SelectStr2 As String
'Dim FromStr As String
'Dim WhereStr As String
'Dim OrderByStr As String
'Dim OrderTypeStr As String
'Dim SubLen As Long
'Dim NewStr As String
'
'   SelectStr1 = "SELECT LT.*, FT.FEATURE_CODE, FT.FEATURE_DESC, FT.FEATURE_LVEL, UTT.UNTTYPE_ID, UT.UNIT_NAME, RT.RTTYPE_NAME  "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM SOC_FEATURE LT, FEATURE FT, UNIT_TYPE UTT, UNIT UT, RATE_TYPE RT "
'
'   WhereStr = "WHERE (FT.FEATURE_ID = LT.FEATURE_ID) AND (UT.UNIT_ID = LT.UNIT) AND (UTT.UNTTYPE_ID = UT.UNIT_TYPE) AND (RT.RTTYPE_ID = LT.RATE_TYPE) "
'   If SOC_FEATURE_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
'      End If
'   End If
'
'   If SOC_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (SOC_ID = " & SOC_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (SOC_ID = " & SOC_ID & ") "
'      End If
'   End If
'
'      If Len(Trim(FROM_CUSTOMER_CODE)) > 0 Then
'      If WildCard(FROM_CUSTOMER_CODE, SubLen, NewStr) Then
'         If WhereStr = "" Then
'            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
'         Else
'            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
'         End If
'      Else
'         If WhereStr = "" Then
'            WhereStr = "WHERE (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
'         Else
'            WhereStr = WhereStr & "AND (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
'         End If
'      End If
'   End If
'
'   If Len(Trim(TO_CUSTOMER_CODE)) > 0 Then
'      If WildCard(TO_CUSTOMER_CODE, SubLen, NewStr) Then
'         If WhereStr = "" Then
'            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
'         Else
'            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
'         End If
'      Else
'         If WhereStr = "" Then
'            WhereStr = "WHERE (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
'         Else
'            WhereStr = WhereStr & "AND (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
'         End If
'      End If
'   End If
'
'   OrderByStr = ""
'   OrderTypeStr = ""
'   If OrderBy > 0 Then
'      If OrderBy = 1 Then
'         OrderByStr = "ORDER BY SOC_FEATURE_ID "
'      ElseIf OrderBy = 2 Then
'         OrderByStr = "ORDER BY SOC_FEATURE_ID "
'      ElseIf OrderBy = 3 Then
'         OrderByStr = "ORDER BY SOC_FEATURE_ID "
'      Else
'         OrderByStr = "ORDER BY SOC_FEATURE_ID "
'      End If
'
'      If OrderType = 1 Then
'         OrderTypeStr = "DESC "
'      ElseIf OrderType = 2 Then
'         OrderTypeStr = "DESC "
'      Else
'         OrderTypeStr = "DESC "
'      End If
'   End If
'
'   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
'   SQL2 = SelectStr2 & FromStr & WhereStr
'
'   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'End Function
'
'Public Function QueryData3(Rs As ADODB.Recordset, ItemCount As Long)
'Dim SQL1 As String
'Dim SQL2 As String
'Dim SelectStr1 As String
'Dim SelectStr2 As String
'Dim FromStr As String
'Dim WhereStr As String
'Dim OrderByStr As String
'Dim OrderTypeStr As String
'
'   SelectStr1 = "SELECT SB.ACCOUNT_ID, AGR.SUBSCRIBER_ID, AGR.SOC_ID, AGR.EXPIRE_DATE, AGR.EFFECTIVE_DATE, SF.* "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM AGREEMENT AGR "
'   FromStr = FromStr & "LEFT OUTER JOIN SUBSCRIBER SB ON (AGR.SUBSCRIBER_ID = SB.SUBSCRIBER_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (SB.ACCOUNT_ID = AC.ACCOUNT_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (AGR.SOC_ID = SC.SOC_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN SOC_FEATURE SF ON (SC.SOC_ID = SF.SOC_ID) "
'
'   WhereStr = "WHERE ((AC.ACCOUNT_ID = " & ACCOUNT_ID & " AND AGR.SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") OR (AC.ACCOUNT_ID = " & ACCOUNT_ID & " AND AGR.SUBSCRIBER_ID <> " & SUBSCRIBER_ID & " AND SB.DUMMY_FLAG = 'Y')) AND ((SF.FEATURE_ID = " & FEATURE_ID & ") OR SF.PART_ITEM_ID = " & PART_ITEM_ID & ") AND (AGR.SOC_ID IS NOT NULL) "
'
''   If FEATURE_ID > 0 Then
''      If Len(WhereStr) <= 0 Then
''         WhereStr = "WHERE (FEATURE_ID = " & FEATURE_ID & ") "
''      Else
''         WhereStr = WhereStr & "AND (FEATURE_ID = " & FEATURE_ID & ") "
''      End If
''   End If
''
''   If PART_ITEM_ID > 0 Then
''      If Len(WhereStr) <= 0 Then
''         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
''      Else
''         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
''      End If
''   End If
''
''   If ACCOUNT_ID > 0 Then
''      If Len(WhereStr) <= 0 Then
''         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
''      Else
''         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
''      End If
''   End If
''
''   If SUBSCRIBER_ID > 0 Then
''      If Len(WhereStr) <= 0 Then
''         WhereStr = "WHERE (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
''      Else
''         WhereStr = WhereStr & "AND (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
''      End If
''   End If
'   OrderByStr = ""
'   OrderTypeStr = ""
'
'   If OrderBy = 1 Then
'      OrderByStr = "ORDER BY ACCOUNT_ID, SB.DUMMY_FLAG ASC, AGR.EXPIRE_DATE DESC, AGR.EFFECTIVE_DATE "
'   ElseIf OrderBy = 2 Then
'      OrderByStr = "ORDER BY SB.DUMMY_FLAG ASC, AGR.EXPIRE_DATE DESC, AGR.EFFECTIVE_DATE "
'   Else
'      OrderByStr = "ORDER BY SB.DUMMY_FLAG ASC, AGR.EXPIRE_DATE DESC, AGR.EFFECTIVE_DATE "
'   End If
'
'   If OrderType = 1 Then
'      OrderTypeStr = "ASC "
'   ElseIf OrderType = 2 Then
'      OrderTypeStr = "ASC "
'   Else
'      OrderTypeStr = "ASC "
'   End If
'
'   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
'   SQL2 = SelectStr2 & FromStr & WhereStr
'
'   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'End Function
'
'Public Function QueryData4(Rs As ADODB.Recordset, ItemCount As Long)
'Dim SQL1 As String
'Dim SQL2 As String
'Dim SelectStr1 As String
'Dim SelectStr2 As String
'Dim FromStr As String
'Dim WhereStr As String
'Dim OrderByStr As String
'Dim OrderTypeStr As String
'
'   SelectStr1 = "SELECT NULL ACCOUNT_ID, NULL SUBSCRIBER_ID, NULL SOC_ID, NULL EXPIRE_DATE, NULL EFFECTIVE_DATE, SF.* "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM SOC_FEATURE SF "
'   FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (SF.SOC_ID = SC.SOC_ID) "
'
'   WhereStr = "WHERE (SC.SOC_LEVEL = 'Y') AND ((SF.FEATURE_ID = " & FEATURE_ID & ") OR (SF.PART_ITEM_ID = " & PART_ITEM_ID & ")) "
'
'   OrderByStr = ""
'   OrderTypeStr = ""
'
'   If OrderBy = 1 Then
'      OrderByStr = ""
'   ElseIf OrderBy = 2 Then
'      OrderByStr = ""
'   Else
'      OrderByStr = ""
'   End If
'
'   If OrderType = 1 Then
'      OrderTypeStr = ""
'   ElseIf OrderType = 2 Then
'      OrderTypeStr = ""
'   Else
'      OrderTypeStr = ""
'   End If
'
'   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
'   SQL2 = SelectStr2 & FromStr & WhereStr
'
'   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'End Function

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
Dim SubLen As Long
 Dim NewStr As String
 
   If Ind = 1 Then
      SelectStr1 = "SELECT EWP.*, EWPI.*, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EX_WORKS_PRICE_ITEM EWPI "
      FromStr = FromStr & "LEFT OUTER JOIN EX_WORKS_PRICE EWP ON (EWPI.EX_WORKS_PRICE_ID = EWP.EX_WORKS_PRICE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EWPI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = ""
'   ElseIf Ind = 2 Then
'      SelectStr1 = "SELECT CM.CUSTOMER_ID, CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME, SC.SOC_CODE, PI.PART_NO, PI.PART_DESC, FT.FEATURE_CODE, FT.FEATURE_DESC, UR.PKG_RATE_AMOUNT, SC.SOC_LEVEL, SF.PART_ITEM_ID, SF.FEATURE_ID "
'      SelectStr2 = "SELECT COUNT(*)  "
'
'      FromStr = "FROM SOC_FEATURE SF "
'      FromStr = FromStr & "LEFT OUTER JOIN UC_RATE UR ON (UR.SOC_FEATURE_ID = SF.SOC_FEATURE_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (SC.SOC_ID = SF.SOC_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (FT.FEATURE_ID = SF.FEATURE_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = SF.PART_ITEM_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN AGREEMENT AGR ON (AGR.SOC_ID = SF.SOC_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN SUBSCRIBER SB ON (SB.SUBSCRIBER_ID = AGR.SUBSCRIBER_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (AC.ACCOUNT_ID = SB.ACCOUNT_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = AC.CUSTOMER_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
'
'      WhereStr = ""
   End If
   
   If EX_WORKS_PRICE_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWPI.EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EWPI.EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID & ") "
      End If
   End If
   
   If EX_WORKS_PRICE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWP.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      End If
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   If Len(ENABLE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AC.ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (AC.ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY  PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EX_WORKS_PRICE_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EX_WORKS_PRICE_ITEM_ID "
      End If
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EX_WORKS_PRICE_ITEM_ID = NVLI(Rs("EX_WORKS_PRICE_ITEM_ID"), -1)
      EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PACKAGE_RATE = NVLD(Rs("PACKAGE_RATE"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
      EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
   End If
End Function



Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

'   TempID1 = FEATURE_ID
'   If FEATURE_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SOC_FEATURE_SEQ", ID, glbErrorLog)
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("SOC_FEATURE_SEQ", ID, glbErrorLog)
      Else
         ID = EX_WORKS_PRICE_ITEM_ID
      End If
      
      SQL1 = "INSERT INTO EX_WORKS_PRICE_ITEM " & _
                     "( " & _
                     "   EX_WORKS_PRICE_ITEM_ID, " & _
                     "   EX_WORKS_PRICE_ID, " & _
                     "   RATE_TYPE, " & _
                     "   PART_ITEM_ID, " & _
                     "   PACKAGE_RATE, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "" & EX_WORKS_PRICE_ID & ", " & _
                        RATE_TYPE & ", " & _
                        TempID2 & ", " & _
                        PACKAGE_RATE & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               EX_WORKS_PRICE_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EX_WORKS_PRICE_ITEM SET " & _
                     "   EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ", " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   PART_ITEM_ID = " & TempID2 & ", " & _
                     "   PACKAGE_RATE = " & PACKAGE_RATE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

'      SQL1 = "DELETE FROM AC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM OC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM RC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM UC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM STPTIER_VOL WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
'      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EX_WORKS_PRICE_ITEM WHERE EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
 Set m_Conn = Nothing
End Sub



