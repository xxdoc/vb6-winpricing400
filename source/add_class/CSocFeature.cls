VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocFeature"
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

Public SOC_FEATURE_ID As Long
Public SOC_ID As Long
Public FEATURE_ID As Long
Public PART_ITEM_ID As Long
Public RC_FLAG As String
Public UC_FLAG As String
Public OC_FLAG As String
Public AC_FLAG As String
Public RATE_TYPE As Long
Public MINIMUM_FLAG As String
Public MINIMUM_UNIT As Double
Public USE_START_FLAG As String
Public USE_END_FLAG As String
Public ROUNDING_FACTOR As Double

Public UNIT_TYPE As Long
Public UNIT_NAME As String
Public FEATURE_CODE As String
Public FEATURE_DESC As String
Public RTTYPE_NAME As String
Public FEATURE_LEVEL As Long
Public SOC_CODE As String
Public SOC_DESC As String
Public SOC_SERVICE_LEVEL As Long
Public SOC_SVCLVL_NAME As String
Public SOC_SVCLVL_ID As String
Public FEATURE_TYPE As Long
Public PART_TYPE As Long
Public PART_NO As String
Public PART_DESC As String
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public CUSTOMER_TYPE As Long
Public PKG_RATE_AMOUNT As Double
Public CUSTOMER_ID As Long
Public SOC_LEVEL As String

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Public ACRate As CAcRate
Public OCRate As COcRate
Public RCRate As CRcRate
Public UCRate As CUcRate

Public Steps As Collection
Public Tiers As Collection
Public VolumeSteps As Collection
Public VolumeTiers As Collection
Public StepTierVol As Collection

Public SUBSCRIBER_ID As Long
Public ACCOUNT_ID As Long
Public ENABLE_FLAG As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ACRate = New CAcRate
   Set OCRate = New COcRate
   Set RCRate = New CRcRate
   Set UCRate = New CUcRate
   
   Set StepTierVol = New Collection
    Set Steps = New Collection
    Set Tiers = New Collection
    Set VolumeSteps = New Collection
    Set VolumeTiers = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT LT.*, FT.FEATURE_CODE, FT.FEATURE_DESC, FT.FEATURE_LVEL, UTT.UNTTYPE_ID, UT.UNIT_NAME, RT.RTTYPE_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SOC_FEATURE LT, FEATURE FT, UNIT_TYPE UTT, UNIT UT, RATE_TYPE RT "
   
   WhereStr = "WHERE (FT.FEATURE_ID = LT.FEATURE_ID) AND (UT.UNIT_ID = LT.UNIT) AND (UTT.UNTTYPE_ID = UT.UNIT_TYPE) AND (RT.RTTYPE_ID = LT.RATE_TYPE) "
   If SOC_FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      End If
   End If
   
   If SOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_ID = " & SOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SOC_ID = " & SOC_ID & ") "
      End If
   End If
   
      If Len(Trim(FROM_CUSTOMER_CODE)) > 0 Then
      If WildCard(FROM_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(TO_CUSTOMER_CODE)) > 0 Then
      If WildCard(TO_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      Else
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "DESC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData3(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT SB.ACCOUNT_ID, AGR.SUBSCRIBER_ID, AGR.SOC_ID, AGR.EXPIRE_DATE, AGR.EFFECTIVE_DATE, SF.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM AGREEMENT AGR "
   FromStr = FromStr & "LEFT OUTER JOIN SUBSCRIBER SB ON (AGR.SUBSCRIBER_ID = SB.SUBSCRIBER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (SB.ACCOUNT_ID = AC.ACCOUNT_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (AGR.SOC_ID = SC.SOC_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SOC_FEATURE SF ON (SC.SOC_ID = SF.SOC_ID) "

   WhereStr = "WHERE ((AC.ACCOUNT_ID = " & ACCOUNT_ID & " AND AGR.SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") OR (AC.ACCOUNT_ID = " & ACCOUNT_ID & " AND AGR.SUBSCRIBER_ID <> " & SUBSCRIBER_ID & " AND SB.DUMMY_FLAG = 'Y')) AND ((SF.FEATURE_ID = " & FEATURE_ID & ") OR SF.PART_ITEM_ID = " & PART_ITEM_ID & ") AND (AGR.SOC_ID IS NOT NULL) "
   
'   If FEATURE_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (FEATURE_ID = " & FEATURE_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (FEATURE_ID = " & FEATURE_ID & ") "
'      End If
'   End If
'
'   If PART_ITEM_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
'      End If
'   End If
'
'   If ACCOUNT_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
'      End If
'   End If
'
'   If SUBSCRIBER_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
'      End If
'   End If
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ACCOUNT_ID, SB.DUMMY_FLAG ASC, AGR.EXPIRE_DATE DESC, AGR.EFFECTIVE_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SB.DUMMY_FLAG ASC, AGR.EXPIRE_DATE DESC, AGR.EFFECTIVE_DATE "
   Else
      OrderByStr = "ORDER BY SB.DUMMY_FLAG ASC, AGR.EXPIRE_DATE DESC, AGR.EFFECTIVE_DATE "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "ASC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData4(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT NULL ACCOUNT_ID, NULL SUBSCRIBER_ID, NULL SOC_ID, NULL EXPIRE_DATE, NULL EFFECTIVE_DATE, SF.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SOC_FEATURE SF "
   FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (SF.SOC_ID = SC.SOC_ID) "

   WhereStr = "WHERE (SC.SOC_LEVEL = 'Y') AND ((SF.FEATURE_ID = " & FEATURE_ID & ") OR (SF.PART_ITEM_ID = " & PART_ITEM_ID & ")) "
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = ""
   ElseIf OrderBy = 2 Then
      OrderByStr = ""
   Else
      OrderByStr = ""
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = ""
   ElseIf OrderType = 2 Then
      OrderTypeStr = ""
   Else
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'In this class we will not handle any error or any transaction
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
 
   If Ind = 1 Then
      SelectStr1 = "SELECT SF.*, FT.FEATURE_CODE, FT.FEATURE_TYPE, FT.FEATURE_DESC, FT.FEATURE_LVEL, SC.*, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SOC_FEATURE SF "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (SF.FEATURE_ID = FT.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (SF.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (SF.SOC_ID = SC.SOC_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME, SC.SOC_CODE, PI.PART_NO, PI.PART_DESC, FT.FEATURE_CODE, FT.FEATURE_DESC, UR.PKG_RATE_AMOUNT, SC.SOC_LEVEL, SF.PART_ITEM_ID, SF.FEATURE_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SOC_FEATURE SF "
      FromStr = FromStr & "LEFT OUTER JOIN UC_RATE UR ON (UR.SOC_FEATURE_ID = SF.SOC_FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (SC.SOC_ID = SF.SOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (FT.FEATURE_ID = SF.FEATURE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = SF.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN AGREEMENT AGR ON (AGR.SOC_ID = SF.SOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUBSCRIBER SB ON (SB.SUBSCRIBER_ID = AGR.SUBSCRIBER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (AC.ACCOUNT_ID = SB.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = AC.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      
      WhereStr = ""
   End If
   
   If SOC_FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SF.SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SF.SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      End If
   End If
   
   If SOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SC.SOC_ID = " & SOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SC.SOC_ID = " & SOC_ID & ") "
      End If
   End If
   
   If FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SF.FEATURE_ID = " & FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SF.FEATURE_ID = " & FEATURE_ID & ") "
      End If
   End If
   
   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CM.CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CM.CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
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
         OrderByStr = "ORDER BY FEATURE_CODE,PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      Else
         OrderByStr = "ORDER BY SOC_FEATURE_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, NM.LONG_NAME ASC, CM.CUSTOMER_ID ASC, SC.SOC_CODE, PI.PART_NO, PI.PART_DESC, FT.FEATURE_CODE, FT.FEATURE_DESC "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "DESC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
      SOC_ID = NVLI(Rs("SOC_ID"), -1)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      RC_FLAG = NVLS(Rs("RC_FLAG"), "")
      UC_FLAG = NVLS(Rs("UC_FLAG"), "")
      OC_FLAG = NVLS(Rs("OC_FLAG"), "")
      AC_FLAG = NVLS(Rs("AC_FLAG"), "")
      RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
      MINIMUM_FLAG = NVLS(Rs("MINIMUM_FLAG"), "")
      MINIMUM_UNIT = NVLD(Rs("MINIMUM_UNIT"), -1)
      USE_START_FLAG = NVLS(Rs("USE_START_FLAG"), "")
      USE_END_FLAG = NVLS(Rs("USE_END_FLAG"), "")
      ROUNDING_FACTOR = NVLD(Rs("ROUNDING_FACTOR"), -1)
   
   '   UNIT_TYPE = NVLI(Rs("UNTTYPE_ID"), -1)
   '   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
      FEATURE_TYPE = NVLI(Rs("FEATURE_TYPE"), -1)
      RTTYPE_NAME = RateTypeToText(RATE_TYPE)
      FEATURE_LEVEL = NVLI(Rs("FEATURE_LVEL"), -1)
      SOC_CODE = NVLS(Rs("SOC_CODE"), "")
      SOC_DESC = NVLS(Rs("SOC_DESC"), "")
   '   SOC_SERVICE_LEVEL = NVLI(Rs("SERVICE_LEVEL"), -1)
   '   SOC_SVCLVL_NAME = NVLS(Rs("SVCLVL_NAME"), "")
   '   SOC_SVCLVL_ID = NVLI(Rs("SVCLVL_ID"), -1)
   ElseIf Ind = 2 Then
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
      SOC_CODE = NVLS(Rs("SOC_CODE"), "")
      SOC_LEVEL = NVLS(Rs("SOC_LEVEL"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      PKG_RATE_AMOUNT = NVLD(Rs("PKG_RATE_AMOUNT"), 0)
   End If
End Function

Public Function PopulateFromRS3(Rs As ADODB.Recordset)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   RC_FLAG = NVLS(Rs("RC_FLAG"), "")
   UC_FLAG = NVLS(Rs("UC_FLAG"), "")
   OC_FLAG = NVLS(Rs("OC_FLAG"), "")
   AC_FLAG = NVLS(Rs("AC_FLAG"), "")
   RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
   MINIMUM_FLAG = NVLS(Rs("MINIMUM_FLAG"), "")
   MINIMUM_UNIT = NVLD(Rs("MINIMUM_UNIT"), -1)
   USE_START_FLAG = NVLS(Rs("USE_START_FLAG"), "")
   USE_END_FLAG = NVLS(Rs("USE_END_FLAG"), "")
   ROUNDING_FACTOR = NVLD(Rs("ROUNDING_FACTOR"), -1)
   
   SUBSCRIBER_ID = NVLI(Rs("SUBSCRIBER_ID"), -1)
   ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = FEATURE_ID
   If FEATURE_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
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
         ID = SOC_FEATURE_ID
      End If
      
      SQL1 = "INSERT INTO SOC_FEATURE " & _
                     "( " & _
                     "   SOC_FEATURE_ID, " & _
                     "   SOC_ID, " & _
                     "   FEATURE_ID, " & _
                     "   RC_FLAG, " & _
                     "   UC_FLAG, " & _
                     "   OC_FLAG, " & _
                     "   AC_FLAG, " & _
                     "   RATE_TYPE, " & _
                     "   MINIMUM_FLAG, " & _
                     "  MINIMUM_UNIT, " & _
                     "   USE_START_FLAG, " & _
                     "   USE_END_FLAG, " & _
                     "   ROUNDING_FACTOR, " & _
                     "   PART_ITEM_ID, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "" & SOC_ID & ", " & _
                        TempID1 & ", " & _
                        "'" & RC_FLAG & "', " & _
                        "'" & UC_FLAG & "', " & _
                        "'" & OC_FLAG & "', " & _
                        "'" & AC_FLAG & "', " & _
                        "" & RATE_TYPE & ", " & _
                        "'" & MINIMUM_FLAG & "', " & _
                        "" & MINIMUM_UNIT & ", " & _
                        "'" & USE_START_FLAG & "', " & _
                        "'" & USE_END_FLAG & "', " & _
                        "" & ROUNDING_FACTOR & ", " & _
                        TempID2 & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               SOC_FEATURE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SOC_FEATURE SET " & _
                     "   SOC_ID = " & SOC_ID & ", " & _
                     "   FEATURE_ID = " & TempID1 & ", " & _
                     "   RC_FLAG = '" & RC_FLAG & "', " & _
                     "   UC_FLAG = '" & UC_FLAG & "', " & _
                     "   OC_FLAG = '" & OC_FLAG & "', " & _
                     "   AC_FLAG = '" & AC_FLAG & "', " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   MINIMUM_FLAG = '" & MINIMUM_FLAG & "', " & _
                     "   MINIMUM_UNIT = " & MINIMUM_UNIT & ", " & _
                     "   USE_START_FLAG = '" & USE_START_FLAG & "', " & _
                     "   USE_END_FLAG = '" & USE_END_FLAG & "', " & _
                     "   ROUNDING_FACTOR = " & ROUNDING_FACTOR & ", " & _
                     "   PART_ITEM_ID = " & TempID2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM AC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM OC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM RC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM UC_RATE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM STPTIER_VOL WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM SOC_FEATURE WHERE SOC_FEATURE_ID = " & SOC_FEATURE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ACRate = Nothing
   Set OCRate = Nothing
   Set RCRate = Nothing
   Set UCRate = Nothing
   
    Set Steps = Nothing
    Set Tiers = Nothing
    Set VolumeSteps = Nothing
    Set VolumeTiers = Nothing
    Set StepTierVol = Nothing
End Sub


