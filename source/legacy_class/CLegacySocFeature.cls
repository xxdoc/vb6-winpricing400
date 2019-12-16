VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLegacySocFeature"
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
Public RC_FLAG As String
Public UC_FLAG As String
Public OC_FLAG As String
Public AC_FLAG As String
Public UNIT As Long
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

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
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

   WhereStr = "WHERE ((AC.ACCOUNT_ID = " & ACCOUNT_ID & " AND AGR.SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") OR (AC.ACCOUNT_ID = " & ACCOUNT_ID & " AND AGR.SUBSCRIBER_ID <> " & SUBSCRIBER_ID & " AND SB.DUMMY_FLAG = 'Y')) AND (SF.FEATURE_ID = " & FEATURE_ID & ") AND (AGR.SOC_ID IS NOT NULL) "
   
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

   WhereStr = "WHERE (SC.SOC_LEVEL = 'Y') AND (SF.FEATURE_ID = " & FEATURE_ID & ") "
   
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
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT LT.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SOC_FEATURE LT "
   
   WhereStr = ""
   
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
   
   If FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEATURE_ID = " & FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FEATURE_ID = " & FEATURE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FEATURE_CODE "
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
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
   RC_FLAG = NVLS(Rs("RC_FLAG"), "")
   UC_FLAG = NVLS(Rs("UC_FLAG"), "")
   OC_FLAG = NVLS(Rs("OC_FLAG"), "")
   AC_FLAG = NVLS(Rs("AC_FLAG"), "")
   UNIT = NVLI(Rs("UNIT"), -1)
   RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
   MINIMUM_FLAG = NVLS(Rs("MINIMUM_FLAG"), "")
   MINIMUM_UNIT = NVLD(Rs("MINIMUM_UNIT"), -1)
   USE_START_FLAG = NVLS(Rs("USE_START_FLAG"), "")
   USE_END_FLAG = NVLS(Rs("USE_END_FLAG"), "")
   ROUNDING_FACTOR = NVLD(Rs("ROUNDING_FACTOR"), -1)
   
'   UNIT_TYPE = NVLI(Rs("UNTTYPE_ID"), -1)
'   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
'   FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
'   FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
'   RTTYPE_NAME = NVLS(Rs("RTTYPE_NAME"), "")
'   FEATURE_LEVEL = NVLI(Rs("FEATURE_LVEL"), -1)
'   SOC_CODE = NVLS(Rs("SOC_CODE"), "")
'   SOC_DESC = NVLS(Rs("SOC_DESC"), "")
'   SOC_SERVICE_LEVEL = NVLI(Rs("SERVICE_LEVEL"), -1)
'   SOC_SVCLVL_NAME = NVLS(Rs("SVCLVL_NAME"), "")
'   SOC_SVCLVL_ID = NVLI(Rs("SVCLVL_ID"), -1)
End Function

Public Function PopulateFromRS3(Rs As ADODB.Recordset)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
   RC_FLAG = NVLS(Rs("RC_FLAG"), "")
   UC_FLAG = NVLS(Rs("UC_FLAG"), "")
   OC_FLAG = NVLS(Rs("OC_FLAG"), "")
   AC_FLAG = NVLS(Rs("AC_FLAG"), "")
   UNIT = NVLI(Rs("UNIT"), -1)
   RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
   MINIMUM_FLAG = NVLS(Rs("MINIMUM_FLAG"), "")
   MINIMUM_UNIT = NVLD(Rs("MINIMUM_UNIT"), -1)
   USE_START_FLAG = NVLS(Rs("USE_START_FLAG"), "")
   USE_END_FLAG = NVLS(Rs("USE_END_FLAG"), "")
   ROUNDING_FACTOR = NVLD(Rs("ROUNDING_FACTOR"), -1)
   
   SUBSCRIBER_ID = NVLI(Rs("SUBSCRIBER_ID"), -1)
   ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SOC_FEATURE_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO SOC_FEATURE " & _
                     "( " & _
                     "   SOC_FEATURE_ID, " & _
                     "   SOC_ID, " & _
                     "   FEATURE_ID, " & _
                     "   RC_FLAG, " & _
                     "   UC_FLAG, " & _
                     "   OC_FLAG, " & _
                     "   AC_FLAG, " & _
                     "   UNIT, " & _
                     "   RATE_TYPE, " & _
                     "   MINIMUM_FLAG, " & _
                     "  MINIMUM_UNIT, " & _
                     "   USE_START_FLAG, " & _
                     "   USE_END_FLAG, " & _
                     "   ROUNDING_FACTOR, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "" & SOC_ID & ", " & _
                        "" & FEATURE_ID & ", " & _
                        "'" & RC_FLAG & "', " & _
                        "'" & UC_FLAG & "', " & _
                        "'" & OC_FLAG & "', " & _
                        "'" & AC_FLAG & "', " & _
                        "" & UNIT & ", " & _
                        "" & RATE_TYPE & ", " & _
                        "'" & MINIMUM_FLAG & "', " & _
                        "" & MINIMUM_UNIT & ", " & _
                        "'" & USE_START_FLAG & "', " & _
                        "'" & USE_END_FLAG & "', " & _
                        "" & ROUNDING_FACTOR & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               SOC_FEATURE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SOC_FEATURE SET " & _
                     "   SOC_ID = " & SOC_ID & ", " & _
                     "   FEATURE_ID = " & FEATURE_ID & ", " & _
                     "   RC_FLAG = '" & RC_FLAG & "', " & _
                     "   UC_FLAG = '" & UC_FLAG & "', " & _
                     "   OC_FLAG = '" & OC_FLAG & "', " & _
                     "   AC_FLAG = '" & AC_FLAG & "', " & _
                     "   UNIT = " & UNIT & ", " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   MINIMUM_FLAG = '" & MINIMUM_FLAG & "', " & _
                     "   MINIMUM_UNIT = " & MINIMUM_UNIT & ", " & _
                     "   USE_START_FLAG = '" & USE_START_FLAG & "', " & _
                     "   USE_END_FLAG = '" & USE_END_FLAG & "', " & _
                     "   ROUNDING_FACTOR = " & ROUNDING_FACTOR & ", " & _
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


