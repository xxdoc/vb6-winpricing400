VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAgreement"
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

Public AGREEMENT_ID As Long
Public SOC_FEATURE_ID As Long
Public SUBSCRIBER_ID As Long
Public EXCLUDE_FLAG As String
Public EFFECTIVE_DATE As Date
Public EXPIRE_DATE As Date
Public ISSUE_DATE As Date
Public SOC_ID As Long

Public ACCOUNT_ID As Long
Public SOC_CODE As String
Public FEATURE_CODE As String
Public FEATURE_DESC As String
Public RC_RATE_AMOUNT As Double
Public ACCOUNT_NO As String
Public SUBSCRIBER_NO As String
Public BILL_DATE As Date
Public FEATURE_ID As Long
Public UNIT_ID As Long
Public USAGE_GROUP_ID As Long
Public OC_FLAG As String

Public SOC_LEVEL_NAME As String
Public SOC_DESC As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

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

   SelectStr1 = "SELECT AG.*, SC.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM AGREEMENT AG "
   FromStr = FromStr & "LEFT OUTER JOIN SOC SC ON (AG.SOC_ID = SC.SOC_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUBSCRIBER SB ON (AG.SUBSCRIBER_ID = SB.SUBSCRIBER_ID) "
   
   WhereStr = ""
   If AGREEMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AGREEMENT_ID = " & AGREEMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AGREEMENT_ID = " & AGREEMENT_ID & ") "
      End If
   End If
   
   If SOC_FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      End If
   End If
   
   If SUBSCRIBER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
      End If
   End If
   
   If ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SB.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SB.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY AGREEMENT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY AGREEMENT_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY AGREEMENT_ID "
      Else
         OrderByStr = "ORDER BY SVL.SVCLVL_NAME, SC.SOC_CODE, FT.FEATURE_CODE "
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

Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT AG.*, SF.*, SC.*, FT.*, RCR.RATE_AMOUNT RC_RATE_AMOUNT, AC.ACCOUNT_NO, SB.SUBSCRIBER_NO "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM AGREEMENT AG, SOC_FEATURE SF, SOC SC, FEATURE FT, SUBSCRIBER SB, RC_RATE RCR, ACCOUNT AC "

   WhereStr = "WHERE (AG.SOC_FEATURE_ID = SF.SOC_FEATURE_ID) "
   WhereStr = WhereStr & "AND (AG.SUBSCRIBER_ID = SB.SUBSCRIBER_ID) "
   WhereStr = WhereStr & "AND (AC.ACCOUNT_ID = SB.ACCOUNT_ID) "
   WhereStr = WhereStr & "AND (SF.SOC_FEATURE_ID = RCR.SOC_FEATURE_ID) "
   WhereStr = WhereStr & "AND (SF.SOC_ID = SC.SOC_ID) "
   WhereStr = WhereStr & "AND (SF.FEATURE_ID = FT.FEATURE_ID) "
   WhereStr = WhereStr & "AND (SF.RC_FLAG = 'Y') "

   If ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
   
   If BILL_DATE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EFFECTIVE_DATE <= '" & DateToStringInt(BILL_DATE) & "') AND (EXPIRE_DATE >= '" & DateToStringInt(BILL_DATE) & "') "
      Else
         WhereStr = WhereStr & "AND (EFFECTIVE_DATE <= '" & DateToStringInt(BILL_DATE) & "') AND (EXPIRE_DATE >= '" & DateToStringInt(BILL_DATE) & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY ACCOUNT_ID, SUBSCRIBER_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY ACCOUNT_ID "
      Else
         OrderByStr = "ORDER BY ACCOUNT_ID "
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


Public Function PopulateFromRS(Rs As ADODB.Recordset)
   AGREEMENT_ID = NVLI(Rs("AGREEMENT_ID"), -1)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   SUBSCRIBER_ID = NVLI(Rs("SUBSCRIBER_ID"), -1)
   EXCLUDE_FLAG = NVLS(Rs("EXCLUDE_FLAG"), "")
   EFFECTIVE_DATE = InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), ""))
   EXPIRE_DATE = InternalDateToDate(NVLS(Rs("EXPIRE_DATE"), ""))
   ISSUE_DATE = InternalDateToDate(NVLS(Rs("ISSUE_DATE"), ""))
'   OC_FLAG = NVLS(Rs("OC_FLAG"), "")
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   
   SOC_CODE = NVLS(Rs("SOC_CODE"), "")
'   FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
'   FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
'   SOC_LEVEL_NAME = NVLS(Rs("SVCLVL_NAME"), "")
'   SOC_DESC = NVLS(Rs("SOC_DESC"), "")
'   SOC_ID = NVLI(Rs("SOC_ID"), -1)
'   FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
'   UNIT_ID = NVLI(Rs("UNIT"), -1)
End Function

Public Function PopulateFromRS2(Rs As ADODB.Recordset)
   AGREEMENT_ID = NVLI(Rs("AGREEMENT_ID"), -1)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   SUBSCRIBER_ID = NVLI(Rs("SUBSCRIBER_ID"), -1)
   EXCLUDE_FLAG = NVLS(Rs("EXCLUDE_FLAG"), "")
   EFFECTIVE_DATE = InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), ""))
   EXPIRE_DATE = InternalDateToDate(NVLS(Rs("EXPIRE_DATE"), ""))
   ISSUE_DATE = InternalDateToDate(NVLS(Rs("ISSUE_DATE"), ""))
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   
   SOC_CODE = NVLS(Rs("SOC_CODE"), "")
   FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
   FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
   SOC_DESC = NVLS(Rs("SOC_DESC"), "")
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
   RC_RATE_AMOUNT = NVLD(Rs("RC_RATE_AMOUNT"), -1)
   ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
   SUBSCRIBER_NO = NVLS(Rs("SUBSCRIBER_NO"), "")
   UNIT_ID = NVLI(Rs("UNIT"), -1)
    FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
End Function

Public Function CopyObject(C As CAgreement)
   C.AGREEMENT_ID = AGREEMENT_ID
   C.SOC_FEATURE_ID = SOC_FEATURE_ID
   C.SUBSCRIBER_ID = SUBSCRIBER_ID
   C.EXCLUDE_FLAG = EXCLUDE_FLAG
   C.EFFECTIVE_DATE = EFFECTIVE_DATE
   C.EXPIRE_DATE = EXPIRE_DATE
   C.ISSUE_DATE = ISSUE_DATE
   
   C.SOC_CODE = SOC_CODE
   C.FEATURE_CODE = FEATURE_CODE
   C.FEATURE_DESC = FEATURE_DESC
   C.SOC_DESC = SOC_DESC
   C.SOC_LEVEL_NAME = SOC_LEVEL_NAME
   C.SOC_ID = SOC_ID
   C.FEATURE_ID = FEATURE_ID
   
   C.Flag = Flag
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   TempID1 = SOC_FEATURE_ID
   If SOC_FEATURE_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = SOC_ID
   If SOC_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("AGREEMENT_SEQ", ID, glbErrorLog)
      Else
         ID = AGREEMENT_ID
      End If
      
      SQL1 = "INSERT INTO AGREEMENT " & _
                     "( " & _
                     "   AGREEMENT_ID, " & _
                     "   SOC_FEATURE_ID, " & _
                     "   SUBSCRIBER_ID, " & _
                     "   EXCLUDE_FLAG, " & _
                     "   EFFECTIVE_DATE, " & _
                     "   EXPIRE_DATE, " & _
                     "   ISSUE_DATE, " & _
                     "   SOC_ID, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "" & TempID1 & ", " & _
                        "" & SUBSCRIBER_ID & ", " & _
                        "'" & EXCLUDE_FLAG & "', " & _
                        "'" & ChangeQuote(Trim(DateToStringInt(EFFECTIVE_DATE))) & "', " & _
                        "'" & ChangeQuote(Trim(DateToStringInt(EXPIRE_DATE))) & "', " & _
                        "'" & ChangeQuote(Trim(DateToStringInt(ISSUE_DATE))) & "', " & _
                        "" & TempID2 & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               AGREEMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE AGREEMENT SET " & _
                     "   SOC_FEATURE_ID = " & TempID1 & ", " & _
                     "   SUBSCRIBER_ID = " & SUBSCRIBER_ID & ", " & _
                     "   EXCLUDE_FLAG = '" & EXCLUDE_FLAG & "', " & _
                     "   EFFECTIVE_DATE = '" & ChangeQuote(Trim(DateToStringInt(EFFECTIVE_DATE))) & "', " & _
                     "   EXPIRE_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXPIRE_DATE))) & "', " & _
                     "   ISSUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(ISSUE_DATE))) & "', " & _
                     "   SOC_ID = " & TempID2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE AGREEMENT_ID = " & AGREEMENT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM AGREEMENT WHERE AGREEMENT_ID = " & AGREEMENT_ID
      m_Conn.Execute (SQL1)
End Function



