VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFeature"
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

Public FEATURE_ID As Long
Public FEATURE_CODE As String
Public FEATURE_TYPE As Long
Public FEATURE_STATUS As String
Public FEATURE_LEVEL As Long
Public FEATURE_DESC As String
Public FEATURE_UNIT As Long
Public SERVICE_FLAG As String

Public FEATURE_TYPE_NAME  As String
Public FEATURE_UNIT_NAME As String
Public SOC_ID As Long
Public EXCEPTION_FLAG As String
Public FEATURE_UNIT_TYPE As Long
Public PERIOD_TYPE As PERIOD_TYPE
Public LOGISTIC_FLAG As String
Public BILL_DIRECT_FLAG As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public SocFeatures As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set SocFeatures = New Collection
End Sub

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

   SelectStr1 = "SELECT ADR.*, UN.UNIT_NAME FEATURE_UNIT_NAME, UN.PERIOD_TYPE, FT.FEATURE_TYPE_NAME, FT.LOGISTIC_FLAG "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FEATURE ADR "
   FromStr = FromStr & "LEFT OUTER JOIN FEATURE_TYPE FT ON (ADR.FEATURE_TYPE = FT.FEATURE_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.FEATURE_UNIT = UN.UNIT_ID) "
   
   WhereStr = ""
   
   If FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.FEATURE_ID = " & FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.FEATURE_ID = " & FEATURE_ID & ") "
      End If
   End If
      
   If SOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((SELECT COUNT(*) FROM SOC_FEATURE SF WHERE (SF.FEATURE_ID = ADR.FEATURE_ID) AND (SF.SOC_ID = " & SOC_ID & ")) <= 0) "
      Else
         WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM SOC_FEATURE SF WHERE (SF.FEATURE_ID = ADR.FEATURE_ID) AND (SF.SOC_ID = " & SOC_ID & ")) <= 0) "
      End If
   End If
   
   If FEATURE_LEVEL > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEATURE_LVEL = " & FEATURE_LEVEL & ") "
      Else
         WhereStr = WhereStr & "AND (FEATURE_LVEL = " & FEATURE_LEVEL & ") "
      End If
   End If
   
    If Len(FEATURE_STATUS) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEATURE_STATUS ='" & FEATURE_STATUS & "') "
      Else
         WhereStr = WhereStr & "AND (FEATURE_STATUS = '" & FEATURE_STATUS & "') "
      End If
   End If
   
   If Len(BILL_DIRECT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILL_DIRECT_FLAG ='" & BILL_DIRECT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (BILL_DIRECT_FLAG = '" & BILL_DIRECT_FLAG & "') "
      End If
   End If
   
   'BILL_DIRECT_FLAG
      
   If FEATURE_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEATURE_TYPE = " & FEATURE_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (FEATURE_TYPE = " & FEATURE_TYPE & ") "
      End If
   End If
   
   If Trim(FEATURE_CODE) <> "" Then
      If WildCard(FEATURE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FEATURE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FEATURE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FEATURE_CODE = '" & ChangeQuote(Trim(FEATURE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FEATURE_CODE = '" & ChangeQuote(Trim(FEATURE_CODE)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
'   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FEATURE_CODE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY FEATURE_CODE "
      Else
         OrderByStr = "ORDER BY FEATURE_CODE "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
'   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
   FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
   FEATURE_STATUS = NVLS(Rs("FEATURE_STATUS"), "")
   FEATURE_LEVEL = NVLI(Rs("FEATURE_LVEL"), -1)
   FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
   FEATURE_TYPE = NVLI(Rs("FEATURE_TYPE"), -1)
   FEATURE_UNIT = NVLI(Rs("FEATURE_UNIT"), -1)
   SERVICE_FLAG = NVLS(Rs("SERVICE_FLAG"), "")
   
   KEY_ID = FEATURE_ID
   KEY_LOOKUP = FEATURE_CODE
   
   FEATURE_TYPE_NAME = NVLS(Rs("FEATURE_TYPE_NAME"), "")
   FEATURE_UNIT_NAME = NVLS(Rs("FEATURE_UNIT_NAME"), "")
   PERIOD_TYPE = NVLI(Rs("PERIOD_TYPE"), -1)
   LOGISTIC_FLAG = NVLS(Rs("LOGISTIC_FLAG"), "")
   BILL_DIRECT_FLAG = NVLS(Rs("BILL_DIRECT_FLAG"), "")
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

'   TempID1 = PART_ITEM_ID
'   If PART_ITEM_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("FEATURE_SEQ", ID, glbErrorLog)
      Else
         ID = FEATURE_ID
      End If

      SQL1 = "INSERT INTO FEATURE " & _
                     "( " & _
                     "   FEATURE_ID, " & _
                     "   FEATURE_CODE, " & _
                     "   FEATURE_STATUS, " & _
                     "   FEATURE_LVEL, " & _
                     "   FEATURE_DESC, " & _
                     "   FEATURE_TYPE, " & _
                     "   FEATURE_UNIT, " & _
                     "   SERVICE_FLAG, " & _
                     "   BILL_DIRECT_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(FEATURE_CODE)) & "'," & _
                         "'" & FEATURE_STATUS & "'," & _
                         " " & FEATURE_LEVEL & "," & _
                         "'" & ChangeQuote(Trim(FEATURE_DESC)) & "'," & _
                         "" & FEATURE_TYPE & "," & _
                         "" & FEATURE_UNIT & "," & _
                         "'" & ChangeQuote(Trim(SERVICE_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(BILL_DIRECT_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         FEATURE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE FEATURE SET " & _
                     "   FEATURE_CODE = '" & ChangeQuote(Trim(FEATURE_CODE)) & "', " & _
                     "   FEATURE_STATUS = '" & FEATURE_STATUS & "', " & _
                     "   FEATURE_LVEL = '" & FEATURE_LEVEL & "', " & _
                     "   FEATURE_DESC = '" & ChangeQuote(Trim(FEATURE_DESC)) & "', " & _
                     "   FEATURE_TYPE = " & FEATURE_TYPE & ", " & _
                     "   FEATURE_UNIT = " & FEATURE_UNIT & ", " & _
                     "   SERVICE_FLAG = '" & ChangeQuote(Trim(SERVICE_FLAG)) & "', " & _
                     "   BILL_DIRECT_FLAG = '" & ChangeQuote(Trim(BILL_DIRECT_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FEATURE_ID = " & FEATURE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM FEATURE WHERE FEATURE_ID = " & FEATURE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set SocFeatures = Nothing
End Sub


