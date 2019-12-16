VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLegacyEnterprise"
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

Public ENTERPRISE_ID As Long
Public BUSINESS_TYPE As Long
Public EMAIL As String
Public TAX_ID As String
Public SETUP_DATE As Date
Public POLICY As String
 Public ENTERPRISE_TYPE As Long
Public WEBSITE As String

Public ENTERPRISE_NAME As String

Public EnpAddresses As Collection
Public EnpNames As Collection
'Public EnpLogo As CEnterpriseLogo
Public EnpPersons As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   Set EnpAddresses = New Collection
   Set EnpNames = New Collection
'   Set EnpLogo = New CEnterpriseLogo
   Set EnpPersons = New Collection
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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT ETP.*, NM.LONG_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ENTERPRISE ETP "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (ETP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
   
   WhereStr = ""
   
   If ENTERPRISE_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ENTERPRISE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ENTERPRISE_ID "
   Else
      OrderByStr = "ORDER BY ENTERPRISE_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
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
   ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
   BUSINESS_TYPE = NVLI(Rs("BUSINESS_TYPE"), -1)
   EMAIL = NVLS(Rs("EMAIL"), "")
   TAX_ID = NVLS(Rs("TAX_ID"), "")
   SETUP_DATE = InternalDateToDate(NVLS(Rs("SETUP_DATE"), ""))
   POLICY = NVLS(Rs("POLICY"), "")
   ENTERPRISE_TYPE = NVLI(Rs("ENTERPRISE_TYPE"), -1)
   ENTERPRISE_NAME = NVLS(Rs("LONG_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("ENTERPRISE_SEQ", ID, glbErrorLog)
   
      SQL1 = "INSERT INTO ENTERPRISE " & _
                     "( " & _
                     "   ENTERPRISE_ID, " & _
                     "   BUSINESS_TYPE, " & _
                     "   EMAIL, " & _
                     "   WEBSITE, " & _
                     "   TAX_ID, " & _
                     "   SETUP_DATE, " & _
                     "   POLICY, " & _
                     "   ENTERPRISE_TYPE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & BUSINESS_TYPE & "," & _
                         "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                         "'" & ChangeQuote(Trim(WEBSITE)) & "'," & _
                         "'" & ChangeQuote(Trim(TAX_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(SETUP_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(POLICY)) & "'," & _
                         "" & ENTERPRISE_TYPE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ENTERPRISE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ENTERPRISE SET " & _
                     "   BUSINESS_TYPE = " & BUSINESS_TYPE & ", " & _
                     "   EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                     "   WEBSITE = '" & ChangeQuote(Trim(WEBSITE)) & "', " & _
                     "   TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   SETUP_DATE = '" & ChangeQuote(Trim(DateToStringInt(SETUP_DATE))) & "', " & _
                     "   POLICY = '" & ChangeQuote(Trim(POLICY)) & "', " & _
                     "   ENTERPRISE_TYPE = " & ENTERPRISE_TYPE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ENTERPRISE WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set EnpAddresses = Nothing
   Set EnpNames = Nothing
'   Set EnpLogo = Nothing
   Set EnpPersons = Nothing
End Sub


