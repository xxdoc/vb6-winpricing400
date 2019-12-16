VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSoc"
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

Public SOC_ID As Long
Public SOC_CODE As String
Public SOC_LEVEL As String
Public SOC_STATUS As Long
Public SOC_DESC As String
Public SERVICE_LEVEL As Long

Public FEATURE_CODE As String
Public SOC_LEVEL_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String
Public MODIFY_DATE As Date

Public SocFeatures As Collection
Public QuoataPlan As Collection

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

   SelectStr1 = "SELECT ADR.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SOC ADR "
   WhereStr = ""
   
   If SOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_ID = " & SOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SOC_ID = " & SOC_ID & ") "
      End If
   End If
      
   If Len(SOC_LEVEL) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_LEVEL = '" & SOC_LEVEL & "') "
      Else
         WhereStr = WhereStr & "AND (SOC_LEVEL = '" & SOC_LEVEL & "') "
      End If
   End If
      
   If SOC_STATUS > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_STATUS = " & SOC_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (SOC_STATUS = " & SOC_STATUS & ") "
      End If
   End If
   
   If Trim(SOC_CODE) <> "" Then
      If WildCard(SOC_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SOC_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SOC_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SOC_CODE = '" & ChangeQuote(Trim(SOC_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SOC_CODE = '" & ChangeQuote(Trim(SOC_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(SOC_DESC) <> "" Then
      If WildCard(SOC_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SOC_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SOC_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SOC_DESC = '" & ChangeQuote(Trim(SOC_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (SOC_DESC = '" & ChangeQuote(Trim(SOC_DESC)) & "')"
         End If
      End If
   End If
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SOC_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SOC_ID "
   Else
      OrderByStr = "ORDER BY SOC_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   SOC_ID = NVLI(Rs("SOC_ID"), -1)
   SOC_CODE = NVLS(Rs("SOC_CODE"), "")
   SOC_LEVEL = NVLS(Rs("SOC_LEVEL"), "")
   SOC_DESC = NVLS(Rs("SOC_DESC"), "")
   SOC_STATUS = NVLI(Rs("SOC_STATUS"), -1)
   MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
   
   KEY_ID = SOC_ID
   KEY_LOOKUP = SOC_CODE
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SOC_SEQ", ID, glbErrorLog)
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("SOC_SEQ", ID, glbErrorLog)
      Else
         ID = SOC_ID
      End If
      
      SQL1 = "INSERT INTO SOC " & _
                     "( " & _
                     "   SOC_ID, " & _
                     "   SOC_CODE, " & _
                     "   SOC_LEVEL, " & _
                     "   SOC_DESC, " & _
                     "   SOC_STATUS, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(SOC_CODE)) & "'," & _
                         "'" & SOC_LEVEL & "'," & _
                         "'" & ChangeQuote(Trim(SOC_DESC)) & "'," & _
                         "" & SOC_STATUS & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SOC SET " & _
                     "   SOC_CODE = '" & ChangeQuote(Trim(SOC_CODE)) & "', " & _
                     "   SOC_LEVEL = '" & SOC_LEVEL & "', " & _
                     "   SOC_DESC = '" & ChangeQuote(Trim(SOC_DESC)) & "', " & _
                     "   SOC_STATUS = " & SOC_STATUS & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SOC_ID = " & SOC_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM RC_RATE WHERE SOC_FEATURE_ID IN (SELECT SOC_FEATURE_ID FROM SOC_FEATURE WHERE SOC_ID = " & SOC_ID & " ) "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM OC_RATE WHERE SOC_FEATURE_ID IN (SELECT SOC_FEATURE_ID FROM SOC_FEATURE WHERE SOC_ID = " & SOC_ID & " ) "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM AC_RATE WHERE SOC_FEATURE_ID IN (SELECT SOC_FEATURE_ID FROM SOC_FEATURE WHERE SOC_ID = " & SOC_ID & " ) "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM UC_RATE WHERE SOC_FEATURE_ID IN (SELECT SOC_FEATURE_ID FROM SOC_FEATURE WHERE SOC_ID = " & SOC_ID & " ) "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM STPTIER_VOL WHERE SOC_FEATURE_ID IN (SELECT SOC_FEATURE_ID FROM SOC_FEATURE WHERE SOC_ID = " & SOC_ID & " ) "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM SOC_FEATURE WHERE SOC_ID = " & SOC_ID
      m_Conn.Execute (SQL1)
            
      SQL1 = "DELETE FROM SOC WHERE SOC_ID = " & SOC_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set SocFeatures = Nothing
End Sub

