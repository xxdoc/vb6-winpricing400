VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocument"
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

Public DOCUMENT_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public ISSUE_DATE As Date
Public EXPIRE_DATE As Date
Public ISSUE_PLACE_ID As Long
Public DOCUMENT_DESC As String

Public CARD_ID_FLAG As String
Public ISSUE_PLACE As String
Public DOCTYPE_NAME As String

Public Address As CAddress

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Address = New CAddress
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

   SelectStr1 = "SELECT ADR.*, DT.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DOCUMENT ADR, DOCUMENT_TYPE DT "
   WhereStr = "WHERE (ADR.DOCUMENT_TYPE = DT.DOCTYPE_ID) "
   
   If DOCUMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_ID = " & DOCUMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_ID = " & DOCUMENT_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY DOCUMENT_ID "
      Else
         OrderByStr = "ORDER BY DOCUMENT_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
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
   DOCUMENT_ID = NVLI(Rs("DOCUMENT_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   ISSUE_DATE = InternalDateToDate(NVLS(Rs("ISSUE_DATE"), ""))
   EXPIRE_DATE = InternalDateToDate(NVLS(Rs("EXPIRE_DATE"), ""))
   ISSUE_PLACE_ID = NVLI(Rs("ISSUE_PLACE_ID"), -1)
   DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
   DOCTYPE_NAME = NVLS(Rs("DOCTYPE_NAME"), "")
   CARD_ID_FLAG = NVLS(Rs("CARD_ID_FLAG"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DOCUMENT_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO DOCUMENT " & _
                     "( " & _
                     "   DOCUMENT_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   ISSUE_DATE, " & _
                     "   EXPIRE_DATE, " & _
                     "   ISSUE_PLACE_ID, " & _
                     "   DOCUMENT_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_TYPE)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ISSUE_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EXPIRE_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(ISSUE_PLACE_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         DOCUMENT_ID = id
   Else 'Show edit
      SQL1 = "UPDATE DOCUMENT SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   DOCUMENT_TYPE = '" & ChangeQuote(Trim(DOCUMENT_TYPE)) & "', " & _
                     "   ISSUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(ISSUE_DATE))) & "', " & _
                     "   EXPIRE_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXPIRE_DATE))) & "', " & _
                     "   ISSUE_PLACE_ID = '" & ChangeQuote(Trim(ISSUE_PLACE_ID)) & "', " & _
                     "   DOCUMENT_DESC = '" & ChangeQuote(Trim(DOCUMENT_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DOCUMENT_ID = " & DOCUMENT_ID
   End If
''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM DOCUMENT WHERE DOCUMENT_ID = " & DOCUMENT_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Address = Nothing
End Sub
