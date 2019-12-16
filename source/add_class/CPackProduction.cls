VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPackProduction"
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

Public PACK_PRODUCTION_ID As Long
Public PACK_PRODUCTION_NO As String
Public PACK_PRODUCTION_AREA As Long
Public PACK_PRODUCTION_DATE As Date
Public PACK_PRODUCTION_DESC As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public KEY_ID As Long
Public KEY_LOOKUP As String

Public PackItems As Collection
'Public CostItemsEx As Collection
'Public ExpenseItem As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PackItems = New Collection
'   Set ExpenseItem = New Collection
End Sub

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
Dim TempDate As String

   SelectStr1 = "SELECT PP.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PACK_PRODUCTION PP "
   
   WhereStr = ""
   
   If PACK_PRODUCTION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PP.PACK_PRODUCTION_ID  = " & PACK_PRODUCTION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PP.PACK_PRODUCTION_ID  = " & PACK_PRODUCTION_ID & ") "
      End If
   End If
   
   If PACK_PRODUCTION_DATE > -1 Then
   TempDate = DateToStringIntLow(Trim(PACK_PRODUCTION_DATE))
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PP.PACK_PRODUCTION_DATE  = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PP.PACK_PRODUCTION_DATE  = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PACK_PRODUCTION_NO  "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PACK_PRODUCTION_DATE  "
   Else
      OrderByStr = "ORDER BY PACK_PRODUCTION_ID  "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   PACK_PRODUCTION_ID = NVLI(Rs("PACK_PRODUCTION_ID"), -1)
   PACK_PRODUCTION_NO = NVLS(Rs("PACK_PRODUCTION_NO"), "")
   PACK_PRODUCTION_AREA = NVLI(Rs("PACK_PRODUCTION_AREA"), -1)
   PACK_PRODUCTION_DATE = InternalDateToDate(NVLS(Rs("PACK_PRODUCTION_DATE"), ""))
   PACK_PRODUCTION_DESC = NVLS(Rs("PACK_PRODUCTION_DESC"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PACK_PRODUCTION_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PACK_PRODUCTION " & _
                     "(" & _
                     "   PACK_PRODUCTION_ID, " & _
                     "   PACK_PRODUCTION_NO, " & _
                     "   PACK_PRODUCTION_AREA, " & _
                     "   PACK_PRODUCTION_DATE, " & _
                     "   PACK_PRODUCTION_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & ") " & _
                     "VALUES " & _
                     "( " & ID & "," & _
                         "'" & ChangeQuote(Trim(PACK_PRODUCTION_NO)) & "'," & _
                         "" & PACK_PRODUCTION_AREA & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PACK_PRODUCTION_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(PACK_PRODUCTION_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & ") "
         PACK_PRODUCTION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PACK_PRODUCTION SET " & _
                     "   PACK_PRODUCTION_NO = '" & ChangeQuote(Trim(PACK_PRODUCTION_NO)) & "', " & _
                     "   PACK_PRODUCTION_AREA = " & PACK_PRODUCTION_AREA & ", " & _
                     "   PACK_PRODUCTION_DATE = '" & ChangeQuote(Trim(DateToStringInt(PACK_PRODUCTION_DATE))) & "', " & _
                     "   PACK_PRODUCTION_DESC = '" & ChangeQuote(Trim(PACK_PRODUCTION_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PACK_PRODUCTION_ID = " & PACK_PRODUCTION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PACK_PRODUCTION_ITEM WHERE PACK_PRODUCTION_ID = " & PACK_PRODUCTION_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM PACK_PRODUCTION WHERE PACK_PRODUCTION_ID = " & PACK_PRODUCTION_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set PackItems = Nothing
End Sub
