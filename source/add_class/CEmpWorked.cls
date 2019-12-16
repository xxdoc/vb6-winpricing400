VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpWorked"
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

Public EMP_WORKED_ID As Long
Public EMP_ID As Long
Public WORK_PLACE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public RESIGN_REASON As Long
Public EMP_POSITION As String
Public RESIGN_REASON_NAME As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT EW.* ,RE.RSGRESON_NAME RESIGN_REASON_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMP_WORKED EW "
FromStr = FromStr & "LEFT OUTER JOIN RESIGN_REASON RE ON (RE.RSGRESON_ID = EW.RESIGN_REASON) "
   WhereStr = ""
   
   If EMP_WORKED_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_WORKED_ID = " & EMP_WORKED_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_WORKED_ID = " & EMP_WORKED_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FROM_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LONG_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY WORK_PLACE "
   Else
      OrderByStr = "ORDER BY EMP_WORKED_ID  "
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   EMP_WORKED_ID = NVLI(Rs("EMP_WORKED_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   WORK_PLACE = NVLS(Rs("WORK_PLACE"), "")
   TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
   FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
   RESIGN_REASON = NVLI(Rs("RESIGN_REASON"), -1)
   EMP_POSITION = NVLS(Rs("EMP_POSITION"), "")
   RESIGN_REASON_NAME = NVLS(Rs("RESIGN_REASON_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMP_WORKED_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMP_WORKED " & _
                     "( " & _
                     "   EMP_WORKED_ID, " & _
                     "   EMP_ID, " & _
                     "   WORK_PLACE, " & _
                     "   TO_DATE, " & _
                     "   FROM_DATE, " & _
                     "   RESIGN_REASON, " & _
                     "   EMP_POSITION, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & EMP_ID & "," & _
                         "'" & WORK_PLACE & "'," & _
                         "'" & DateToStringInt(Trim(TO_DATE)) & "'," & _
                         "'" & DateToStringInt(Trim(FROM_DATE)) & "'," & _
                         "" & RESIGN_REASON & "," & _
                         "'" & EMP_POSITION & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMP_WORKED_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMP_WORKED SET " & _
                     "   EMP_ID = " & EMP_ID & ", " & _
                     "   WORK_PLACE = '" & WORK_PLACE & "', " & _
                     "    TO_DATE = '" & DateToStringInt(TO_DATE) & "', " & _
                     "    FROM_DATE = '" & DateToStringInt(FROM_DATE) & "', " & _
                     "   RESIGN_REASON = " & RESIGN_REASON & ", " & _
                     "   EMP_POSITION = '" & EMP_POSITION & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_WORKED_ID = " & EMP_WORKED_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM EMP_WORKED WHERE EMP_WORKED_ID = " & EMP_WORKED_ID
      m_Conn.Execute (SQL1)
End Function

