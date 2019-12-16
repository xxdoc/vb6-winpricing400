VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpHealty"
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

Public EMP_HEALTY_ID As Long
Public EMP_ID As Long
Public HOSPITAL_NAME As String
Public HEALT_DESC As String
Public FROM_DATE As Date
Public TO_DATE As Date

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

   SelectStr1 = "SELECT EE.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMP_HEALTY EE "

   WhereStr = ""
   
   If EMP_HEALTY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_HEALTY_ID = " & EMP_HEALTY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_HEALTY_ID = " & EMP_HEALTY_ID & ") "
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
      OrderByStr = "ORDER BY HOSPITAL_NAME "
   Else
      OrderByStr = "ORDER BY EMP_HEALTY_ID  "
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
   EMP_HEALTY_ID = NVLI(Rs("EMP_HEALTY_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   HOSPITAL_NAME = NVLS(Rs("HOSPITAL_NAME"), "")
   HEALT_DESC = NVLS(Rs("HEALT_DESC"), "")
   TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
   FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMP_HEALTY_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMP_HEALTY " & _
                     "( " & _
                     "   EMP_HEALTY_ID, " & _
                     "   EMP_ID, " & _
                     "   HOSPITAL_NAME, " & _
                     "   HEALT_DESC, " & _
                     "   TO_DATE, " & _
                     "   FROM_DATE, " & _
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
                         "'" & HOSPITAL_NAME & "'," & _
                         "'" & HEALT_DESC & "'," & _
                         "'" & DateToStringInt(Trim(TO_DATE)) & "'," & _
                         "'" & DateToStringInt(Trim(FROM_DATE)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMP_HEALTY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMP_HEALTY SET " & _
                     "   EMP_ID = " & EMP_ID & ", " & _
                     "   HOSPITAL_NAME = '" & HOSPITAL_NAME & "', " & _
                     "   HEALT_DESC = '" & HEALT_DESC & "', " & _
                     "    TO_DATE = '" & DateToStringInt(TO_DATE) & "', " & _
                     "    FROM_DATE = '" & DateToStringInt(FROM_DATE) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_HEALTY_ID = " & EMP_HEALTY_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EMP_HEALTY WHERE EMP_HEALTY_ID = " & EMP_HEALTY_ID
      m_Conn.Execute (SQL1)
End Function

