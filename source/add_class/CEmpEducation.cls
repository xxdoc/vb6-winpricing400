VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpEducation"
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

Public EMP_EDUCATION_ID As Long
Public EMP_ID As Long
Public INSTITUTE As String
Public SCORE As Double
Public MASTER As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public QUALIFICATION_ID As Long
Public QUALIFICATION_NAME As String

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

   SelectStr1 = "SELECT EE.*, QT.QUALIFICATION_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMP_EDUCATION EE "
FromStr = FromStr & "LEFT OUTER JOIN QUALIFICATION_TYPE QT ON (QT.QUALIFICATION_TYPE_ID = EE.QUALIFICATION_ID) "
   WhereStr = ""
   
   If EMP_EDUCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_EDUCATION_ID = " & EMP_EDUCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_EDUCATION_ID = " & EMP_EDUCATION_ID & ") "
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
      OrderByStr = "ORDER BY INSTITUTE "
   Else
      OrderByStr = "ORDER BY EMP_EDUCATION_ID  "
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
   EMP_EDUCATION_ID = NVLI(Rs("EMP_EDUCATION_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   INSTITUTE = NVLS(Rs("INSTITUTE"), "")
   SCORE = NVLD(Rs("SCORE"), 0)
   MASTER = NVLS(Rs("MASTER"), "")
   TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
   FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
   QUALIFICATION_ID = NVLI(Rs("QUALIFICATION_ID"), -1)
   QUALIFICATION_NAME = NVLS(Rs("QUALIFICATION_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMP_EDUCATION_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMP_EDUCATION " & _
                     "( " & _
                     "   EMP_EDUCATION_ID, " & _
                     "   EMP_ID, " & _
                     "   INSTITUTE, " & _
                     "   SCORE, " & _
                     "   MASTER, " & _
                     "   TO_DATE, " & _
                     "   FROM_DATE, " & _
                     "   QUALIFICATION_ID, " & _
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
                         "'" & INSTITUTE & "'," & _
                         "" & SCORE & "," & _
                         "'" & MASTER & "'," & _
                         "'" & DateToStringInt(Trim(TO_DATE)) & "'," & _
                         "'" & DateToStringInt(Trim(FROM_DATE)) & "'," & _
                         "" & QUALIFICATION_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMP_EDUCATION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMP_EDUCATION SET " & _
                     "   EMP_ID = " & EMP_ID & ", " & _
                     "   INSTITUTE = '" & INSTITUTE & "', " & _
                     "   SCORE = " & SCORE & ", " & _
                     "   MASTER = '" & MASTER & "', " & _
                     "    TO_DATE = '" & DateToStringInt(TO_DATE) & "', " & _
                     "    FROM_DATE = '" & DateToStringInt(FROM_DATE) & "', " & _
                     "   QUALIFICATION_ID = " & QUALIFICATION_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_EDUCATION_ID = " & EMP_EDUCATION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EMP_EDUCATION WHERE EMP_EDUCATION_ID = " & EMP_EDUCATION_ID
      m_Conn.Execute (SQL1)
End Function

