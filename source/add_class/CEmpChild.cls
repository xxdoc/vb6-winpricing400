VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpChild"
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

Public EMP_CHILD_ID As Long
Public EMP_ID As Long
Public GENDER As Long
Public NAME_ID As Long
Public BIRTH_DATE As Date
Public PHONE As String
Public Name As cName

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Name = New cName
   BIRTH_DATE = -1
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
   
   FromStr = "FROM EMP_CHILD EE "

   WhereStr = ""
   
   If EMP_CHILD_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_CHILD_ID = " & EMP_CHILD_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_CHILD_ID = " & EMP_CHILD_ID & ") "
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
      OrderByStr = "ORDER BY BIRTH_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LONG_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY GENDER "
   Else
      OrderByStr = "ORDER BY EMP_CHILD_ID  "
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
   EMP_CHILD_ID = NVLI(Rs("EMP_CHILD_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   GENDER = NVLI(Rs("GENDER"), -1)
   NAME_ID = NVLD(Rs("NAME_ID"), 0)
   BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
   PHONE = NVLS(Rs("PHONE"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMP_CHILD_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMP_CHILD " & _
                     "( " & _
                     "   EMP_CHILD_ID, " & _
                     "   EMP_ID, " & _
                     "   GENDER, " & _
                     "   NAME_ID, " & _
                     "   BIRTH_DATE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                       "    MODIFY_BY, " & _
                    "    PHONE " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & EMP_ID & "," & _
                         "'" & GENDER & "'," & _
                         "" & NAME_ID & "," & _
                         "'" & DateToStringInt(Trim(BIRTH_DATE)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                             glbUser.USER_ID & "," & _
                        "'" & ChangeQuote(Trim(PHONE)) & "'" & _
                     ") "
         EMP_CHILD_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMP_CHILD SET " & _
                     "   EMP_ID = " & EMP_ID & ", " & _
                     "   GENDER = '" & GENDER & "', " & _
                     "   NAME_ID = " & NAME_ID & ", " & _
                     "    BIRTH_DATE = '" & DateToStringInt(BIRTH_DATE) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                          "    MODIFY_BY = " & glbUser.USER_ID & ", " & _
                     "    PHONE = '" & ChangeQuote(Trim(PHONE)) & "' " & _
                     "WHERE EMP_CHILD_ID = " & EMP_CHILD_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EMP_CHILD WHERE EMP_CHILD_ID = " & EMP_CHILD_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Name = Nothing
End Sub
