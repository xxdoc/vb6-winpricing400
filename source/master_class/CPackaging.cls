VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPackaging"
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

Public PACKAGING_ID As Long
Public PACKAGING_NO As String
Public PACKAGING_NAME As String
Public WEIGHT_RATE As Double

Public KEY_ID As Long
Public KEY_LOOKUP As String

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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PACKAGING UG "
   
   WhereStr = ""
   If PACKAGING_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PACKAGING_ID = " & PACKAGING_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PACKAGING_ID = " & PACKAGING_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PACKAGING_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PACKAGING_NAME "
   Else
      OrderByStr = "ORDER BY PACKAGING_ID "
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
   PACKAGING_ID = NVLI(Rs("PACKAGING_ID"), -1)
   PACKAGING_NAME = NVLS(Rs("PACKAGING_NAME"), "")
   PACKAGING_NO = NVLS(Rs("PACKAGING_NO"), "")
   WEIGHT_RATE = NVLD(Rs("WEIGHT_RATE"), 0)
   
   KEY_ID = PACKAGING_ID
   KEY_LOOKUP = PACKAGING_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PACKAGING_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PACKAGING " & _
                     "( " & _
                     "   PACKAGING_ID, " & _
                     "   PACKAGING_NAME, " & _
                     "   PACKAGING_NO, " & _
                     "   WEIGHT_RATE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PACKAGING_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(PACKAGING_NO)) & "'," & _
                         WEIGHT_RATE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PACKAGING_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PACKAGING SET " & _
                     "   PACKAGING_NO = '" & ChangeQuote(Trim(PACKAGING_NO)) & "', " & _
                     "   PACKAGING_NAME = '" & ChangeQuote(Trim(PACKAGING_NAME)) & "', " & _
                     "   WEIGHT_RATE = " & WEIGHT_RATE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PACKAGING_ID = " & PACKAGING_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PACKAGING WHERE PACKAGING_ID = " & PACKAGING_ID
      m_Conn.Execute (SQL1)
End Function


