VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPurchaseExpense"
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

Public PUREXP_ID As Long
Public PUREXP_NO As String
Public PUREXP_NAME As String
Public EXPENSE_RATE As Double

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
   
   FromStr = "FROM PURCHASE_EXPENSE UG "
   
   WhereStr = ""
   If PUREXP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PUREXP_ID = " & PUREXP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PUREXP_ID = " & PUREXP_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PUREXP_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PUREXP_NAME "
   Else
      OrderByStr = "ORDER BY PUREXP_ID "
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
   PUREXP_ID = NVLI(Rs("PUREXP_ID"), -1)
   PUREXP_NAME = NVLS(Rs("PUREXP_NAME"), "")
   PUREXP_NO = NVLS(Rs("PUREXP_NO"), "")
   EXPENSE_RATE = NVLD(Rs("EXPENSE_RATE"), 0)
   
   KEY_ID = PUREXP_ID
   KEY_LOOKUP = PUREXP_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PURCHASE_EXPENSE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PURCHASE_EXPENSE " & _
                     "( " & _
                     "   PUREXP_ID, " & _
                     "   PUREXP_NAME, " & _
                     "   PUREXP_NO, " & _
                     "   EXPENSE_RATE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PUREXP_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(PUREXP_NO)) & "'," & _
                         EXPENSE_RATE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PUREXP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PURCHASE_EXPENSE SET " & _
                     "   PUREXP_NO = '" & ChangeQuote(Trim(PUREXP_NO)) & "', " & _
                     "   PUREXP_NAME = '" & ChangeQuote(Trim(PUREXP_NAME)) & "', " & _
                     "   EXPENSE_RATE = " & EXPENSE_RATE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PUREXP_ID = " & PUREXP_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PURCHASE_EXPENSE WHERE PUREXP_ID = " & PUREXP_ID
      m_Conn.Execute (SQL1)
End Function




