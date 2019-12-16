VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFreelanceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public FREELANCE_ITEM_ID As Long
Public FREELANCE_ID As Long
Public CUSTOMER_ID As Long
Public FREELANCE_CODE As String
Public FREELANCE_NAME As String
Public FREELANCE_LASTNAME As String
Public FREELANCE_RESIGN_FLAG As String

Public OrderBy As Long
Public OrderType As Long
Public QueryFlag As Long
Public AddEditMode As SHOW_MODE_TYPE
Public Flag As String
Public KEY_ID As Long
Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   QueryFlag = -1
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Temp1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FREELANCE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO FREELANCE_ITEM " & _
                     "( " & _
                     "   FREELANCE_ITEM_ID, " & _
                     "   FREELANCE_ID ," & _
                     "   CUSTOMER_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(FREELANCE_ID)) & "', " & _
                        "'" & ChangeQuote(Trim(CUSTOMER_ID)) & "', " & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ") "
                  FREELANCE_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE FREELANCE_ITEM SET " & _
                     "   FREELANCE_ID = '" & ChangeQuote(Trim(FREELANCE_ID)) & "', " & _
                     "   CUSTOMER_ID = '" & ChangeQuote(Trim(CUSTOMER_ID)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & "" & _
                    "WHERE FREELANCE_ITEM_ID = " & FREELANCE_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM FREELANCE_ITEM WHERE FREELANCE_ITEM_ID = " & FREELANCE_ITEM_ID
      m_Conn.Execute (SQL1)

End Function

Public Function QueryData(Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT FLI.*,FL.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FREELANCE_ITEM FLI "
   FromStr = FromStr & " LEFT OUTER JOIN FREELANCE FL ON (FLI.FREELANCE_ID = FL.FREELANCE_ID) "
   WhereStr = ""
   
   If FREELANCE_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FLI.FREELANCE_ITEM_ID = " & FREELANCE_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FLI.FREELANCE_ITEM_ID = " & FREELANCE_ITEM_ID & ") "
      End If
   End If
   
   If FREELANCE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FLI.FREELANCE_ID = " & FREELANCE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FLI.FREELANCE_ID = " & FREELANCE_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FLI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FLI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FREELANCE_ITEM_ID  "
   Else
      OrderByStr = "ORDER BY FREELANCE_ITEM_ID  "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
      FREELANCE_ITEM_ID = NVLI(Rs("FREELANCE_ITEM_ID"), -1)
      FREELANCE_ID = NVLI(Rs("FREELANCE_ID"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      FREELANCE_CODE = NVLS(Rs("FREELANCE_CODE"), "")
      FREELANCE_NAME = NVLS(Rs("FREELANCE_NAME"), "")
      FREELANCE_LASTNAME = NVLS(Rs("FREELANCE_LASTNAME"), "")
      FREELANCE_RESIGN_FLAG = NVLS(Rs("FREELANCE_RESIGN_FLAG"), "")
   End If
End Function
