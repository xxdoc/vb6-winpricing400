VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployeeProof"
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

Public EMPLOYEE_PROOF_ID As Long
Public EMP_ID As Long
Public DOCUMENT_ID As Long

Public EMP_CODE As String
Public EMP_NAME As String
Public EMP_LASTNAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DOCUMENT_TYPE As Long
Public POSITION_ID As Long

Public Doc As CDocument

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Doc = New CDocument
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
   
   FromStr = "FROM EMPLOYEE_PROOF UG "

   WhereStr = ""
   If EMPLOYEE_PROOF_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMPLOYEE_PROOF_ID = " & EMPLOYEE_PROOF_ID & ") "
      Else
         WhereStr = "AND (EMPLOYEE_PROOF_ID = " & EMPLOYEE_PROOF_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EMPLOYEE_PROOF_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EMPLOYEE_PROOF_ID "
      Else
         OrderByStr = "ORDER BY EMPLOYEE_PROOF_ID "
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
   EMPLOYEE_PROOF_ID = NVLI(Rs("EMPLOYEE_PROOF_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   DOCUMENT_ID = NVLI(Rs("DOCUMENT_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMPLOYEE_PROOF_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO EMPLOYEE_PROOF " & _
                     "( " & _
                     "   EMPLOYEE_PROOF_ID, " & _
                     "   EMP_ID, " & _
                     "   DOCUMENT_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         id & "," & _
                         "" & EMP_ID & "," & _
                         "" & DOCUMENT_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMPLOYEE_PROOF_ID = id
   Else 'Show edit
      SQL1 = "UPDATE EMPLOYEE_PROOF SET " & _
                     "   EMP_ID = " & EMP_ID & ", " & _
                     "   DOCUMENT_ID = " & DOCUMENT_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMPLOYEE_PROOF_ID = " & EMPLOYEE_PROOF_ID
   End If
   ''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EMPLOYEE_PROOF WHERE EMPLOYEE_PROOF_ID = " & EMPLOYEE_PROOF_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Doc = Nothing
End Sub



Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
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
Dim TempDate As String

   SelectStr1 = "SELECT EM.*, NM.LONG_NAME, NM.LAST_NAME, DM.*,DT.*, EP.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMPLOYEE_PROOF UG "
   FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT DM ON (UG.DOCUMENT_ID = DM.DOCUMENT_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_TYPE DT ON (DM.DOCUMENT_TYPE = DT.DOCTYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (UG.EMP_ID = EM.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EM.EMP_ID = EN.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMP_POSITION EP ON (EM.CURRENT_POSITION = EP.POSITION_ID) "
   
   WhereStr = ""
   If EMPLOYEE_PROOF_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMPLOYEE_PROOF_ID = " & EMPLOYEE_PROOF_ID & ") "
      Else
         WhereStr = "AND (EMPLOYEE_PROOF_ID = " & EMPLOYEE_PROOF_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   If POSITION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CURRENT_POSITION = " & POSITION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CURRENT_POSITION = " & POSITION_ID & ") "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Trim(EMP_CODE) <> "" Then
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(EMP_NAME) <> "" Then
      If WildCard(EMP_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.LONG_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.LONG_NAME = '" & ChangeQuote(Trim(EMP_NAME)) & "')"
         End If
      End If
   End If
   
   If Trim(EMP_LASTNAME) <> "" Then
      If WildCard(EMP_LASTNAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.LAST_NAME = '" & ChangeQuote(Trim(EMP_LASTNAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.LAST_NAME = '" & ChangeQuote(Trim(EMP_LASTNAME)) & "')"
         End If
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (EXPIRE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EXPIRE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (EXPIRE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EXPIRE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EMP_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY NM.LAST_NAME "
   ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY EXPIRE_DATE "
   Else
      OrderByStr = "ORDER BY EMPLOYEE_PROOF_ID "
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
