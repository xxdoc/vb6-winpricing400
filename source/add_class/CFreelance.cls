VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFreelance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public FREELANCE_ID As Long
Public FREELANCE_CODE As String
Public FREELANCE_NAME As String
Public FREELANCE_LASTNAME As String
Public FREELANCE_RESIGN_FLAG As String
Public EMP_ID As Long
Public EMP_CODE As String

Public OrderBy As Long
Public OrderType As Long
Public QueryFlag As Long
Public AddEditMode As SHOW_MODE_TYPE
Public Flag As String


Public KEY_ID As Long
Public KEY_LOOKUP As String

Public CREATE_NAME As String
Public MODIFY_NAME As String
Public CREATE_BY As Long
Public MODIFY_BY As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   QueryFlag = -1
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim Temp1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FREELANCE_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO FREELANCE " & _
                     "( " & _
                     "   FREELANCE_ID, " & _
                     "   FREELANCE_CODE, " & _
                     "   FREELANCE_NAME, " & _
                     "   FREELANCE_LASTNAME, " & _
                     "    FREELANCE_RESIGN_FLAG, " & _
                     "    EMP_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         id & "," & _
                        "'" & ChangeQuote(Trim(FREELANCE_CODE)) & "', " & _
                        "'" & ChangeQuote(Trim(FREELANCE_NAME)) & "', " & _
                        "'" & ChangeQuote(Trim(FREELANCE_LASTNAME)) & "', " & _
                        "'" & ChangeQuote(Trim(FREELANCE_RESIGN_FLAG)) & "', " & _
                        ChangeQuote(Trim(EMP_ID)) & ", " & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ") "
                  FREELANCE_ID = id
   Else 'Show edit
      SQL1 = "UPDATE FREELANCE SET " & _
                     "   FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "', " & _
                     "   FREELANCE_NAME = '" & ChangeQuote(Trim(FREELANCE_NAME)) & "', " & _
                     "   FREELANCE_LASTNAME = '" & ChangeQuote(Trim(FREELANCE_LASTNAME)) & "', " & _
                     "   FREELANCE_RESIGN_FLAG = '" & ChangeQuote(Trim(FREELANCE_RESIGN_FLAG)) & "', " & _
                     "   EMP_ID = " & EMP_ID & ", "
         SQL1 = SQL1 & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & "" & _
                    "WHERE FREELANCE_ID = " & FREELANCE_ID
   End If
 '  ''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
 m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM FREELANCE WHERE FREELANCE_ID = " & FREELANCE_ID
      m_Conn.Execute (SQL1)

End Function

Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long, Optional Ind As Long = 1)
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

 If Ind = 1 Then
   SelectStr1 = "SELECT FL.*,EP.EMP_CODE   "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FREELANCE FL "
   FromStr = FromStr & "LEFT JOIN EMPLOYEE EP ON (FL.EMP_ID=EP.EMP_ID)  "
   WhereStr = ""
ElseIf Ind = 2 Then
   SelectStr1 = "SELECT FL.*,EP.EMP_CODE   "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FREELANCE FL "
   FromStr = FromStr & "LEFT JOIN FREELANCE_ITEM FLI ON (FL.FREELANCE_ID=FLI.FREELANCE_ID)  "
   FromStr = FromStr & "LEFT JOIN CUSTOMER CM ON (FLI.CUSTOMER_ID=CM.CUSTOMER_ID)  "
   WhereStr = ""
End If
   
   If FREELANCE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FL.FREELANCE_ID = " & FREELANCE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FL.FREELANCE_ID = " & FREELANCE_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FL.EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FL.EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   If Trim(FREELANCE_CODE) <> "" Then
      If WildCard(FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(FREELANCE_NAME) <> "" Then
      If WildCard(FREELANCE_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FREELANCE_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FREELANCE_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FREELANCE_NAME = '" & ChangeQuote(Trim(FREELANCE_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (FREELANCE_NAME = '" & ChangeQuote(Trim(FREELANCE_NAME)) & "')"
         End If
      End If
   End If
   
   If Trim(FREELANCE_LASTNAME) <> "" Then
      If WildCard(FREELANCE_LASTNAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FREELANCE_LASTNAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FREELANCE_LASTNAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FREELANCE_LASTNAME = '" & ChangeQuote(Trim(FREELANCE_LASTNAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (FREELANCE_LASTNAME = '" & ChangeQuote(Trim(FREELANCE_LASTNAME)) & "')"
         End If
      End If
   End If
   
   If Len(FREELANCE_RESIGN_FLAG) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (FREELANCE_RESIGN_FLAG = '" & ChangeQuote(Trim(FREELANCE_RESIGN_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (FREELANCE_RESIGN_FLAG = '" & ChangeQuote(Trim(FREELANCE_RESIGN_FLAG)) & "')"
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FREELANCE_ID  "
   Else
      OrderByStr = "ORDER BY FREELANCE_CODE  "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
      FREELANCE_ID = NVLI(Rs("FREELANCE_ID"), -1)
      FREELANCE_CODE = NVLS(Rs("FREELANCE_CODE"), "")
      FREELANCE_NAME = NVLS(Rs("FREELANCE_NAME"), "")
      FREELANCE_LASTNAME = NVLS(Rs("FREELANCE_LASTNAME"), "")
      FREELANCE_RESIGN_FLAG = NVLS(Rs("FREELANCE_RESIGN_FLAG"), "")
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      
      KEY_ID = FREELANCE_ID
      KEY_LOOKUP = FREELANCE_CODE
      
      CREATE_BY = NVLI(Rs("CREATE_BY"), -1)
      MODIFY_BY = NVLI(Rs("MODIFY_BY"), -1)
   End If
End Function
