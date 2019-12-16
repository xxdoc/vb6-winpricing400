VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobVerify"
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

Public JOB_VERIFY_ID As Long
Public JOB_ID As Long
Public PART_ITEM_ID As Long
Public NOTE As String
Public VERIFY_FLAG As String

Public PART_NO  As String
Public PART_DESC  As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As String

'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT INP.*, PI.PART_NO, PI.PART_DESC "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM JOB_VERIFY INP "
   FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = INP.PART_ITEM_ID) "
   
   If JOB_VERIFY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INP.JOB_VERIFY_ID = " & JOB_VERIFY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INP.JOB_VERIFY_ID = " & JOB_VERIFY_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JB.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JB.JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY INP.JOB_VERIFY_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY INP.JOB_VERIFY_ID "
   Else
      OrderByStr = "ORDER BY INP.JOB_VERIFY_ID "
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
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   JOB_VERIFY_ID = NVLI(Rs("JOB_VERIFY_ID"), -1)
   JOB_ID = NVLI(Rs("JOB_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   NOTE = NVLS(Rs("NOTE"), "")
   VERIFY_FLAG = NVLS(Rs("VERIFY_FLAG"), "")
   
   PART_NO = NVLS(Rs("PART_NO"), "")
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   
    KEY_ID = PART_ITEM_ID
   KEY_LOOKUP = PART_NO
  End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_VERIFY_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO JOB_VERIFY " & _
                     "( " & _
                     "   JOB_VERIFY_ID, " & _
                     "   JOB_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   NOTE, " & _
                     "   VERIFY_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                             "'" & JOB_ID & "'," & _
                             TempID1 & "," & _
                        "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                        "'" & ChangeQuote(Trim(VERIFY_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         JOB_VERIFY_ID = id
   Else 'Show edit
                        SQL1 = "UPDATE JOB_VERIFY SET " & _
                        "  PART_ITEM_ID = " & TempID1 & ", " & _
                     "   NOTE= '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "   VERIFY_FLAG= '" & ChangeQuote(Trim(VERIFY_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_VERIFY_ID = " & JOB_VERIFY_ID
   End If
   
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM JOB_VERIFY WHERE JOB_VERIFY_ID = " & JOB_VERIFY_ID
      m_Conn.Execute (SQL1)
End Function


