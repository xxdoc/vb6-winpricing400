VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobParameter"
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

Public JOB_PARAMETER_ID As Long
Public JOB_ID As Long
Public PARAMETER_PROCESS_ID As Long
Public JOB_PARAMETER_DESC As String
Public PARAM_AMOUNT As Double

Public PARAMETER_PROCESS_NAME  As String
Public PART_ITEM_ID As Long
Public FROM_DATE As Date
Public TO_DATE As Date
Public PROCESS_ID As Long

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As String

'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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
Dim GroupByStr As String
Dim TempDate As String

   If Ind = 1 Then
      SelectStr1 = "SELECT INP.*,PI.PARAMETER_PROCESS_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM JOB_PARAMETER INP "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETER_PROCESS PI ON (PI.PARAMETER_PROCESS_ID = INP.PARAMETER_PROCESS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT JB.PART_ITEM_ID PART_ITEM_ID, INP.PARAMETER_PROCESS_ID, SUM(INP.PARAM_AMOUNT) PARAM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM JOB_PARAMETER INP "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (INP.JOB_ID = JB.JOB_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = JB.PART_ITEM_ID) "

      WhereStr = ""
   End If
   
   If JOB_PARAMETER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_PARAMETER_ID = " & JOB_PARAMETER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_PARAMETER_ID = " & JOB_PARAMETER_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (jb.PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (jb.PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
      
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JB.JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JB.JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JB.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JB.JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY JB.PART_ITEM_ID, INP.PARAMETER_PROCESS_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY JOB_PARAMETER_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY JOB_PARAMETER_ID "
      Else
         OrderByStr = "ORDER BY JOB_PARAMETER_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      JOB_PARAMETER_ID = NVLI(Rs("JOB_PARAMETER_ID"), -1)
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      PARAMETER_PROCESS_ID = NVLI(Rs("PARAMETER_PROCESS_ID"), -1)
      JOB_PARAMETER_DESC = NVLS(Rs("JOB_PARAMETER_DESC"), "")
      PARAM_AMOUNT = NVLD(Rs("PARAM_AMOUNT"), 0)
      
      PARAMETER_PROCESS_NAME = NVLS(Rs("PARAMETER_PROCESS_NAME"), "")
      
       KEY_ID = PARAMETER_PROCESS_ID
      KEY_LOOKUP = PARAMETER_PROCESS_NAME
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PARAMETER_PROCESS_ID = NVLI(Rs("PARAMETER_PROCESS_ID"), -1)
      PARAM_AMOUNT = NVLD(Rs("PARAM_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = PARAMETER_PROCESS_ID
   If PARAMETER_PROCESS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_PARAMETER_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO JOB_PARAMETER " & _
                     "( " & _
                     "   JOB_PARAMETER_ID, " & _
                     "   JOB_ID, " & _
                     "   PARAMETER_PROCESS_ID, " & _
                     "   JOB_PARAMETER_DESC, " & _
                     "   PARAM_AMOUNT, " & _
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
                        "'" & ChangeQuote(Trim(JOB_PARAMETER_DESC)) & "'," & _
                        PARAM_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         JOB_PARAMETER_ID = id
   Else 'Show edit
                        SQL1 = "UPDATE JOB_PARAMETER SET " & _
                     "   PARAMETER_PROCESS_ID = " & TempID1 & ", " & _
                     "   JOB_PARAMETER_DESC= '" & ChangeQuote(Trim(JOB_PARAMETER_DESC)) & "', " & _
                     "   PARAM_AMOUNT = " & PARAM_AMOUNT & ", " & _
                     "   MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_PARAMETER_ID = " & JOB_PARAMETER_ID
   End If
   
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM JOB_PARAMETER WHERE JOB_PARAMETER_ID = " & JOB_PARAMETER_ID
      m_Conn.Execute (SQL1)
End Function


