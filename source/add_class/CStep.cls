VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStep"
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

Public STEP_ID As Long
Public SOC_FEATURE_ID As Long
Public FROM_QUANTITY As Double
Public TO_QUANTITY As Double

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any tranSOC_tion
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
   
   FromStr = "FROM STEP UG "

   WhereStr = ""
   If STEP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STEP_ID = " & STEP_ID & ") "
      Else
         WhereStr = "AND (STEP_ID = " & STEP_ID & ") "
      End If
   End If
   
   If SOC_FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      Else
         WhereStr = "AND (SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY STEP_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY STEP_ID "
      Else
         OrderByStr = "ORDER BY STEP_ID "
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
   STEP_ID = NVLI(Rs("STEP_ID"), -1)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   FROM_QUANTITY = NVLD(Rs("FROM_QUANTITY"), -1)
   TO_QUANTITY = NVLD(Rs("TO_QUANTITY"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("STEP_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO STEP " & _
                     "( " & _
                     "   STEP_ID, " & _
                     "   SOC_FEATURE_ID, " & _
                     "   FROM_QUANTITY, " & _
                     "   TO_QUANTITY, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & SOC_FEATURE_ID & "," & _
                         "" & FROM_QUANTITY & "," & _
                         "" & TO_QUANTITY & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         STEP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STEP SET " & _
                     "   SOC_FEATURE_ID = " & SOC_FEATURE_ID & ", " & _
                     "   FROM_QUANTITY = " & FROM_QUANTITY & ", " & _
                     "   TO_QUANTITY = " & TO_QUANTITY & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STEP_ID = " & STEP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM STEP WHERE STEP_ID = " & STEP_ID
      m_Conn.Execute (SQL1)
End Function




