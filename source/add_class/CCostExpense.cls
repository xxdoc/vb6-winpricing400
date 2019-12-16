VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostExpense"
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

Public COST_EXPENSE_ID As Long
Public COST_PRODUCTION_ID As Long
Public EXPENSE_TYPE As Long
Public PACKAGE_TYPE As PARCEL_TYPE
Public RATIO_TYPE As RATIO_TYPE
Public EXPENSE_AMOUNT As Double

Public PARAMETER_PROCESS_NAME  As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public id As String

Public CostRaws As Collection

'Public CstDescs As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CostRaws = New Collection
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

   SelectStr1 = "SELECT INP.*, PI.PARAMETER_PROCESS_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM COST_EXPENSE INP "
   FromStr = FromStr & "LEFT OUTER JOIN PARAMETER_PROCESS PI ON (PI.PARAMETER_PROCESS_ID= INP.EXPENSE_TYPE ) "
   
   If COST_EXPENSE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COST_EXPENSE_ID = " & COST_EXPENSE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COST_EXPENSE_ID = " & COST_EXPENSE_ID & ") "
      End If
   End If
   
   If COST_PRODUCTION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY COST_EXPENSE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY COST_EXPENSE_ID "
   Else
      OrderByStr = "ORDER BY COST_EXPENSE_ID "
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
   COST_EXPENSE_ID = NVLI(Rs("COST_EXPENSE_ID"), -1)
   COST_PRODUCTION_ID = NVLI(Rs("COST_PRODUCTION_ID"), -1)
   EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
   RATIO_TYPE = NVLI(Rs("RATIO_TYPE"), -1)
   PACKAGE_TYPE = NVLI(Rs("PACKAGE_TYPE"), -1)
   EXPENSE_AMOUNT = NVLD(Rs("EXPENSE_AMOUNT"), 0)
   
   PARAMETER_PROCESS_NAME = NVLS(Rs("PARAMETER_PROCESS_NAME"), "")
  End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
   
   TempID1 = EXPENSE_TYPE
   If EXPENSE_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("COST_EXPENSE_SEQ", id, glbErrorLog)

      SQL1 = "INSERT INTO COST_EXPENSE " & _
                     "( " & _
                     "   COST_EXPENSE_ID, " & _
                     "   COST_PRODUCTION_ID, " & _
                     "   EXPENSE_TYPE, " & _
                     "   EXPENSE_AMOUNT, " & _
                     "   PACKAGE_TYPE, " & _
                     "   RATIO_TYPE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                             "'" & COST_PRODUCTION_ID & "'," & _
                             TempID1 & "," & _
                        EXPENSE_AMOUNT & "," & _
                        PACKAGE_TYPE & "," & _
                        RATIO_TYPE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         COST_EXPENSE_ID = id
   Else 'Show edit
                        SQL1 = "UPDATE COST_EXPENSE SET " & _
                        "  EXPENSE_TYPE = " & TempID1 & ", " & _
                     "   EXPENSE_AMOUNT= " & EXPENSE_AMOUNT & ", " & _
                     "   RATIO_TYPE= " & RATIO_TYPE & ", " & _
                     "   PACKAGE_TYPE= " & PACKAGE_TYPE & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE COST_EXPENSE_ID = " & COST_EXPENSE_ID
   End If
   
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM COST_RAW WHERE COST_EXPENSE_ID = " & COST_EXPENSE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COST_EXPENSE WHERE COST_EXPENSE_ID = " & COST_EXPENSE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CostRaws = Nothing
End Sub
