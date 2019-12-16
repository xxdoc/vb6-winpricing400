VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostProduction"
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

Public COST_PRODUCTION_ID As Long
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public JOB_FROM_DATE As Date
Public JOB_TO_DATE As Date

Public FROM_DATE As Date
Public TO_DATE As Date
Public KEY_ID As Long
Public KEY_LOOKUP As String

Public CostItems As Collection
Public CostItemsEx As Collection
Public ExpenseItem As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CostItems = New Collection
   Set ExpenseItem = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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
   
   FromStr = "FROM COST_PRODUCTION UG "
   
   WhereStr = ""
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
      OrderByStr = "ORDER BY DOCUMENT_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY DOCUMENT_DATE "
   Else
      OrderByStr = "ORDER BY COST_PRODUCTION_ID "
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
   COST_PRODUCTION_ID = NVLI(Rs("COST_PRODUCTION_ID"), -1)
   DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   JOB_FROM_DATE = InternalDateToDate(NVLS(Rs("JOB_FROM_DATE"), ""))
   JOB_TO_DATE = InternalDateToDate(NVLS(Rs("JOB_TO_DATE"), ""))
   
   KEY_ID = COST_PRODUCTION_ID
   KEY_LOOKUP = DOCUMENT_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = COST_PRODUCTION_ID
   If COST_PRODUCTION_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("COST_PRODUCTION_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO COST_PRODUCTION " & _
                     "( " & _
                     "   COST_PRODUCTION_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_NO, " & _
                     "   JOB_FROM_DATE, " & _
                     "   JOB_TO_DATE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(JOB_FROM_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(JOB_TO_DATE))) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         COST_PRODUCTION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE COST_PRODUCTION SET " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   JOB_FROM_DATE = '" & ChangeQuote(Trim(DateToStringInt(JOB_FROM_DATE))) & "', " & _
                     "   JOB_TO_DATE = '" & ChangeQuote(Trim(DateToStringInt(JOB_TO_DATE))) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM COST_RAW WHERE COST_EXPENSE_ID IN (SELECT COST_EXPENSE_ID FROM COST_EXPENSE WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COST_EXPENSE WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CSTITEM_RAW WHERE COST_ITEM_ID IN (SELECT DISTINCT CI.COST_ITEM_ID  FROM COSTPRD_ITEM  CS LEFT OUTER JOIN COST_ITEM CI ON (CS.COSTPRD_ITEM_ID = CI.COSTPRD_ITEM_ID) WHERE CS.COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COST_ITEM WHERE COSTPRD_ITEM_ID IN (SELECT COSTPRD_ITEM_ID FROM COSTPRD_ITEM WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM COSTPRD_ITEM WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COST_PRODUCTION WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID
      m_Conn.Execute (SQL1)
End Function
Public Function DeleteDataCstItemRaw()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM CSTITEM_RAW WHERE COST_ITEM_ID IN (SELECT DISTINCT CI.COST_ITEM_ID  FROM COSTPRD_ITEM  CS LEFT OUTER JOIN COST_ITEM CI ON (CS.COSTPRD_ITEM_ID = CI.COSTPRD_ITEM_ID) WHERE CS.COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM COST_ITEM WHERE COSTPRD_ITEM_ID IN (SELECT COSTPRD_ITEM_ID FROM COSTPRD_ITEM WHERE COST_PRODUCTION_ID = " & COST_PRODUCTION_ID & ") "
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ExpenseItem = Nothing
End Sub
