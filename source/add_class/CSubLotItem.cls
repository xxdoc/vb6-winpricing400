VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubLotItem"
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
'
Public SUBLOT_ITEM_ID As Long
Public LOT_ITEM_ID As Long
Public FROM_LOT_ID As Long
Public SUBLOT_AMT As Double
Public PREVIOUS_AMT As Double
Public NEXT_AMT As Double

Public KEY_ID As Long
Public KEY_LOOKUP As String
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public IncludeFlag As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT ADR.*, IVD.DOCUMENT_NO, IVD.DOCUMENT_DATE "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SUBLOT_ITEM ADR  "
   FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI ON (ADR.FROM_LOT_ID = LI.LOT_ITEM_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   
   WhereStr = ""
   
   If SUBLOT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUBLOT_ITEM_ID = " & SUBLOT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUBLOT_ITEM_ID = " & SUBLOT_ITEM_ID & ") "
      End If
   End If
         
   If LOT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.LOT_ITEM_ID = " & LOT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.LOT_ITEM_ID = " & LOT_ITEM_ID & ") "
      End If
   End If
         
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SUBLOT_ITEM_ID "
   Else
      OrderByStr = "ORDER BY SUBLOT_ITEM_ID "
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
   SUBLOT_ITEM_ID = NVLI(Rs("SUBLOT_ITEM_ID"), -1)
   LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
   SUBLOT_AMT = NVLD(Rs("SUBLOT_AMT"), 0)
   FROM_LOT_ID = NVLI(Rs("FROM_LOT_ID"), -1)
   PREVIOUS_AMT = NVLD(Rs("PREVIOUS_AMT"), 0)
   NEXT_AMT = NVLD(Rs("NEXT_AMT"), 0)
   
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = FROM_LOT_ID
   If FROM_LOT_ID <= 0 Then
      TempID1 = "NULL "
   End If
      
   TempID2 = LOT_ITEM_ID
   If LOT_ITEM_ID <= 0 Then
      TempID2 = "NULL "
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SUBLOT_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO SUBLOT_ITEM " & _
                     "( " & _
                     "   SUBLOT_ITEM_ID, " & _
                     "   LOT_ITEM_ID, " & _
                     "   SUBLOT_AMT, " & _
                     "   FROM_LOT_ID, " & _
                     "   PREVIOUS_AMT, " & _
                     "   NEXT_AMT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & TempID2 & "," & _
                         "'" & ChangeQuote(Trim(SUBLOT_AMT)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PREVIOUS_AMT)) & "'," & _
                         "'" & ChangeQuote(Trim(NEXT_AMT)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SUBLOT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUBLOT_ITEM SET " & _
                     "   LOT_ITEM_ID = " & TempID2 & ", " & _
                     "   SUBLOT_AMT = '" & SUBLOT_AMT & "', " & _
                     "   FROM_LOT_ID = " & TempID1 & ", " & _
                     "   PREVIOUS_AMT = '" & ChangeQuote(Trim(PREVIOUS_AMT)) & "', " & _
                     "   NEXT_AMT = '" & ChangeQuote(Trim(NEXT_AMT)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUBLOT_ITEM_ID = " & SUBLOT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM SUBLOT_ITEM WHERE SUBLOT_ITEM_ID = " & SUBLOT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function DeleteAllData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM SUBLOT_ITEM WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

