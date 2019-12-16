VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBulkHole"
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

Public BULK_HOLE_ID As Long
Public BILLING_DOC_ID As Double
Public PART_ITEM_ID As Long
Public HOLE_AMOUNT As Double
Public NOTE As String

Public PART_NO As String
Public PART_DESC As String

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
Dim TempDate As String
Dim FromDate1 As String
Dim ToDate1 As String
Dim FromDate2 As String
Dim ToDate2 As String

   If Ind = 1 Then
      SelectStr1 = "SELECT LT.*, PI.PART_NO, PI.PART_DESC PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BULK_HOLE LT "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LT.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If BULK_HOLE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BULK_HOLE_ID = " & BULK_HOLE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BULK_HOLE_ID = " & BULK_HOLE_ID & ") "
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BULK_HOLE_ID "
      Else
         OrderByStr = "ORDER BY BULK_HOLE_ID "
      End If
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
   If Ind = 1 Then
      BULK_HOLE_ID = NVLI(Rs("BULK_HOLE_ID"), -1)
      BILLING_DOC_ID = NVLD(Rs("BILLING_DOC_ID"), 0)
      NOTE = NVLS(Rs("NOTE"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      HOLE_AMOUNT = NVLD(Rs("HOLE_AMOUNT"), 0)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BULK_HOLE_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO BULK_HOLE " & _
                     "( " & _
                     "   BULK_HOLE_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   NOTE, " & _
                     "   HOLE_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                         HOLE_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               BULK_HOLE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BULK_HOLE SET " & _
                     "   BILLING_DOC_ID = " & BILLING_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & TempID2 & ", " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "   HOLE_AMOUNT = " & HOLE_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BULK_HOLE_ID = " & BULK_HOLE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM BULK_HOLE WHERE BULK_HOLE_ID = " & BULK_HOLE_ID
      m_Conn.Execute (SQL1)
End Function





