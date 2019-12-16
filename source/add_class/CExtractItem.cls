VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExtractItem"
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

Public EXTRACT_ITEM_ID As Long
Public QUANTITY_EXTRACT_ID As Long
Public PART_ITEM_ID As Long
Public TOTAL_AMT As Double
 Public STD_AMOUNT As Double
 
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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT CM.*, PI.PART_NO, PI.PART_DESC  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXTRACT_ITEM CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If EXTRACT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXTRACT_ITEM_ID = " & EXTRACT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXTRACT_ITEM_ID = " & EXTRACT_ITEM_ID & ") "
      End If
   End If
      
   If QUANTITY_EXTRACT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID & ") "
      End If
   End If

   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
            
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY EXTRACT_ITEM_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If

   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EXTRACT_ITEM_ID = NVLI(Rs("EXTRACT_ITEM_ID"), 0)
      QUANTITY_EXTRACT_ID = NVLI(Rs("QUANTITY_EXTRACT_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TOTAL_AMT = NVLD(Rs("TOTAL_AMT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
      
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
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String
Dim TempID8 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXTRACT_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EXTRACT_ITEM " & _
                     "( " & _
                     "   EXTRACT_ITEM_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   TOTAL_AMT, " & _
                     "   QUANTITY_EXTRACT_ID, " & _
                     "   STD_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        PART_ITEM_ID & "," & _
                        TOTAL_AMT & "," & _
                        QUANTITY_EXTRACT_ID & "," & _
                        STD_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EXTRACT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXTRACT_ITEM SET " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    TOTAL_AMT = " & TOTAL_AMT & ", " & _
                     "    QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID & ", " & _
                     "    STD_AMOUNT = " & STD_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXTRACT_ITEM_ID = " & EXTRACT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM EXTRACT_ITEM WHERE EXTRACT_ITEM_ID = " & EXTRACT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

