VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartLocation"
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

Public PART_LOCATION_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public AVG_PRICE As Double
Public LAST_PRICE As Double
Public MINIMUM_ALLOW As Double
Public CURRENT_AMOUNT As Double
Public TEMP_AVG_PRICE As Double
Public TEMP_LAST_PRICE As Double
Public TEMP_CURRENT_AMOUNT As Double

Public LOCATION_NAME As String
Public PART_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public COMMIT_FLAG As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
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
      SelectStr1 = "SELECT LT.*, LC.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_LOCATION LT "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LT.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
   ElseIf Ind = 3 Then
   ElseIf Ind = 4 Then
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT PL.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_LOCATION PL "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PL.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
   End If
   
   If PART_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_LOCATION_ID = " & PART_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_LOCATION_ID = " & PART_LOCATION_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LT.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LT.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_LOCATION_ID "
      Else
         OrderByStr = "ORDER BY PART_LOCATION_ID "
      End If
   ElseIf (Ind = 2) Or (Ind = 3) Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NAME "
      Else
         OrderByStr = "ORDER BY PART_TYPE_NO "
      End If
   ElseIf (Ind = 4) Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NAME "
      Else
         OrderByStr = "ORDER BY PART_TYPE_ID "
      End If
   ElseIf (Ind = 5) Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   PART_LOCATION_ID = NVLI(Rs("PART_LOCATION_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
   AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   LOCATION_ID = NVLD(Rs("LOCATION_ID"), -1)
   LAST_PRICE = NVLD(Rs("LAST_PRICE"), 0)
   MINIMUM_ALLOW = NVLD(Rs("MINIMUM_ALLOW"), 0)
   CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
   TEMP_AVG_PRICE = NVLD(Rs("TEMP_AVG_PRICE"), 0)
   TEMP_LAST_PRICE = NVLD(Rs("TEMP_LAST_PRICE"), 0)
   TEMP_CURRENT_AMOUNT = NVLD(Rs("TEMP_CURRENT_AMOUNT"), 0)
   
   LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = LAST_PRICE
   If LAST_PRICE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = MINIMUM_ALLOW
   If MINIMUM_ALLOW <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_LOCATION_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PART_LOCATION " & _
                     "( " & _
                     "   PART_LOCATION_ID, " & _
                     "   AVG_PRICE, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   LAST_PRICE, " & _
                     "   MINIMUM_ALLOW, " & _
                     "   CURRENT_AMOUNT, " & _
                     "   TEMP_LAST_PRICE, " & _
                     "   TEMP_AVG_PRICE, " & _
                     "   TEMP_CURRENT_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & AVG_PRICE & "," & _
                         "'" & ChangeQuote(Trim(PART_ITEM_ID)) & "'," & _
                         "" & TempID1 & "," & _
                         "" & TempID2 & "," & _
                         "" & TempID3 & "," & _
                         CURRENT_AMOUNT & "," & _
                         "" & TEMP_LAST_PRICE & "," & _
                         "" & TEMP_AVG_PRICE & "," & _
                         "" & TEMP_CURRENT_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               PART_LOCATION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_LOCATION SET " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                     "   PART_ITEM_ID = '" & ChangeQuote(Trim(PART_ITEM_ID)) & "', " & _
                     "   LOCATION_ID = " & TempID1 & ", " & _
                     "   LAST_PRICE = " & TempID2 & ", " & _
                     "   MINIMUM_ALLOW = " & TempID3 & ", " & _
                     "   CURRENT_AMOUNT = " & CURRENT_AMOUNT & ", " & _
                     "   TEMP_LAST_PRICE = " & TEMP_LAST_PRICE & ", " & _
                     "   TEMP_CURRENT_AMOUNT = " & TEMP_CURRENT_AMOUNT & ", " & _
                     "   TEMP_AVG_PRICE = " & TEMP_AVG_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_LOCATION_ID = " & PART_LOCATION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PART_LOCATION WHERE PART_LOCATION_ID = " & PART_LOCATION_ID
      m_Conn.Execute (SQL1)
End Function

