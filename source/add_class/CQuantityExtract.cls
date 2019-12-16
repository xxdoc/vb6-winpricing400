VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CQuantityExtract"
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

Public QUANTITY_EXTRACT_ID As Long
Public FROM_JOB_DATE As Date
Public TO_JOB_DATE As Long
 Public PROCESS_TYPE As Long
 Public AREA As Long
 
 Public ExtractItems As Collection
 
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ExtractItems = New Collection
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
      SelectStr1 = "SELECT CM.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM QUANTITY_EXTRACT CM "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT MAX(QUANTITY_EXTRACT_ID) QUANTITY_EXTRACT_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM QUANTITY_EXTRACT CM "
      
      WhereStr = ""
   End If
   
   If QUANTITY_EXTRACT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID & ") "
      End If
   End If
   
   If AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AREA = " & AREA & ") "
      Else
         WhereStr = WhereStr & "AND (AREA = " & AREA & ") "
      End If
   End If
               
   If PROCESS_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PROCESS_TYPE = " & PROCESS_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PROCESS_TYPE = " & PROCESS_TYPE & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FROM_JOB_DATE "
      Else
         OrderByStr = "ORDER BY QUANTITY_EXTRACT_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY QUANTITY_EXTRACT_ID "
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
      QUANTITY_EXTRACT_ID = NVLI(Rs("QUANTITY_EXTRACT_ID"), -1)
      FROM_JOB_DATE = InternalDateToDate(NVLS(Rs("FROM_JOB_DATE"), ""))
      TO_JOB_DATE = InternalDateToDate(NVLS(Rs("TO_JOB_DATE"), ""))
      PROCESS_TYPE = NVLI(Rs("PROCESS_TYPE"), -1)
      AREA = NVLI(Rs("AREA"), -1)
   ElseIf Ind = 2 Then
      QUANTITY_EXTRACT_ID = NVLI(Rs("QUANTITY_EXTRACT_ID"), -1)
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

   TempID1 = TO_JOB_DATE
   If TO_JOB_DATE <= 0 Then
      TempID1 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("QUANTITY_EXTRACT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO QUANTITY_EXTRACT " & _
                     "( " & _
                     "   QUANTITY_EXTRACT_ID, " & _
                     "   FROM_JOB_DATE, " & _
                     "   TO_JOB_DATE, " & _
                     "   PROCESS_TYPE, " & _
                     "   AREA, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(FROM_JOB_DATE))) & "'," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(TO_JOB_DATE))) & "'," & _
                        PROCESS_TYPE & "," & _
                        AREA & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         QUANTITY_EXTRACT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE QUANTITY_EXTRACT SET " & _
                     "    FROM_JOB_DATE = '" & ChangeQuote(DateToStringInt(Trim(FROM_JOB_DATE))) & "', " & _
                     "    TO_JOB_DATE = '" & ChangeQuote(DateToStringInt(Trim(TO_JOB_DATE))) & "', " & _
                     "    PROCESS_TYPE = " & PROCESS_TYPE & ", " & _
                     "    AREA = " & AREA & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM QUANTITY_EXTRACT WHERE QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM EXTRACT_ITEM WHERE QUANTITY_EXTRACT_ID = " & QUANTITY_EXTRACT_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ExtractItems = Nothing
End Sub
