VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLot"
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
Public LOT_ID As Long
Public LOT_NO As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DATE_IN As String
Public PART_ITEM_ID As Long
Public DOCUMENT_DATE As Date
Public LOT_DATE As Date
Public CREATE_DATE As Date
Public START_DATE As Date
Public LOT_AMOUNT As Long
Public LOT_PAYOFF As Long
Public LOT_BAL As Long
Public Pallets  As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Pallets = New Collection
   FROM_DATE = -1
   TO_DATE = -1
   START_DATE = -1
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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

   If Ind = 1 Then ' ใช้แล้ว
      SelectStr1 = "SELECT LT.* "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT LT "
   ElseIf Ind = 4 Then ' ใช้แล้ว
      SelectStr1 = "SELECT DISTINCT LT.LOT_NO "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM LOT LT "
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LTW.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LTW.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LT.LOT_ID = " & LOT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LT.LOT_ID = " & LOT_ID & ") "
      End If
   End If
   
   If Len(LOT_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LT.LOT_NO = '" & LOT_NO & "') "
      Else
         WhereStr = WhereStr & "AND (LT.LOT_NO = '" & LOT_NO & "') "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LT.LOT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LT.LOT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LT.LOT_DATE  <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LT.LOT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
      If Len(DATE_IN) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (LT.LOT_DATE  IN  " & Trim(DATE_IN) & ")"
      Else
         WhereStr = WhereStr & "AND (LT.LOT_DATE IN " & Trim(DATE_IN) & ")"
      End If
   End If
   
   GroupByStr = ""
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Or Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LT.LOT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LT.LOT_NO "
      End If
   Else
      OrderByStr = "ORDER BY LT.LOT_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
'      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         LOT_DATE = InternalDateToDate(NVLS(Rs("LOT_DATE"), ""))
   ElseIf Ind = 2 Then
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         LOT_NO = NVLS(Rs("LOT_NO"), "")
         START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
   ElseIf Ind = 3 Then
         LOT_ID = NVLI(Rs("LOT_ID"), -1)
         CREATE_DATE = InternalDateToDate(NVLS(Rs("CREATE_DATE"), ""))
   ElseIf Ind = 4 Then
         LOT_NO = NVLS(Rs("LOT_NO"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LOT_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO LOT " & _
                     "( " & _
                     "LOT_ID, " & _
                     "LOT_NO, " & _
                     "LOT_DATE, " & _
                      "CREATE_DATE, " & _
                    "CREATE_BY, " & _
                    "MODIFY_DATE, " & _
                    "MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     "'" & LOT_NO & "'," & _
                     "'" & ChangeQuote(Trim(DateToStringInt(LOT_DATE))) & "'," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
                     LOT_ID = ID
   Else 'Show edit
                     SQL1 = "UPDATE LOT SET " & _
                     "  LOT_NO = '" & LOT_NO & "', " & _
                     "  LOT_DATE = '" & ChangeQuote(Trim(DateToStringInt(LOT_DATE))) & "', " & _
                     "  MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "  MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "  WHERE LOT_ID = " & LOT_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
   SQL1 = "DELETE FROM LOT WHERE LOT_ID = " & LOT_ID
   m_Conn.Execute (SQL1)
End Function
Public Function UpdateLotDate()
Dim SQL1 As String
                     SQL1 = "UPDATE LOT SET " & _
                     "  LOT_DATE = '" & ChangeQuote(Trim(DateToStringInt(LOT_DATE))) & "'" & _
                     "  WHERE LOT_ID = " & LOT_ID
                     
                     m_Conn.Execute (SQL1)
End Function


Private Sub Class_Terminate()
   Set Pallets = Nothing
End Sub
