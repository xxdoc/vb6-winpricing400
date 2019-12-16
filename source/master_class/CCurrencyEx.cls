VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCurrencyEx"
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

Public CURRENCY_EX_ID As Long
Public EXCHANGE_DATE As Date
Public US As Double
Public EURO As Double
Public YEN As Double
Public SS As Double

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   EXCHANGE_DATE = -1
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
Dim TempDate As String

   SelectStr1 = "SELECT CU.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM CURRENCY_EX CU "
    
   If CURRENCY_EX_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CURRENCY_EX_ID = " & CURRENCY_EX_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CURRENCY_EX_ID = " & CURRENCY_EX_ID & ") "
      End If
   End If
   
   If EXCHANGE_DATE > -1 Then
      TempDate = DateToStringIntLow(EXCHANGE_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE (EXCHANGE_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EXCHANGE_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
    OrderByStr = "ORDER BY EXCHANGE_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY US "
    ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY EURO "
    ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY YEN "
   ElseIf OrderBy = 5 Then
      OrderByStr = "ORDER BY SS "
      Else
      OrderByStr = "ORDER BY CURRENCY_EX_ID "
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
   CURRENCY_EX_ID = NVLI(Rs("CURRENCY_EX_ID"), -1)
   EXCHANGE_DATE = InternalDateToDate(NVLS(Rs("EXCHANGE_DATE"), ""))
   US = NVLD(Rs("US"), -1)
   EURO = NVLD(Rs("EURO"), -1)
   YEN = NVLD(Rs("YEN"), -1)
   SS = NVLD(Rs("SS"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CURRENCY_EX_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO CURRENCY_EX " & _
                     "( " & _
                     "   CURRENCY_EX_ID, " & _
                     "   EXCHANGE_DATE, " & _
                     "   US, " & _
                     "   EURO, " & _
                     "   YEN, " & _
                     "   SS, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EXCHANGE_DATE))) & "'," & _
                           "'" & ChangeQuote(Trim(US)) & "'," & _
                           "'" & ChangeQuote(Trim(EURO)) & "'," & _
                           "'" & ChangeQuote(Trim(YEN)) & "'," & _
                           "'" & ChangeQuote(Trim(SS)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CURRENCY_EX_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CURRENCY_EX SET " & _
                     "   EXCHANGE_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXCHANGE_DATE))) & "', " & _
                    "  US = '" & ChangeQuote(Trim(US)) & "', " & _
                    "  EURO = '" & ChangeQuote(Trim(EURO)) & "', " & _
                    "  YEN = '" & ChangeQuote(Trim(YEN)) & "', " & _
                    "  SS = '" & ChangeQuote(Trim(SS)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CURRENCY_EX_ID = " & CURRENCY_EX_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM CURRENCY_EX WHERE CURRENCY_EX_ID = " & CURRENCY_EX_ID
      m_Conn.Execute (SQL1)
End Function


