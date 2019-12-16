VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoBank"
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

Public MEMO_BANK_ID As Long
Public BILLING_DOC_ID As Long
Public MEMO_BANK_NO As String
Public DESCRIPTION As String
Public MEMO_BANK_DATE As Date
Public COEFFICIENT As Double
Public AMOUNT_THAI As Double
Public AMOUNT_OTHER As Double
Public TAX As Double
Public EXCHANGE_DATE As Date
Public UNIT As Long

Public TICKET_TYPE As String
Public INVOICE_NO As String
Public CURRENCY_OTHER As String
Public Ratio As Double
Public COST_BAHT As Double
Public TAX_PERCENT As Double
Public TAX_IN As Double
Public VAT As Double
Public PAID_FLAG As String




Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   MEMO_BANK_DATE = -1
   PAID_FLAG = "N"
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
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MEMO_BANK UG "
      'FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
   
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY MEMO_BANK_DATE "
       Else
         OrderByStr = "ORDER BY MEMO_BANK_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      
      MEMO_BANK_ID = NVLI(Rs("MEMO_BANK_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      MEMO_BANK_NO = NVLS(Rs("MEMO_BANK_NO"), "")
      DESCRIPTION = NVLS(Rs("DESCRIPTION"), "")
      MEMO_BANK_DATE = NVLS(Rs("MEMO_BANK_DATE"), "")
      EXCHANGE_DATE = NVLS(Rs("EXCHANGE_DATE"), "")
'      MEMO_BANK_DATE = InternalDateToDate(NVLS(Rs("MEMO_BANK_DATE"), ""))
'     EXCHANGE_DATE = InternalDateToDate(NVLS(Rs("EXCHANGE_DATE"), ""))
      AMOUNT_THAI = NVLD(Rs("AMOUNT_THAI"), 0)
      AMOUNT_OTHER = NVLD(Rs("AMOUNT_OTHER"), 0)
      COEFFICIENT = NVLD(Rs("COEFFICIENT"), 0)
      TAX = NVLD(Rs("TAX"), 0)
      UNIT = NVLD(Rs("UNIT"), 0)
      
      TICKET_TYPE = NVLS(Rs("TICKET_TYPE"), "")
      INVOICE_NO = NVLS(Rs("INVOICE_NO"), "")
      CURRENCY_OTHER = NVLS(Rs("CURRENCY_OTHER"), "")
      Ratio = NVLD(Rs("RATIO"), 0)
      COST_BAHT = NVLD(Rs("COST_BAHT"), 0)
      TAX_PERCENT = NVLD(Rs("TAX_PERCENT"), 0)
      TAX_IN = NVLD(Rs("TAX_IN"), 0)
      VAT = NVLD(Rs("VAT"), 0)
      PAID_FLAG = NVLS(Rs("PAID_FLAG"), "")
      
         End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim ID2 As Long
Dim TempID1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MEMO_BANK_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO MEMO_BANK " & _
                     "( " & _
                     "   MEMO_BANK_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   MEMO_BANK_NO, " & _
                     "   DESCRIPTION, " & _
                     "   MEMO_BANK_DATE, " & _
                     "   EXCHANGE_DATE, " & _
                     "   COEFFICIENT, " & _
                     "   AMOUNT_THAI, " & _
                     "   AMOUNT_OTHER, " & _
                     "   TAX, " & _
                     "   UNIT, "
     SQL1 = SQL1 & _
                      "   TICKET_TYPE, " & _
                     "   INVOICE_NO, " & _
                     "   CURRENCY_OTHER, " & _
                     "   RATIO, " & _
                     "   COST_BAHT, " & _
                     "   TAX_PERCENT, " & _
                     "   TAX_IN, " & _
                     "   VAT, " & _
                     "   PAID_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
            SQL1 = SQL1 & _
                     "( " & _
                         id & "," & _
                         "" & BILLING_DOC_ID & "," & _
                           "'" & ChangeQuote(Trim(MEMO_BANK_NO)) & "', " & _
                           "'" & ChangeQuote(Trim(DESCRIPTION)) & "', " & _
                          "'" & ChangeQuote(Trim(MEMO_BANK_DATE)) & "'," & _
                          "'" & ChangeQuote(Trim(EXCHANGE_DATE)) & "'," & _
                         "" & COEFFICIENT & "," & _
                         "" & AMOUNT_THAI & "," & _
                         "" & AMOUNT_OTHER & "," & _
                         "" & TAX & "," & _
                         "" & UNIT & ","
            SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(TICKET_TYPE)) & "', " & _
                          "'" & ChangeQuote(Trim(INVOICE_NO)) & "'," & _
                          "'" & ChangeQuote(Trim(CURRENCY_OTHER)) & "'," & _
                         "" & Ratio & "," & _
                         "" & COST_BAHT & "," & _
                         "" & TAX_PERCENT & "," & _
                         "" & TAX_IN & "," & _
                         "" & VAT & "," & _
                         "'" & ChangeQuote(Trim(PAID_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         MEMO_BANK_ID = id
   Else
      SQL1 = "UPDATE MEMO_BANK SET " & _
                    "   MEMO_BANK_NO = '" & ChangeQuote(Trim(MEMO_BANK_NO)) & "', " & _
                    "   DESCRIPTION = '" & ChangeQuote(Trim(DESCRIPTION)) & "', " & _
                    "   MEMO_BANK_DATE = '" & ChangeQuote(Trim(MEMO_BANK_DATE)) & "', " & _
                    "   EXCHANGE_DATE = '" & ChangeQuote(Trim(EXCHANGE_DATE)) & "', " & _
                     "   COEFFICIENT = " & COEFFICIENT & ", " & _
                     "   AMOUNT_THAI = " & AMOUNT_THAI & ", " & _
                     "   AMOUNT_OTHER = " & AMOUNT_OTHER & ", " & _
                     "   TAX = " & TAX & ", " & _
                     "   UNIT = " & UNIT & ", " & _
                    "   TICKET_TYPE = '" & ChangeQuote(Trim(TICKET_TYPE)) & "', " & _
                    "   INVOICE_NO = '" & ChangeQuote(Trim(INVOICE_NO)) & "', " & _
                    "   CURRENCY_OTHER = '" & ChangeQuote(Trim(CURRENCY_OTHER)) & "', " & _
                     "   RATIO = " & Ratio & ", " & _
                     "   COST_BAHT = " & COST_BAHT & ", " & _
                     "   TAX_PERCENT = " & TAX_PERCENT & ", " & _
                     "   TAX_IN = " & TAX_IN & ", " & _
                     "   VAT = " & VAT & ", " & _
                     "   PAID_FLAG = '" & ChangeQuote(Trim(PAID_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE MEMO_BANK_ID = " & MEMO_BANK_ID
   End If
   ''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM MEMO_BANK WHERE MEMO_BANK_ID = " & MEMO_BANK_ID
      m_Conn.Execute (SQL1)
End Function

