VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillingPayment"
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
Public ExtraFlag As Long
' End of generic data member

Public CheckTotalAmount  As Boolean
Public FindPayType As Boolean
Public PrintReportNormalPO2 As Boolean

Public BILLING_PAYMENT_ID As Long
Public BILLING_PAYMENT_ID_REF As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE  As Date
Public DOCUMENT_NO_JV As String
Public DOCUMENT_DATE_JV As Date
Public DOCUMENT_TYPE As Long
Public PAYMENT_TO As String
Public DOC_ASSEMBLE As String
Public PAYMENT_COST As String
Public PAYMENT_DUE As String
Public PAYMENT_PART As String
Public PAYMENT_DEPT As String
Public PAYMENT_DEPT2 As String
Public PAYMENT_AMOUNT As Double
Public PAYMENT_BY As Long
Public NOTE As String

Public FROM_DATE As Date
Public TO_DATE As Date

Public RUNNING_NO As Long
Public RUNNING_NO2 As Long
Public CONFIG_DOC_TYPE As CONFIG_DOC_TYPE
Public CONFIG_DOC_TYPE2 As CONFIG_DOC_TYPE

Public BillingPaymentItem As Collection
Public GlDetails As Collection
Public GlDetails2 As Collection


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Set BillingPaymentItem = New Collection
   Set GlDetails = New Collection
   Set GlDetails2 = New Collection

   DOCUMENT_DATE = -1
   FROM_DATE = -1
   TO_DATE = -1
   
   DOCUMENT_DATE_JV = -1

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
Dim GroupByStr As String

  If Ind = 1 Then
      SelectStr1 = "SELECT BP.*,BP2.DOCUMENT_NO DOCUMENT_NO_JV,BP2.DOCUMENT_DATE DOCUMENT_DATE_JV "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_PAYMENT BP "
      FromStr = FromStr & "LEFT JOIN BILLING_PAYMENT BP2 ON (BP.BILLING_PAYMENT_ID_REF=BP2.BILLING_PAYMENT_ID)  "
      
      WhereStr = ""
     ElseIf Ind = 2 Then
      SelectStr1 = "SELECT BP2.*,BP.DOCUMENT_NO DOCUMENT_NO_JV,BP.DOCUMENT_DATE DOCUMENT_DATE_JV "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_PAYMENT BP2 "
      FromStr = FromStr & "LEFT JOIN BILLING_PAYMENT BP ON (BP2.BILLING_PAYMENT_ID_REF=BP.BILLING_PAYMENT_ID)  "
      
      WhereStr = ""
   End If
      
   If BILLING_PAYMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BP.BILLING_PAYMENT_ID = " & BILLING_PAYMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BP.BILLING_PAYMENT_ID = " & BILLING_PAYMENT_ID & ") "
      End If
   End If
   
   
   If Len(DOCUMENT_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BP.DOCUMENT_NO LIKE '" & DOCUMENT_NO & "') "
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_NO LIKE '" & DOCUMENT_NO & "') "
      End If
   End If
   
   If DOCUMENT_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(DOCUMENT_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BP.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Len(DOCUMENT_NO_JV) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BP.DOCUMENT_NO LIKE '" & DOCUMENT_NO_JV & "') "
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_NO LIKE '" & DOCUMENT_NO_JV & "') "
      End If
   End If
   
   If DOCUMENT_DATE_JV > -1 Then
      TempDate = DateToStringIntLow(Trim(DOCUMENT_DATE_JV))
      If WhereStr = "" Then
         WhereStr = "WHERE (BP.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BP.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BP.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BP.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (BP.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   GroupByStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BP.DOCUMENT_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY BP.DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY BP.DOCUMENT_DATE DESC, BP.DOCUMENT_NO "
      End If
   End If
   
   If Ind = 1 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BILLING_PAYMENT_ID = NVLI(Rs("BILLING_PAYMENT_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      BILLING_PAYMENT_ID_REF = NVLI(Rs("BILLING_PAYMENT_ID_REF"), -1)
      DOCUMENT_NO_JV = NVLS(Rs("DOCUMENT_NO_JV"), "")
      DOCUMENT_DATE_JV = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE_JV"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      PAYMENT_TO = NVLS(Rs("PAYMENT_TO"), "")
      DOC_ASSEMBLE = NVLS(Rs("DOC_ASSEMBLE"), "")
      PAYMENT_COST = NVLS(Rs("PAYMENT_COST"), "")
      PAYMENT_DUE = NVLS(Rs("PAYMENT_DUE"), "")
      PAYMENT_PART = NVLS(Rs("PAYMENT_PART"), "")
      PAYMENT_DEPT = NVLS(Rs("PAYMENT_DEPT"), "")
      PAYMENT_DEPT2 = NVLS(Rs("PAYMENT_DEPT2"), "")
      PAYMENT_AMOUNT = NVLD(Rs("PAYMENT_AMOUNT"), 0)
      PAYMENT_BY = NVLI(Rs("PAYMENT_BY"), -1)
      NOTE = NVLS(Rs("NOTE"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID0 As String

   TempID0 = BILLING_PAYMENT_ID_REF
   If BILLING_PAYMENT_ID_REF <= 0 Then
      TempID0 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILLING_PAYMENT_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO BILLING_PAYMENT " & _
                     "( " & _
                     "   BILLING_PAYMENT_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   PAYMENT_TO, " & _
                     "   DOC_ASSEMBLE, " & _
                     "   PAYMENT_COST, " & _
                     "   PAYMENT_DUE, " & _
                     "   PAYMENT_PART, " & _
                     "   PAYMENT_DEPT, " & _
                     "   PAYMENT_DEPT2, "
               SQL1 = SQL1 & _
                     "   PAYMENT_AMOUNT, " & _
                     "   PAYMENT_BY, " & _
                     "   NOTE, " & _
                     "   BILLING_PAYMENT_ID_REF, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY" & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         "" & DOCUMENT_TYPE & "," & _
                         "'" & ChangeQuote(Trim(PAYMENT_TO)) & "'," & _
                         "'" & ChangeQuote(Trim(DOC_ASSEMBLE)) & "'," & _
                         "'" & ChangeQuote(Trim(PAYMENT_COST)) & "'," & _
                         "'" & ChangeQuote(Trim(PAYMENT_DUE)) & "'," & _
                         "'" & ChangeQuote(Trim(PAYMENT_PART)) & "'," & _
                         "'" & ChangeQuote(Trim(PAYMENT_DEPT)) & "'," & _
                         "'" & ChangeQuote(Trim(PAYMENT_DEPT2)) & "'," & _
                         PAYMENT_AMOUNT & "," & _
                         PAYMENT_BY & "," & _
                         "'" & ChangeQuote(Trim(NOTE)) & "',"
            SQL1 = SQL1 & _
                          "" & TempID0 & "," & _
                         "'" & ServerDateTime & "'," & _
                         "'" & ChangeQuote(Trim(glbUser.USER_ID)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         "'" & ChangeQuote(Trim(glbUser.USER_ID)) & "'" & _
                     ") "
         BILLING_PAYMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BILLING_PAYMENT SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "    DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "   PAYMENT_TO  = '" & ChangeQuote(Trim(PAYMENT_TO)) & "', " & _
                     "   DOC_ASSEMBLE = '" & ChangeQuote(Trim(DOC_ASSEMBLE)) & "', " & _
                     "   PAYMENT_COST = '" & ChangeQuote(Trim(PAYMENT_COST)) & "', " & _
                     "   PAYMENT_DUE = '" & ChangeQuote(Trim(PAYMENT_DUE)) & "', " & _
                     "   PAYMENT_PART = '" & ChangeQuote(Trim(PAYMENT_PART)) & "', " & _
                     "   PAYMENT_DEPT = '" & ChangeQuote(Trim(PAYMENT_DEPT)) & "', " & _
                     "   PAYMENT_DEPT2 = '" & ChangeQuote(Trim(PAYMENT_DEPT2)) & "', " & _
                     "   PAYMENT_AMOUNT  = " & PAYMENT_AMOUNT & ", " & _
                     "   PAYMENT_BY = " & PAYMENT_BY & ", " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', "
         SQL1 = SQL1 & _
                   "   BILLING_PAYMENT_ID_REF = " & TempID0 & ", " & _
                   "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                   "   MODIFY_BY= '" & glbUser.USER_ID & "'" & _
                     " WHERE BILLING_PAYMENT_ID = " & BILLING_PAYMENT_ID
   End If
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String

   SQL1 = "DELETE FROM GL_DETAIL  WHERE BILLING_PAYMENT_ID = " & BILLING_PAYMENT_ID
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM BILLING_PAYMENT  WHERE BILLING_PAYMENT_ID = " & BILLING_PAYMENT_ID
   m_Conn.Execute (SQL1)

End Function

Private Sub Class_Terminate()
   Set BillingPaymentItem = Nothing
   Set GlDetails = Nothing
   Set GlDetails2 = Nothing
End Sub
