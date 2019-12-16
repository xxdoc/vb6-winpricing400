VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCashDocPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

Public Post2BD As Collection

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Public CASH_DOC_POST_ID As Long
Public CASH_DOC_ID As Long
Public BANK_BRANCH As Long
Public BANK_ID As Long

Public CHEQUE_NO  As String

Public CHEQUE_ID As Long
Public CHEQUE_AMOUNT As Double
Public POST_TYPE  As Long
Public BILLING_DOC_ID  As Long
Public WH_AMOUNT  As Double
Public INTERREST_AMOUNT  As Double
Public BILLING_DOC_NO As String
Public CHEQUE_SUPPLIER_ID As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public BANK_ACCOUNT As Long
Public DOCUMENT_TYPE As Long
Public ENTERPRISE_ID As Long
Public EMP_ID As Long
Public CUSTOMER_ID As Long
Public SUPPLIER_ID As Long
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Public FROM_DATE  As Date
Public TO_DATE  As Date

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
      
   Set Post2BD = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   
   Set Post2BD = Nothing
End Sub

Private Sub AddField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FName, FType, FValue, FCat)
   Call m_FieldList.add(F, FName)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldName As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   GetFieldValue = Tf.GetValue
End Function

Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.SetValue (Value)
End Sub


Public Sub AddEditData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)

   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField
Dim TempRs As ADODB.Recordset
Dim Ct As CCashTran
Dim iCount As Long
Dim Cq As CCheque

   Set Tf = m_FieldList(PKField)
        
   If GetFieldValue("POST_TYPE") = POST_CLEAR Then
      SQL = "UPDATE CHEQUE  SET  POST_FLAG = 'N'  WHERE  CHEQUE_ID  = " & GetFieldValue("CHEQUE_ID")
      Call m_ADOConn.Execute(SQL)
   ElseIf GetFieldValue("POST_TYPE") = WAITING_CLEAR Then
      SQL = "UPDATE CHEQUE  SET  BANK_FLAG = 'N'  WHERE  CHEQUE_ID  = " & GetFieldValue("CHEQUE_ID")
      Call m_ADOConn.Execute(SQL)
      
   ElseIf GetFieldValue("POST_TYPE") = PASSED_CLEAR Then
      SQL = "UPDATE CHEQUE  SET  POST_FLAG = 'N'  WHERE  CHEQUE_ID  = " & GetFieldValue("CHEQUE_ID")
      Call m_ADOConn.Execute(SQL)
   
   End If

   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   If GetFieldValue("POST_TYPE") = WAITING_CLEAR Then
      SQL = "DELETE FROM BILLING_DOC  WHERE BILLING_DOC_ID = " & GetFieldValue("BILLING_DOC_ID")
      Call m_ADOConn.Execute(SQL)
   End If
   
End Sub
Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim NewStr As String
Dim SubLen As Long
Dim TempDate As String

Dim CASH_DOC_ID As Long
Dim CASH_DOC_POST_ID As Long
Dim ORDER_TYPE As Long
   
   
   CASH_DOC_POST_ID = Val(GetFieldValue("CASH_DOC_POST_ID"))
   CASH_DOC_ID = Val(GetFieldValue("CASH_DOC_ID"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
      
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.BANK_NAME BANK_NAME, MR2.BBRANCH_NAME BRANCH_NAME,CQ.CHEQUE_NO,CQ.CHEQUE_DATE,CQ.SUPPLIER_ID CHEQUE_SUPPLIER_ID, CQ.EFFECTIVE_DATE  "
      SelectStr1 = SelectStr1 & ",BD.PV_NO BILLING_DOC_NO "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN BANK MR1 ON (NM.BANK_ID = MR1.BANK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_BRANCH MR2 ON (NM.BANK_BRANCH = MR2.BBRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE CQ ON (NM.CHEQUE_ID = CQ.CHEQUE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN CASH_TRAN CT ON (CT.CHECK_ID = CQ.CHEQUE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = CT.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT NM.*,CD.*,NM.BANK_BRANCH BANK_BRANCH,NM.BANK_ID BANK_ID,CQ.CHEQUE_NO  "
      SelectStr1 = SelectStr1 & ",BD.PV_NO BILLING_DOC_NO,CQ.SUPPLIER_ID CHEQUE_SUPPLIER_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      FromStr = FromStr & "LEFT OUTER JOIN CASH_DOC CD ON (CD.CASH_DOC_ID = NM.CASH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CHEQUE CQ ON (NM.CHEQUE_ID = CQ.CHEQUE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN CASH_TRAN CT ON (CT.CHECK_ID = CQ.CHEQUE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = CT.BILLING_DOC_ID) "
      
      WhereStr = " WHERE CD.DOCUMENT_TYPE =  " & WAITING_CHEQUE
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT CD.DOCUMENT_NO "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      FromStr = FromStr & "LEFT OUTER JOIN CASH_DOC CD ON (CD.CASH_DOC_ID = NM.CASH_DOC_ID) "
      
      WhereStr = " WHERE CD.DOCUMENT_TYPE =  " & WAITING_CHEQUE
   End If
   
    If CASH_DOC_POST_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_POST_ID = " & CASH_DOC_POST_ID & ") " & vbCrLf
   End If
   
   If CASH_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_ID = " & CASH_DOC_ID & ") " & vbCrLf
   End If
   
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY  CASH_DOC_POST_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY  CD.DOCUMENT_DATE,CD.DOCUMENT_NO,CD.CASH_DOC_ID "
      ORDER_TYPE = 1
   
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   If Ind = 3 Then
      OrderByStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CASH_DOC_POST_ID", NVLI(Rs("CASH_DOC_POST_ID"), -1))
      Call SetFieldValue("CASH_DOC_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("CHEQUE_ID", NVLI(Rs("CHEQUE_ID"), -1))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("CHEQUE_AMOUNT", NVLD(Rs("CHEQUE_AMOUNT"), 0))
      Call SetFieldValue("CHEQUE_DATE", InternalDateToDate(NVLS(Rs("CHEQUE_DATE"), "")))
      Call SetFieldValue("EFFECTIVE_DATE", InternalDateToDate(NVLS(Rs("EFFECTIVE_DATE"), "")))
      Call SetFieldValue("POST_TYPE", NVLI(Rs("POST_TYPE"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("WH_AMOUNT", NVLD(Rs("WH_AMOUNT"), 0))
      Call SetFieldValue("INTERREST_AMOUNT", NVLD(Rs("INTERREST_AMOUNT"), 0))
      
'      Call SetFieldValue("KEY_ID", NVLI(Rs("CASH_DOC_ID"), -1))
'      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("DOCUMENT_NO"), ""))
      'Call SetFieldValue("BANK_NO", NVLS(Rs("BANK_NO"), ""))
      'Call SetFieldValue("BRANCH_NO", NVLS(Rs("BRANCH_NO"), ""))
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("CHEQUE_NO", NVLS(Rs("CHEQUE_NO"), ""))
      Call SetFieldValue("CHEQUE_SUPPLIER_ID", NVLI(Rs("CHEQUE_SUPPLIER_ID"), -1))
      
      Call SetFieldValue("BILLING_DOC_NO", NVLS(Rs("BILLING_DOC_NO"), ""))
   ElseIf Ind = 2 Then
      CASH_DOC_POST_ID = NVLI(Rs("CASH_DOC_POST_ID"), -1)
      CASH_DOC_ID = NVLI(Rs("CASH_DOC_ID"), -1)
      BANK_BRANCH = NVLI(Rs("BANK_BRANCH"), -1)
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      
      CHEQUE_NO = NVLS(Rs("CHEQUE_NO"), "")
      
      CHEQUE_ID = NVLI(Rs("CHEQUE_ID"), -1)
      CHEQUE_AMOUNT = NVLD(Rs("CHEQUE_AMOUNT"), 0)
      POST_TYPE = NVLI(Rs("POST_TYPE"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      INTERREST_AMOUNT = NVLD(Rs("INTERREST_AMOUNT"), 0)
      CHEQUE_SUPPLIER_ID = NVLI(Rs("CHEQUE_SUPPLIER_ID"), -1)
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      
      BANK_ACCOUNT = NVLI(Rs("BANK_ACCOUNT"), -1)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      
      BILLING_DOC_NO = NVLS(Rs("BILLING_DOC_NO"), "")
      
   ElseIf Ind = 3 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CASH_DOC_POST"
   SequenceName = "CASH_DOC_POST_SEQ"
   PKField = "CASH_DOC_POST_ID"
   
   Call AddField("CASH_DOC_POST_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("CASH_DOC_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_BRANCH", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CHEQUE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CHEQUE_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("POST_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("BILLING_DOC_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("WH_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("INTERREST_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   
   'Call AddField("BANK_NO", STRING_TYPE, "", TEMP_CAT)
   'Call AddField("BRANCH_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BANK_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CHEQUE_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CHEQUE_DATE", DATE_TYPE, "", TEMP_CAT)
   Call AddField("CHEQUE_SUPPLIER_ID", INT_TYPE, -1, TEMP_CAT)
   Call AddField("EFFECTIVE_DATE", DATE_TYPE, "", TEMP_CAT)
   
   Call AddField("BILLING_DOC_NO", STRING_TYPE, "", TEMP_CAT)
End Sub
Public Sub GenerateCpHeader(FileID As Long)
Dim TempStr As String
   
   TempStr = "CD;"
   
   TempStr = TempStr & DOCUMENT_NO & ";"
   TempStr = TempStr & DateToStringInt(DOCUMENT_DATE) & ";"
   TempStr = TempStr & BANK_ACCOUNT & ";"
   TempStr = TempStr & DOCUMENT_TYPE & ";"
   TempStr = TempStr & ENTERPRISE_ID & ";"
   TempStr = TempStr & EMP_ID & ";"
   TempStr = TempStr & CUSTOMER_ID & ";"
   TempStr = TempStr & SUPPLIER_ID & ";"
   
   Print #FileID, TempStr
End Sub
Public Sub GenerateCpTailer(FileID As Long)
Dim TempStr As String
   
   TempStr = "CP;"
   
   TempStr = TempStr & CHEQUE_NO & ";"
   TempStr = TempStr & CHEQUE_ID & ";"
   TempStr = TempStr & BANK_BRANCH & ";"
   TempStr = TempStr & BANK_ID & ";"
   TempStr = TempStr & CHEQUE_AMOUNT & ";"
   TempStr = TempStr & POST_TYPE & ";"
   TempStr = TempStr & BILLING_DOC_ID & ";"
   TempStr = TempStr & WH_AMOUNT & ";"
   TempStr = TempStr & INTERREST_AMOUNT & ";"
   TempStr = TempStr & BILLING_DOC_NO & ";"
   TempStr = TempStr & CHEQUE_SUPPLIER_ID & ";"
   
   Print #FileID, TempStr
End Sub
