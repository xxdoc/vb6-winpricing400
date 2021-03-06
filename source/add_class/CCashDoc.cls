VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCashDoc"
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

Public CashTranItems As Collection
Public TransferItems As Collection
Public JournalItems As Collection
Public PostItems As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set CashTranItems = New Collection
   Set TransferItems = New Collection
   Set JournalItems = New Collection
   Set PostItems = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set CashTranItems = Nothing
   Set TransferItems = Nothing
   Set JournalItems = Nothing
   Set PostItems = Nothing
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
   Call SetFieldValue("ENTERPRISE_ID", glbUser.ENTERPRISE_ID)
   
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
Dim CP As CCashDocPost

   Set Tf = m_FieldList(PKField)
   
'   SQL = "DELETE FROM JOURNAL_ITEM WHERE " & PKField & " = " & Tf.GetValue
'   Call m_ADOConn.Execute(SQL)
   
   Set TempRs = New ADODB.Recordset
   Set Ct = New CCashTran
   Call Ct.SetFieldValue("CASH_DOC_ID", Tf.GetValue)
   Call Ct.QueryData(1, TempRs, iCount)
   Set Ct = Nothing
   While Not TempRs.EOF
      Set Ct = New CCashTran
      Call Ct.PopulateFromRS(1, TempRs)
      Call Ct.DeleteData
      If Ct.GetFieldValue("CHECK_ID") > 0 Then
         Set Cq = New CCheque
         Call Cq.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
         Call Cq.SetFieldValue("BANK_FLAG", "N")
         Call Cq.SetFieldValue("POST_FLAG", "N")
         Call Cq.UpdateBankFlag
         Set Cq = Nothing
      End If
      Set Ct = Nothing
      Call TempRs.MoveNext
   Wend
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   If GetFieldValue("DOCUMENT_TYPE") = POST_CHEQUE Then
      SQL = "UPDATE CHEQUE CH SET CH.POST_FLAG = 'N',CH.BANK_FLAG = 'Y'  WHERE CH.CHEQUE_ID IN (SELECT CP.CHEQUE_ID FROM CASH_DOC_POST CP WHERE CP.CASH_DOC_ID = " & Tf.GetValue & ")"
      Call m_ADOConn.Execute(SQL)
   ElseIf GetFieldValue("DOCUMENT_TYPE") = WAITING_CHEQUE Then
      SQL = "UPDATE CHEQUE CH SET CH.POST_FLAG = 'N',CH.BANK_FLAG = 'N'  WHERE CH.CHEQUE_ID IN (SELECT CP.CHEQUE_ID FROM CASH_DOC_POST CP WHERE CP.CASH_DOC_ID = " & Tf.GetValue & ")"
      Call m_ADOConn.Execute(SQL)
      
   ElseIf GetFieldValue("DOCUMENT_TYPE") = PASSED_CHEQUE Then
      SQL = "UPDATE CHEQUE CH SET CH.POST_FLAG = 'N',CH.BANK_FLAG = 'Y'  WHERE CH.CHEQUE_ID IN (SELECT CP.CHEQUE_ID FROM CASH_DOC_POST CP WHERE CP.CASH_DOC_ID = " & Tf.GetValue & ")"
      Call m_ADOConn.Execute(SQL)
   End If
   
   If GetFieldValue("DOCUMENT_TYPE") = WAITING_CHEQUE Then
      
      Set TempRs = New ADODB.Recordset
      Set CP = New CCashDocPost
      Call CP.SetFieldValue("CASH_DOC_ID", Tf.GetValue)
      Call CP.QueryData(1, TempRs, iCount)
      Set CP = Nothing
      While Not TempRs.EOF
         Set CP = New CCashDocPost
         Call CP.PopulateFromRS(1, TempRs)
         Call CP.DeleteData
         Set CP = Nothing
         Call TempRs.MoveNext
      Wend
      If TempRs.State = adStateOpen Then
         Call TempRs.Close
      End If
      Set TempRs = Nothing
   
   End If
   
   SQL = "DELETE FROM CASH_DOC_POST  WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
End Sub

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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

Dim DOCUMENT_NO As String
Dim CASH_DOC_ID As Long
Dim DOCUMENT_TYPE As Long
Dim ENP_ID As Long
Dim ACCOUNT_NAME As String
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim CHEQUE_SEARCH_NO As String
Dim TempDate As String
   
   CASH_DOC_ID = Val(GetFieldValue("CASH_DOC_ID"))
   DOCUMENT_TYPE = Val(GetFieldValue("DOCUMENT_TYPE"))
   DOCUMENT_NO = GetFieldValue("DOCUMENT_NO")
   ACCOUNT_NAME = GetFieldValue("ACCOUNT_NAME")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   ENP_ID = glbUser.ENTERPRISE_ID
   CHEQUE_SEARCH_NO = GetFieldValue("CHEQUE_SEARCH_NO")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*, MR1.BANK_NAME BANK_NAME, MR2.BBRANCH_NAME BRANCH_NAME, MR3.KEY_NAME ACCOUNT_NAME, NM1.LONG_NAME EMP_NAME, NM1.LAST_NAME EMP_LASTNAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN BANK MR1 ON (NM.BANK_ID = MR1.BANK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_BRANCH MR2 ON (NM.BANK_BRANCH = MR2.BBRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.BANK_ACCOUNT = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (NM.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (EN.NAME_ID = NM1.NAME_ID) "
      
      WhereStr = ""
   End If
   
   If CASH_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CASH_DOC_ID = " & CASH_DOC_ID & ") " & vbCrLf
   End If
  
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") " & vbCrLf
   End If
      
'   If Len(Trim(COMMIT_FLAG)) > 0 Then
'      If Len(WhereStr) > 0 Then
'         TempStr = "AND "
'      Else
'         TempStr = "WHERE "
'      End If
'      If WildCard(COMMIT_FLAG, SubLen, NewStr) Then
'         WhereStr = WhereStr & TempStr & " (SUBSTR(COMMIT_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
'      Else
'         WhereStr = WhereStr & TempStr & " (COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "')"
'      End If
'   End If
   
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
      
   If ENP_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.ENTERPRISE_ID = " & ENP_ID & ") " & vbCrLf
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
   
   If Len(CHEQUE_SEARCH_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(CHEQUE_SEARCH_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM CASH_DOC_POST DI LEFT OUTER JOIN CHEQUE PI ON (DI.CHEQUE_ID = PI.CHEQUE_ID) WHERE (DI.CASH_DOC_ID = NM.CASH_DOC_ID) AND (SUBSTR(PI.CHEQUE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM CASH_DOC_POST DI LEFT OUTER JOIN CHEQUE PI ON (DI.CHEQUE_ID = PI.CHEQUE_ID) WHERE (DI.CASH_DOC_ID = NM.CASH_DOC_ID) AND (PI.CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_SEARCH_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(CHEQUE_SEARCH_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM CASH_DOC_POST DI LEFT OUTER JOIN CHEQUE PI ON (DI.CHEQUE_ID = PI.CHEQUE_ID) WHERE (DI.CASH_DOC_ID = NM.CASH_DOC_ID) AND (SUBSTR(PI.CHEQUE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM CASH_DOC_POST DI LEFT OUTER JOIN CHEQUE PI ON (DI.CHEQUE_ID = PI.CHEQUE_ID) WHERE (DI.CASH_DOC_ID = NM.CASH_DOC_ID) AND (PI.CHEQUE_NO = '" & ChangeQuote(Trim(CHEQUE_SEARCH_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If Len(Trim(ACCOUNT_NAME)) > 0 Then
      If WildCard(ACCOUNT_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(MR3.KEY_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(MR3.KEY_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (MR3.KEY_NAME = '" & ChangeQuote(Trim(ACCOUNT_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (MR3.KEY_NAME = '" & ChangeQuote(Trim(ACCOUNT_NAME)) & "')"
         End If
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY DOCUMENT_DATE DESC, CASH_DOC_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CASH_DOC_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("DOCUMENT_DATE", InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), "")))
      Call SetFieldValue("DOCUMENT_NO", NVLS(Rs("DOCUMENT_NO"), ""))
      Call SetFieldValue("BANK_BRANCH", NVLI(Rs("BANK_BRANCH"), -1))
      Call SetFieldValue("BANK_ID", NVLI(Rs("BANK_ID"), -1))
      Call SetFieldValue("BANK_ACCOUNT", NVLI(Rs("BANK_ACCOUNT"), -1))
      Call SetFieldValue("DOCUMENT_TYPE", NVLI(Rs("DOCUMENT_TYPE"), -1))
      Call SetFieldValue("ENTERPRISE_ID", NVLI(Rs("ENTERPRISE_ID"), -1))
      Call SetFieldValue("EMP_ID", NVLI(Rs("EMP_ID"), -1))
      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
      Call SetFieldValue("SUPPLIER_ID", NVLI(Rs("SUPPLIER_ID"), -1))
      
      Call SetFieldValue("KEY_ID", NVLI(Rs("CASH_DOC_ID"), -1))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("DOCUMENT_NO"), ""))
      Call SetFieldValue("BANK_NAME", NVLS(Rs("BANK_NAME"), ""))
      Call SetFieldValue("BRANCH_NAME", NVLS(Rs("BRANCH_NAME"), ""))
      Call SetFieldValue("ACCOUNT_NAME", NVLS(Rs("ACCOUNT_NAME"), ""))
      Call SetFieldValue("EMP_NAME", NVLS(Rs("EMP_NAME"), ""))
      Call SetFieldValue("EMP_LASTNAME", NVLS(Rs("EMP_LASTNAME"), ""))
      Call SetFieldValue("EMP_NAME_LASTNAME", NVLS(Rs("EMP_NAME"), "") & " " & NVLS(Rs("EMP_LASTNAME"), ""))
      
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CASH_DOC"
   SequenceName = "CASH_DOC_SEQ"
   PKField = "CASH_DOC_ID"
   
   Call AddField("CASH_DOC_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("DOCUMENT_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("DOCUMENT_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("BANK_BRANCH", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ACCOUNT", INT_TYPE, -1, DATA_CAT)
   Call AddField("DOCUMENT_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("ENTERPRISE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("EMP_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CUSTOMER_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("SUPPLIER_ID", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("ACCOUNT_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BANK_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BRANCH_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_LASTNAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("EMP_NAME_LASTNAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CHEQUE_SEARCH_NO", STRING_TYPE, "", TEMP_CAT)
   
   
End Sub
