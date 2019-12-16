VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomerAccountList"
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

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
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
   GetFieldValue = Tf.getValue
End Function

Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.setValue (Value)
End Sub
Public Sub AddEditData()
Dim SQL As String
Dim Tf As CTableField
Dim NET_AMOUNT As Double
Dim FEE_AMOUNT As Double
Dim AMOUNT As Double

   Set Tf = m_FieldList(PKField)
   
   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If
   '''Debug.Print (SQL)
   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
   
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.getValue
   Call m_ADOConn.Execute(SQL)
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

Dim CUSTOMER_ACCOUNT_LIST_ID As Long
Dim CUSTOMER_ID As Long
Dim ACCOUNT_LIST_TYPE As Long
Dim PART_GROUP_ID As Long
Dim FEATURE_TYPE As Long
Dim BANK_ACCOUNT_ID As String
Dim DEBIT_ID As String
Dim CREDIT_ID As String
Dim CUSTOMER_CODE As String
Dim FEATURE_TYPE_NO As String
Dim PART_GROUP_NO As String

Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   CUSTOMER_ACCOUNT_LIST_ID = Val(GetFieldValue("CUSTOMER_ACCOUNT_LIST_ID"))
   CUSTOMER_ID = Val(GetFieldValue("CUSTOMER_ID"))
   ACCOUNT_LIST_TYPE = Val(GetFieldValue("ACCOUNT_LIST_TYPE"))
   PART_GROUP_ID = Val(GetFieldValue("PART_GROUP_ID"))
   FEATURE_TYPE = Val(GetFieldValue("FEATURE_TYPE"))
   BANK_ACCOUNT_ID = Val(GetFieldValue("BANK_ACCOUNT_ID"))
   DEBIT_ID = Val(GetFieldValue("DEBIT_ID"))
   CREDIT_ID = Val(GetFieldValue("CREDIT_ID"))
   
   CUSTOMER_CODE = GetFieldValue("CUSTOMER_CODE")
   FEATURE_TYPE_NO = GetFieldValue("FEATURE_TYPE_NO")
   PART_GROUP_NO = GetFieldValue("PART_GROUP_NO")
   
   
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",PG.PART_GROUP_NO,PG.PART_GROUP_NAME "
      SelectStr1 = SelectStr1 & ",FT.FEATURE_TYPE_NO,FT.FEATURE_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",MR3.KEY_CODE BANK_ACCOUNT_NO,MR3.KEY_NAME BANK_ACCOUNT_NAME "
      SelectStr1 = SelectStr1 & ",MR4.KEY_CODE DEBIT_NO,MR4.KEY_NAME DEBIT_NAME "
      SelectStr1 = SelectStr1 & ",MR5.KEY_CODE CREDIT_NO,MR5.KEY_NAME CREDIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (NM.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE_TYPE FT ON (NM.FEATURE_TYPE = FT.FEATURE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON (NM.BANK_ACCOUNT_ID = MR3.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR4 ON (NM.DEBIT_ID = MR4.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR5 ON (NM.CREDIT_ID = MR5.KEY_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then           ' for Patch_2007_05_29_2_jill
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = NM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE_TYPE FT ON (NM.FEATURE_TYPE = FT.FEATURE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (NM.PART_GROUP_ID = PG.PART_GROUP_ID) "
      
   End If
   
   If CUSTOMER_ACCOUNT_LIST_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_ACCOUNT_LIST_ID = " & CUSTOMER_ACCOUNT_LIST_ID & ") " & vbCrLf
   End If
   
   If ACCOUNT_LIST_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.ACCOUNT_LIST_TYPE = " & ACCOUNT_LIST_TYPE & ") " & vbCrLf
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PART_GROUP_ID = " & PART_GROUP_ID & ") " & vbCrLf
   End If
      
   If FEATURE_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.FEATURE_TYPE = " & FEATURE_TYPE & ") " & vbCrLf
   End If
   
   If BANK_ACCOUNT_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BANK_ACCOUNT_ID = " & BANK_ACCOUNT_ID & ") " & vbCrLf
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_ID = " & CUSTOMER_ID & ") " & vbCrLf
   End If
   
   If Trim(CUSTOMER_CODE) <> "" Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CM.CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CM.CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CM.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   
   If Trim(FEATURE_TYPE_NO) <> "" Then
      If WildCard(FEATURE_TYPE_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FT.FEATURE_TYPE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FT.FEATURE_TYPE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FT.FEATURE_TYPE_NO = '" & ChangeQuote(Trim(FEATURE_TYPE_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (FT.FEATURE_TYPE_NO = '" & ChangeQuote(Trim(FEATURE_TYPE_NO)) & "')"
         End If
      End If
   End If
   
   If Trim(PART_GROUP_NO) <> "" Then
      If WildCard(PART_GROUP_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PG.PART_GROUP_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PG.PART_GROUP_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PG.PART_GROUP_NO = '" & ChangeQuote(Trim(PART_GROUP_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PG.PART_GROUP_NO = '" & ChangeQuote(Trim(PART_GROUP_NO)) & "')"
         End If
      End If
   End If
   
   GroupByStr = ""
   If Ind = 3 Then
   '   GroupByStr = "GROUP BY NM.BANK_ACCOUNT, NM.TX_TYPE "
   End If
   
   If ORDER_BY = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_ACCOUNT_LIST_ID "
   Else
      OrderByStr = "ORDER BY CUSTOMER_ACCOUNT_LIST_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CUSTOMER_ACCOUNT_LIST_ID", NVLI(Rs("CUSTOMER_ACCOUNT_LIST_ID"), -1))
      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
      Call SetFieldValue("ACCOUNT_LIST_TYPE", NVLI(Rs("ACCOUNT_LIST_TYPE"), -1))
      
      Call SetFieldValue("PART_GROUP_ID", NVLI(Rs("PART_GROUP_ID"), -1))
      Call SetFieldValue("FEATURE_TYPE", NVLI(Rs("FEATURE_TYPE"), -1))
      Call SetFieldValue("BANK_ACCOUNT_ID", NVLI(Rs("BANK_ACCOUNT_ID"), -1))
      
      Call SetFieldValue("DEBIT_ID", NVLI(Rs("DEBIT_ID"), -1))
      Call SetFieldValue("CREDIT_ID", NVLI(Rs("CREDIT_ID"), -1))
      
      Call SetFieldValue("PART_GROUP_NO", NVLS(Rs("PART_GROUP_NO"), ""))
      Call SetFieldValue("PART_GROUP_NAME", NVLS(Rs("PART_GROUP_NAME"), ""))
      Call SetFieldValue("FEATURE_TYPE_NO", NVLS(Rs("FEATURE_TYPE_NO"), ""))
      Call SetFieldValue("FEATURE_TYPE_NAME", NVLS(Rs("FEATURE_TYPE_NAME"), ""))
      Call SetFieldValue("BANK_ACCOUNT_NO", NVLS(Rs("BANK_ACCOUNT_NO"), ""))
      Call SetFieldValue("BANK_ACCOUNT_NAME", NVLS(Rs("BANK_ACCOUNT_NAME"), ""))
      Call SetFieldValue("DEBIT_NO", NVLS(Rs("DEBIT_NO"), ""))
      Call SetFieldValue("DEBIT_NAME", NVLS(Rs("DEBIT_NAME"), ""))
      Call SetFieldValue("CREDIT_NO", NVLS(Rs("CREDIT_NO"), ""))
      Call SetFieldValue("CREDIT_NAME", NVLS(Rs("CREDIT_NAME"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("CUSTOMER_ACCOUNT_LIST_ID", NVLI(Rs("CUSTOMER_ACCOUNT_LIST_ID"), -1))
      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
      Call SetFieldValue("ACCOUNT_LIST_TYPE", NVLI(Rs("ACCOUNT_LIST_TYPE"), -1))
      
      Call SetFieldValue("PART_GROUP_ID", NVLI(Rs("PART_GROUP_ID"), -1))
      Call SetFieldValue("FEATURE_TYPE", NVLI(Rs("FEATURE_TYPE"), -1))
      Call SetFieldValue("BANK_ACCOUNT_ID", NVLI(Rs("BANK_ACCOUNT_ID"), -1))
      
      Call SetFieldValue("DEBIT_ID", NVLI(Rs("DEBIT_ID"), -1))
      Call SetFieldValue("CREDIT_ID", NVLI(Rs("CREDIT_ID"), -1))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CUSTOMER_ACCOUNT_LIST"
   SequenceName = "CUSTOMER_ACCOUNT_LIST_SEQ"
   PKField = "CUSTOMER_ACCOUNT_LIST_ID"
   
   Call AddField("CUSTOMER_ACCOUNT_LIST_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("CUSTOMER_ID", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("ACCOUNT_LIST_TYPE", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("PART_GROUP_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("FEATURE_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("BANK_ACCOUNT_ID", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("DEBIT_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CREDIT_ID", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("PART_GROUP_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PART_GROUP_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("FEATURE_TYPE_NO", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("BANK_ACCOUNT_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("BANK_ACCOUNT_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("DEBIT_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("DEBIT_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CREDIT_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CREDIT_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CUSTOMER_CODE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("TEMP_CODE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("FEATURE_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
   
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
End Sub
