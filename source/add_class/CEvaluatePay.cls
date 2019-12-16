VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEvaluatePay"
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

   Set Tf = m_FieldList(PKField)
   
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
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

Dim EVALUATE_PAY_ID As Long
Dim SUPPLIER_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim SUPPLIER_CODE As String

   EVALUATE_PAY_ID = Val(GetFieldValue("EVALUATE_PAY_ID"))
   SUPPLIER_ID = Val(GetFieldValue("SUPPLIER_ID"))
   SUPPLIER_CODE = GetFieldValue("SUPPLIER_CODE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ", NM1.LONG_NAME SUPPLIER_NAME "
      SelectStr1 = SelectStr1 & ", SP.SUPPLIER_CODE SUPPLIER_CODE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = NM.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      WhereStr = ""
   End If
   
   If EVALUATE_PAY_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EVALUATE_PAY_ID = " & EVALUATE_PAY_ID & ") " & vbCrLf
   End If
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.SUPPLIER_ID = " & SUPPLIER_ID & ") " & vbCrLf
   End If
      
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EVALUATE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EVALUATE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY EVALUATE_DATE,EVALUATE_PAY_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY SUPPLIER_CODE "
   Else
      OrderByStr = "ORDER BY EVALUATE_DATE,EVALUATE_PAY_ID "
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
      Call SetFieldValue("EVALUATE_PAY_ID", NVLI(Rs("EVALUATE_PAY_ID"), -1))
      Call SetFieldValue("SUPPLIER_ID", NVLI(Rs("SUPPLIER_ID"), -1))
      Call SetFieldValue("EVALUATE_DATE", InternalDateToDate(NVLS(Rs("EVALUATE_DATE"), "")))
      Call SetFieldValue("EVALUATE_AMOUNT", NVLI(Rs("EVALUATE_AMOUNT"), 0))
      Call SetFieldValue("EVALUATE_PAY_DESC", NVLS(Rs("EVALUATE_PAY_DESC"), ""))
      
      Call SetFieldValue("SUPPLIER_CODE", NVLS(Rs("SUPPLIER_CODE"), ""))
      Call SetFieldValue("SUPPLIER_NAME", NVLS(Rs("SUPPLIER_NAME"), ""))
      
   End If
End Sub

Public Sub RegisterFields()
   TableName = "EVALUATE_PAY"
   SequenceName = "EVALUATE_PAY_SEQ"
   PKField = "EVALUATE_PAY_ID"
   
   Call AddField("EVALUATE_PAY_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("SUPPLIER_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("EVALUATE_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("EVALUATE_AMOUNT", MONEY_TYPE, -1, DATA_CAT)
   Call AddField("EVALUATE_PAY_DESC", STRING_TYPE, "", DATA_CAT)
   
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
   
   Call AddField("SUPPLIER_CODE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("SUPPLIER_NAME", STRING_TYPE, "", TEMP_CAT)
   
End Sub

