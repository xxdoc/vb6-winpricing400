VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGLDetail"
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

Dim GL_DETAIL_ID As Long
Dim BILLING_DOC_ID As Long
Dim BILLING_PAYMENT_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   GL_DETAIL_ID = Val(GetFieldValue("GL_DETAIL_ID"))
   BILLING_DOC_ID = Val(GetFieldValue("BILLING_DOC_ID"))
   BILLING_PAYMENT_ID = Val(GetFieldValue("BILLING_PAYMENT_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",MR.KEY_CODE GL_NO,MR.KEY_NAME GL_NAME,MR.SUM_FLAG SUM_FLAG,MR.SHOW_FLAG SHOW_FLAG "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR ON (NM.GL_ID = MR.KEY_ID) "
      
      WhereStr = ""
   End If
   
   If GL_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.GL_DETAIL_ID = " & GL_DETAIL_ID & ") " & vbCrLf
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BILLING_DOC_ID = " & BILLING_DOC_ID & ") " & vbCrLf
   End If
   
   If BILLING_PAYMENT_ID >= 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.BILLING_PAYMENT_ID = " & BILLING_PAYMENT_ID & ") " & vbCrLf
   End If
      
   If Ind = 1 Then
      OrderByStr = "ORDER BY GL_TYPE "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("GL_DETAIL_ID", NVLI(Rs("GL_DETAIL_ID"), -1))
      Call SetFieldValue("BILLING_DOC_ID", NVLI(Rs("BILLING_DOC_ID"), -1))
      Call SetFieldValue("BILLING_PAYMENT_ID", NVLI(Rs("BILLING_PAYMENT_ID"), -1))
      Call SetFieldValue("GL_ID", NVLI(Rs("GL_ID"), -1))
      Call SetFieldValue("GL_DESC", NVLS(Rs("GL_DESC"), ""))
      Call SetFieldValue("GL_TYPE", NVLI(Rs("GL_TYPE"), -1))
      Call SetFieldValue("GL_AMOUNT", NVLD(Rs("GL_AMOUNT"), 0))
      
      Call SetFieldValue("GL_NO", NVLS(Rs("GL_NO"), ""))
      Call SetFieldValue("GL_NAME", NVLS(Rs("GL_NAME"), ""))
      Call SetFieldValue("SUM_FLAG", NVLS(Rs("SUM_FLAG"), ""))
      Call SetFieldValue("SHOW_FLAG", NVLS(Rs("SHOW_FLAG"), ""))

   End If
End Sub

Public Sub RegisterFields()
   TableName = "GL_DETAIL"
   SequenceName = "GL_DETAIL_SEQ"
   PKField = "GL_DETAIL_ID"
   
   Call AddField("GL_DETAIL_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BILLING_DOC_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("BILLING_PAYMENT_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("GL_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("GL_DESC", STRING_TYPE, -1, DATA_CAT)
   Call AddField("GL_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("GL_AMOUNT", MONEY_TYPE, -1, DATA_CAT)
   
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
   
   Call AddField("GL_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("GL_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("SUM_FLAG", STRING_TYPE, "", TEMP_CAT)
   
    Call AddField("SHOW_FLAG", MONEY_TYPE, -1, TEMP_CAT)
   
End Sub

