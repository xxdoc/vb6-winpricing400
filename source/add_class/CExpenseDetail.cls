VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpenseDetail"
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
Dim TempRs As ADODB.Recordset
Dim Ct As CCashTran
Dim iCount As Long
Dim Cq As CCheque

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

Dim EXPENSE_DETAIL_ID As Long
Dim EXPENSE_DETAIL_TYPE As Long
Dim EXPENSE_ID As Long

Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim TempDate As String

   EXPENSE_DETAIL_ID = Val(GetFieldValue("EXPENSE_DETAIL_ID"))
   EXPENSE_DETAIL_TYPE = Val(GetFieldValue("EXPENSE_DETAIL_TYPE"))
   EXPENSE_ID = Val(GetFieldValue("EXPENSE_ID"))
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.*,PP.PARAMETER_PROCESS_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETER_PROCESS PP ON (PP.PARAMETER_PROCESS_ID= NM.EXPENSE_DETAIL_TYPE ) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE EP ON (EP.EXPENSE_ID= NM.EXPENSE_ID ) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SUM(NM.EXPENSE_DETAIL_PRICE) EXPENSE_DETAIL_PRICE,NM.EXPENSE_DETAIL_TYPE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE EP ON (EP.EXPENSE_ID= NM.EXPENSE_ID ) "
      
      WhereStr = ""
   End If
   
   If EXPENSE_DETAIL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EXPENSE_DETAIL_ID = " & EXPENSE_DETAIL_ID & ") " & vbCrLf
   End If
  
  If EXPENSE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EXPENSE_ID = " & EXPENSE_ID & ") " & vbCrLf
   End If
   
   If EXPENSE_DETAIL_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.EXPENSE_DETAIL_TYPE = " & EXPENSE_DETAIL_TYPE & ") " & vbCrLf
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EP.EXPENSE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EP.EXPENSE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY NM.EXPENSE_DETAIL_TYPE "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY NM.EXPENSE_DETAIL_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   If Ind = 2 Then
      OrderByStr = ""
   End If
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("EXPENSE_DETAIL_ID", NVLI(Rs("EXPENSE_DETAIL_ID"), -1))
      Call SetFieldValue("EXPENSE_DETAIL_TYPE", NVLI(Rs("EXPENSE_DETAIL_TYPE"), -1))
      Call SetFieldValue("EXPENSE_ID", NVLI(Rs("EXPENSE_DETAIL_TYPE"), -1))
      Call SetFieldValue("EXPENSE_DETAIL_DESC", NVLS(Rs("EXPENSE_DETAIL_DESC"), ""))
      Call SetFieldValue("EXPENSE_DETAIL_AMOUNT", NVLD(Rs("EXPENSE_DETAIL_AMOUNT"), 0))
      Call SetFieldValue("EXPENSE_DETAIL_AVG", NVLD(Rs("EXPENSE_DETAIL_AVG"), 0))
      Call SetFieldValue("EXPENSE_DETAIL_PRICE", NVLD(Rs("EXPENSE_DETAIL_PRICE"), 0))
      
      Call SetFieldValue("PARAMETER_PROCESS_NAME", NVLS(Rs("PARAMETER_PROCESS_NAME"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("EXPENSE_DETAIL_TYPE", NVLI(Rs("EXPENSE_DETAIL_TYPE"), -1))
      Call SetFieldValue("EXPENSE_DETAIL_PRICE", NVLD(Rs("EXPENSE_DETAIL_PRICE"), 0))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "EXPENSE_DETAIL"
   SequenceName = "EXPENSE_DETAIL_SEQ"
   PKField = "EXPENSE_DETAIL_ID"
   
   Call AddField("EXPENSE_DETAIL_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("EXPENSE_DETAIL_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("EXPENSE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("EXPENSE_DETAIL_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("EXPENSE_DETAIL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("EXPENSE_DETAIL_AVG", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("EXPENSE_DETAIL_PRICE", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   
   Call AddField("PARAMETER_PROCESS_NAME", STRING_TYPE, "", TEMP_CAT)
   
End Sub
