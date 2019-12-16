VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostItemRaw"
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

Public DoItems As Collection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   Set DoItems = New Collection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set DoItems = Nothing
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
Dim BANK_FLAG As String
Dim POST_FLAG As String

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

Dim CSTITEM_RAW_ID As Long
Dim COST_ITEM_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim PARAM_PROCESS_ID As Long

   CSTITEM_RAW_ID = Val(GetFieldValue("CSTITEM_RAW_ID"))
   COST_ITEM_ID = Val(GetFieldValue("COST_ITEM_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   PARAM_PROCESS_ID = Val(GetFieldValue("PARAM_PROCESS_ID"))
   
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT  CP.PART_ITEM_ID FEED_ID,PI.PART_NO,PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN COST_ITEM CI ON (NM.COST_ITEM_ID = CI.COST_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN COSTPRD_ITEM CP ON (CP.COSTPRD_ITEM_ID = CI.COSTPRD_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN COST_PRODUCTION CD ON (CD.COST_PRODUCTION_ID = CP.COST_PRODUCTION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CP.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      TempStr = "WHERE "
      WhereStr = WhereStr & TempStr & " (NM.ITEM_COST > 0 ) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (NM.ITEM_AMOUNT > 0 ) " & vbCrLf
      TempStr = "AND "
      WhereStr = WhereStr & TempStr & " (CP.PART_ITEM_ID IS NOT NULL ) " & vbCrLf
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT  NM.PART_ITEM_ID,PI.PART_NO,PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN COST_ITEM CI ON (NM.COST_ITEM_ID = CI.COST_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN COSTPRD_ITEM CP ON (CP.COSTPRD_ITEM_ID = CI.COSTPRD_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN COST_PRODUCTION CD ON (CD.COST_PRODUCTION_ID = CP.COST_PRODUCTION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (NM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
'      TempStr = "WHERE "
'      WhereStr = WhereStr & TempStr & " (NM.ITEM_COST > 0 ) " & vbCrLf
'      TempStr = "AND "
'      WhereStr = WhereStr & TempStr & " (NM.ITEM_AMOUNT > 0 ) " & vbCrLf
      
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT SUM(NM.ITEM_COST) ITEM_COST,SUM(NM.ITEM_AMOUNT) ITEM_AMOUNT,NM.PART_ITEM_ID,CP.PART_ITEM_ID FEED_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN COST_ITEM CI ON (NM.COST_ITEM_ID = CI.COST_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN COSTPRD_ITEM CP ON (CP.COSTPRD_ITEM_ID = CI.COSTPRD_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN COST_PRODUCTION CD ON (CD.COST_PRODUCTION_ID = CP.COST_PRODUCTION_ID) "
      
      WhereStr = ""
   End If
   
   If CSTITEM_RAW_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CSTITEM_RAW_ID = " & CSTITEM_RAW_ID & ") " & vbCrLf
   End If
   
   If COST_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.COST_ITEM_ID = " & COST_ITEM_ID & ") " & vbCrLf
   End If
   
   If PARAM_PROCESS_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (CI.PARAM_PROCESS_ID = " & PARAM_PROCESS_ID & ") " & vbCrLf
   End If
   
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (CD.DOCUMENT_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (CD.DOCUMENT_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   GroupByStr = ""
   If Ind = 4 Then
      GroupByStr = "GROUP BY NM.PART_ITEM_ID,CP.PART_ITEM_ID "
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY NM.CSTITEM_RAW_ID "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY NM.PART_ITEM_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   If Ind = 4 Then
      OrderByStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CSTITEM_RAW_ID", NVLI(Rs("CSTITEM_RAW_ID"), -1))
      Call SetFieldValue("COST_ITEM_ID", NVLI(Rs("COST_ITEM_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
   ElseIf Ind = 2 Then
      Call SetFieldValue("FEED_ID", NVLI(Rs("FEED_ID"), -1))
      Call SetFieldValue("PART_NO", NVLS(Rs("PART_NO"), ""))
      Call SetFieldValue("PART_DESC", NVLS(Rs("PART_DESC"), ""))
   ElseIf Ind = 3 Then
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("PART_NO", NVLS(Rs("PART_NO"), ""))
      Call SetFieldValue("PART_DESC", NVLS(Rs("PART_DESC"), ""))
   ElseIf Ind = 4 Then
      Call SetFieldValue("FEED_ID", NVLI(Rs("FEED_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("ITEM_COST", NVLD(Rs("ITEM_COST"), 0))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "CSTITEM_RAW"
   SequenceName = "CSTITEM_RAW_SEQ"
   PKField = "CSTITEM_RAW_ID"
   
   Call AddField("CSTITEM_RAW_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("COST_ITEM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PART_ITEM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("ITEM_COST", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   
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
   
   Call AddField("FEED_ID", INT_TYPE, -1, TEMP_CAT)
   Call AddField("PART_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PART_DESC", STRING_TYPE, "", TEMP_CAT)
   
   
   Call AddField("PARAM_PROCESS_ID", INT_TYPE, -1, TEMP_CAT)
End Sub


