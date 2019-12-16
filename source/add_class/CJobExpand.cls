VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobExpand"
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
   Call SetFieldValue("ENTERPRISE_ID", glbUser.ENTERPRISE_ID)
   
   BANK_FLAG = GetFieldValue("BANK_FLAG")
   If BANK_FLAG = "" Then
      BANK_FLAG = "N"
      Call SetFieldValue("BANK_FLAG", BANK_FLAG)
   End If
   
   POST_FLAG = GetFieldValue("POST_FLAG")
   If POST_FLAG = "" Then
      POST_FLAG = "N"
      Call SetFieldValue("POST_FLAG", POST_FLAG)
   End If
   
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

Public Sub DeleteByProcesstypeArea()
Dim SQL As String
Dim Tf As CTableField
Dim Area As Long
Dim PROCESS_TYPE As Long

   Set Tf = m_FieldList(PKField)
   PROCESS_TYPE = GetFieldValue("PROCESS_TYPE")
   Area = GetFieldValue("AREA")
   
   SQL = "DELETE FROM " & TableName & " WHERE (AREA = " & Area & ") AND (PROCESS_TYPE = " & PROCESS_TYPE & ") "
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

Dim JOB_EXPAND_ID As Long
Dim Area As Long
Dim PROCESS_TYPE As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   JOB_EXPAND_ID = Val(GetFieldValue("JOB_EXPAND_ID"))
   Area = Val(GetFieldValue("AREA"))
   PROCESS_TYPE = Val(GetFieldValue("PROCESS_TYPE"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (NM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If JOB_EXPAND_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.JOB_EXPAND_ID = " & JOB_EXPAND_ID & ") " & vbCrLf
   End If
   
   If Area > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.AREA = " & Area & ") " & vbCrLf
   End If
      
   If PROCESS_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PROCESS_TYPE = " & PROCESS_TYPE & ") " & vbCrLf
   End If
      
   If Ind = 1 Then
      OrderByStr = "ORDER BY PI.PART_NO "
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
      Call SetFieldValue("JOB_EXPAND_ID", NVLI(Rs("JOB_EXPAND_ID"), -1))
      Call SetFieldValue("AREA", NVLI(Rs("AREA"), -1))
      Call SetFieldValue("PROCESS_TYPE", NVLI(Rs("PROCESS_TYPE"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("JOB_FROM_DATE", InternalDateToDate(NVLS(Rs("JOB_FROM_DATE"), "")))
      Call SetFieldValue("JOB_TO_DATE", InternalDateToDate(NVLS(Rs("JOB_TO_DATE"), "")))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "JOB_EXPAND"
   SequenceName = "JOB_EXPAND_SEQ"
   PKField = "JOB_EXPAND_ID"
   
   Call AddField("JOB_EXPAND_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("AREA", INT_TYPE, -1, DATA_CAT)
   Call AddField("PROCESS_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("PART_ITEM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("JOB_FROM_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("JOB_TO_DATE", DATE_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PART_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PART_DESC", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub

