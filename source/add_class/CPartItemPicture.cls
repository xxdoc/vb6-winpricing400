VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartItemPicture"
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

Dim PART_ITEM_PICTURE_ID As Long
Dim PART_ITEM_ID As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim PART_ITEM_PICTURE_TYPE As Long
'Dim ENTERPRISE_ID As Long

   PART_ITEM_PICTURE_ID = Val(GetFieldValue("PART_ITEM_PICTURE_ID"))
   PART_ITEM_ID = Val(GetFieldValue("PART_ITEM_ID"))
   PART_ITEM_PICTURE_TYPE = Val(GetFieldValue("PART_ITEM_PICTURE_TYPE"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
'   ENTERPRISE_ID = glbUser.ENTERPRISE_ID
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT NM.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
      
   If PART_ITEM_PICTURE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PART_ITEM_PICTURE_ID = " & PART_ITEM_PICTURE_ID & ") " & vbCrLf
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PART_ITEM_ID = " & PART_ITEM_ID & ") " & vbCrLf
   End If
   
   If PART_ITEM_PICTURE_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.PART_ITEM_PICTURE_TYPE = " & PART_ITEM_PICTURE_TYPE & ") " & vbCrLf
   End If
      
      
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = ""
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_TYPE = 1 Then
         OrderByStr = OrderByStr & ""
      Else
         OrderByStr = OrderByStr & ""
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("PART_ITEM_PICTURE_ID", NVLI(Rs("PART_ITEM_PICTURE_ID"), -1))
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
      Call SetFieldValue("PART_ITEM_PICTURE_TYPE", NVLI(Rs("PART_ITEM_PICTURE_TYPE"), -1))
      Call SetFieldValue("PART_ITEM_PICTURE_PATH", NVLS(Rs("PART_ITEM_PICTURE_PATH"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("PART_ITEM_ID", NVLI(Rs("PART_ITEM_ID"), -1))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "PART_ITEM_PICTURE"
   SequenceName = "PART_ITEM_PICTURE_SEQ"
   PKField = "PART_ITEM_PICTURE_ID"
   
   Call AddField("PART_ITEM_PICTURE_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PART_ITEM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PART_ITEM_PICTURE_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("PART_ITEM_PICTURE_PATH", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
End Sub
