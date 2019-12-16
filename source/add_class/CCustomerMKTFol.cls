VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomerMKTFol"
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

Dim MKT_FOL_ID As Long
Dim CUSTOMER_ID As Long
Dim FOL_NOTE As String
Dim FOL_DATE As Date
Dim CANCEL_FLAG As String

Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   MKT_FOL_ID = Val(GetFieldValue("MKT_FOL_ID"))
   CUSTOMER_ID = Val(GetFieldValue("CUSTOMER_ID"))
   FOL_NOTE = Val(GetFieldValue("FOL_NOTE"))
    FOL_DATE = Val(GetFieldValue("FOL_DATE"))
    CANCEL_FLAG = Val(GetFieldValue("CANCEL_FLAG"))

   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",CS.*"
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CS ON (NM.CUSTOMER_ID  =  CS.CUSTOMER_ID) "
      
      WhereStr = ""
  
   End If
   
   If MKT_FOL_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.MKT_FOL_ID = " & MKT_FOL_ID & ") " & vbCrLf
   End If

   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CUSTOMER_ID = " & CUSTOMER_ID & ") " & vbCrLf
   End If
   
   
   GroupByStr = ""
   If Ind = 3 Then
   '   GroupByStr = "GROUP BY NM.BANK_ACCOUNT, NM.TX_TYPE "
   End If
   
   If ORDER_BY = 1 Then
      OrderByStr = "ORDER BY MKT_FOL__ID "
   Else
      OrderByStr = "ORDER BY CUSTOMER_ID "
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
      Call SetFieldValue("MKT_FOL_ID", NVLI(Rs("MKT_FOL_ID"), -1))
      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
      Call SetFieldValue("FOL_NOTE", NVLS(Rs("FOL_NOTE"), ""))
      Call SetFieldValue("FOL_DATE ", InternalDateToDate(NVLS(Rs("FOL_DATE "), "")))
      Call SetFieldValue("CANCEL_FLAG", NVLS(Rs("CANCEL_FLAG"), ""))


   End If
End Sub
Public Sub RegisterFields()
   TableName = "MKT_FOL"
   SequenceName = "MKT_FOL_SEQ"
   PKField = "MKT_FOL_ID"
   
   Call AddField("MKT_FOL_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("CUSTOMER_ID", INT_TYPE, -1, DATA_CAT)
   
   
   
   Call AddField("FOL_NOTE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("FOL_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("CANCEL_FLAG", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
 
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
End Sub



