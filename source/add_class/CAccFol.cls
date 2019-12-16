VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccFol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String

Public ACC_FOL_ID As Long
Public CUSTOMER_ID As Long
Public FOL_NOTE As String
Public FOL_DATE As Date
Public CANCEL_FLAG As String

Public ORDER_BY As Long
Public ORDER_TYPE As Long

Private m_AccountFollowCancelFlag_N As Collection

Private Sub Class_Initialize()
  Set m_Conn = glbDatabaseMngr.DBConnection
  Set m_AccountFollowCancelFlag_N = New Collection
End Sub

'Private Sub AddField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
'Dim F As CTableField
'
'   Set F = New CTableField
'   Call F.InitField(FName, FType, FValue, FCat)
'   Call m_FieldList.add(F, FName)
'   Set F = Nothing
'End Sub
'
'Public Function GetFieldValue(FieldName As String)
'Dim Tf As CTableField
'
'   Set Tf = m_FieldList(FieldName)
'   GetFieldValue = Tf.GetValue
'End Function

'Public Sub SetFieldValue(FieldName As String, Value As Variant)
'Dim Tf As CTableField
'
'   Set Tf = m_FieldList(FieldName)
'   Tf.SetValue (Value)
'End Sub

'Public Sub DeleteData()
'Dim SQL As String
'Dim Tf As CTableField
'
'   Set Tf = m_FieldList(PKField)
'
'   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
'   Call m_ADOConn.Execute(SQL)
'End Sub
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)

Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String

'   ACC_FOL_ID = Val(GetFieldValue("ACC_FOL_ID"))
'   CUSTOMER_ID = Val(GetFieldValue("CUSTOMER_ID"))
'   FOL_NOTE = Val(GetFieldValue("FOL_NOTE"))
'    FOL_DATE = Val(GetFieldValue("FOL_DATE"))
'    CANCEL_FLAG = Val(GetFieldValue("CANCEL_FLAG"))
'
'   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
'   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   

      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",CS.*"
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM ACC_FOL NM "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CS ON (NM.CUSTOMER_ID  =  CS.CUSTOMER_ID) "
      
      WhereStr = ""
      
  
   

If ACC_FOL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (NM.ACC_FOL_ID= " & ACC_FOL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (NM.ACC_FOL_ID = " & ACC_FOL_ID & ") "
      End If
   End If

   
     If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (NM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (NM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   
  OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ACC_FOL_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY FOL_DATE "
   Else
      OrderByStr = "ORDER BY CS.CUSTOMER_ID "
   End If
   



 
 
 
  If OrderType = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function
Public Function QueryData2(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String


   If Ind = 1 Then
      SelectStr1 = "SELECT AF.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM ACC_FOL AF "
      
      WhereStr = ""
   End If

If ACC_FOL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AF.ACC_FOL_ID= " & ACC_FOL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AF.ACC_FOL_ID = " & ACC_FOL_ID & ") "
      End If
   End If

   
     If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AF.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AF.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   
     If Len(CANCEL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AF.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (AF.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      End If
   End If
   
  
   
  OrderByStr = ""
   OrderTypeStr = ""
 If Ind = 1 Then
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY AF.ACC_FOL_ID,AF.FOL_DATE "
   Else
      OrderByStr = "ORDER BY AF.FOL_DATE"
   End If
   End If

 
  If OrderType = 1 Then
          OrderByStr = OrderByStr & " DESC "
   Else
           OrderByStr = OrderByStr & " ASC "
   End If
   
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function


Public Sub PopulateFromRS(Rs As ADODB.Recordset)
  ' If Ind = 1 Then
'      Call SetFieldValue("ACC_FOL_ID", NVLI(Rs("ACC_FOL_ID"), -1))
'      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
'      Call SetFieldValue("FOL_NOTE", NVLS(Rs("FOL_NOTE"), ""))
'      Call SetFieldValue("FOL_DATE", InternalDateToDate(NVLS(Rs("FOL_DATE"), "")))
'      Call SetFieldValue("CANCEL_FLAG", NVLS(Rs("CANCEL_FLAG"), ""))
   ACC_FOL_ID = NVLI(Rs("ACC_FOL_ID"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   
   FOL_NOTE = NVLS(Rs("FOL_NOTE"), "")
   FOL_DATE = InternalDateToDate(NVLS(Rs("FOL_DATE"), ""))
   CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
 '  End If
End Sub
Public Sub PopulateFromRS2(Ind As Long, Rs As ADODB.Recordset)
 'If Ind = 1 Then
'      Call SetFieldValue("ACC_FOL_ID", NVLI(Rs("ACC_FOL_ID"), -1))
'      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
'      Call SetFieldValue("FOL_NOTE", NVLS(Rs("FOL_NOTE"), ""))
'      Call SetFieldValue("FOL_DATE", InternalDateToDate(NVLS(Rs("FOL_DATE"), "")))
'      Call SetFieldValue("CANCEL_FLAG", NVLS(Rs("CANCEL_FLAG"), ""))
   ACC_FOL_ID = NVLI(Rs("ACC_FOL_ID"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   
   FOL_NOTE = NVLS(Rs("FOL_NOTE"), "")
   FOL_DATE = InternalDateToDate(NVLS(Rs("FOL_DATE"), ""))
   CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
 '  End If
End Sub
'Public Sub RegisterFields()
'   TableName = "ACC_FOL"
'   SequenceName = "ACC_FOL_SEQ"
'   PKField = "ACC_FOL_ID"
'
'   Call AddField("ACC_FOL_ID", INT_TYPE, -1, ID_CAT)
'   Call AddField("CUSTOMER_ID", INT_TYPE, -1, DATA_CAT)
'
'
'   Call AddField("FOL_NOTE", STRING_TYPE, "", DATA_CAT)
'   Call AddField("FOL_DATE", DATE_TYPE, -1, DATA_CAT)
'   Call AddField("CANCEL_FLAG", STRING_TYPE, "", DATA_CAT)
'
'   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
'   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
'   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
'   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
'
'
'   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
'   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
'End Sub
'
'Public Sub AddEditData()
'Dim SQL As String
'Dim Tf As CTableField
'Dim NET_AMOUNT As Double
'Dim FEE_AMOUNT As Double
'Dim AMOUNT As Double
'
'   Set Tf = m_FieldList(PKField)
'
'   Call PopulateInternalField(ShowMode, Me)
'
'   If ShowMode = SHOW_ADD Then
'      SQL = GenerateInsertSQL(Me)
'   ElseIf ShowMode = SHOW_EDIT Then
'      SQL = GenerateUpdateSQL(Me)
'   End If
'   '''Debug.Print (SQL)
'   Call m_ADOConn.Execute(SQL)
'End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("ACC_FOL_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO ACC_FOL " & _
                     "( " & _
                     "   ACC_FOL_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   FOL_NOTE, " & _
                     "   FOL_DATE, " & _
                     "   CANCEL_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & CUSTOMER_ID & "," & _
                         "'" & FOL_NOTE & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(FOL_DATE))) & "'," & _
                         "'" & CANCEL_FLAG & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ACC_FOL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ACC_FOL SET " & _
                     "   CUSTOMER_ID = " & CUSTOMER_ID & ", " & _
                     "   FOL_NOTE = '" & FOL_NOTE & "', " & _
                    "    FOL_DATE = '" & ChangeQuote(Trim(DateToStringInt(FOL_DATE))) & "', " & _
                     "   CANCEL_FLAG = '" & CANCEL_FLAG & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                   " WHERE ACC_FOL_ID = " & ACC_FOL_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ACC_FOL WHERE ACC_FOL_ID = " & ACC_FOL_ID
      m_Conn.Execute (SQL1)
End Function

