VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMKTFol"
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

Public MKT_FOL_ID As Long
Public CUSTOMER_ID As Long
Public FOL_NOTE As String
Public FOL_DATE As Date
Public CANCEL_FLAG As String


Public ORDER_BY As Long
Public ORDER_TYPE As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public DOCUMENT_TYPE As Long
Private m_MarketingFollowCancelFlag_N As Collection

Private Sub Class_Initialize()
  Set m_Conn = glbDatabaseMngr.DBConnection
  Set m_MarketingFollowCancelFlag_N = New Collection
  FROM_DATE = -1
   TO_DATE = -1
End Sub


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
   
'   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",CS.*"
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MKT_FOL NM "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CS ON (NM.CUSTOMER_ID  =  CS.CUSTOMER_ID) "
      
      WhereStr = ""
  
  ' End If
   

If MKT_FOL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (NM.MKT_FOL_ID= " & MKT_FOL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (NM.MKT_FOL_ID = " & MKT_FOL_ID & ") "
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
      OrderByStr = "ORDER BY MKT_FOL_ID "
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
      SelectStr1 = "SELECT MKT.* "
    SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM MKT_FOL MKT "
      
      WhereStr = ""
   End If
      

If MKT_FOL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MKT.MKT_FOL_ID= " & MKT_FOL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MKT.MKT_FOL_ID = " & MKT_FOL_ID & ") "
      End If
   End If

 If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MKT.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MKT.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   

    If Len(CANCEL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MKT.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (MKT.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
      End If
   End If
   
  OrderByStr = ""
   OrderTypeStr = ""
    
    If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY MKT.MKT_FOL_ID,MKT.FOL_DATE "
      Else
        OrderByStr = "ORDER BY  MKT.FOL_DATE  "
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
 
   MKT_FOL_ID = NVLI(Rs("MKT_FOL_ID"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   
   FOL_NOTE = NVLS(Rs("FOL_NOTE"), "")
   FOL_DATE = InternalDateToDate(NVLS(Rs("FOL_DATE"), ""))
   CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
 '  End If
End Sub
Public Sub PopulateFromRS2(Ind As Long, Rs As ADODB.Recordset)
 
   MKT_FOL_ID = NVLI(Rs("MKT_FOL_ID"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   
   FOL_NOTE = NVLS(Rs("FOL_NOTE"), "")
   FOL_DATE = InternalDateToDate(NVLS(Rs("FOL_DATE"), ""))
   CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
 '  End If
End Sub




Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MKT_FOL_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO MKT_FOL " & _
                     "( " & _
                     "   MKT_FOL_ID, " & _
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
         MKT_FOL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MKT_FOL SET " & _
                     "   CUSTOMER_ID = " & CUSTOMER_ID & ", " & _
                     "   FOL_NOTE = '" & FOL_NOTE & "', " & _
                    "    FOL_DATE = '" & ChangeQuote(Trim(DateToStringInt(FOL_DATE))) & "', " & _
                     "   CANCEL_FLAG = '" & CANCEL_FLAG & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                   " WHERE MKT_FOL_ID = " & MKT_FOL_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM MKT_FOL WHERE MKT_FOL_ID = " & MKT_FOL_ID
      m_Conn.Execute (SQL1)
End Function

