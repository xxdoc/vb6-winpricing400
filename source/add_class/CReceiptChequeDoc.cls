VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReceiptChequeDoc"
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

Public RECEIPT_CHEQUE_DOC_ID As Long
Public RECEIPT_CHEQUE_DOC_NO As String
Public RECEIPT_CHEQUE_DOC_DATE As Date


Public CHEQUE_DOC_ID As Long
Public BILLING_DOC_ID As Long

Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public RECEIPT_ITEM_AMOUNT As Double
Public CASH_DISCOUNT As Double
Public PAID_AMOUNT As Double


Private m_ChequeDoc_coll As Collection
Public ReceiptCheqDocItems  As Collection

Private Sub Class_Initialize()
  Set m_Conn = glbDatabaseMngr.DBConnection

  Set m_ChequeDoc_coll = New Collection
  Set ReceiptCheqDocItems = New Collection
End Sub

'Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
'
'Dim SQL1 As String
'Dim SQL2 As String
'Dim SelectStr1 As String
'Dim SelectStr2 As String
'Dim FromStr As String
'Dim WhereStr As String
'Dim OrderByStr As String
'Dim OrderTypeStr As String
'Dim TempDate As String
'Dim SubLen As Long
'Dim NewStr As String
'
'
'      SelectStr1 = "SELECT NM.* "
'      SelectStr1 = SelectStr1 & ",CS.*"
'
'      SelectStr2 = "SELECT COUNT(*) "
'
'      FromStr = "FROM ACC_FOL NM "
'      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CS ON (NM.CUSTOMER_ID  =  CS.CUSTOMER_ID) "
'
'      WhereStr = ""
'
'
'
'
'If ACC_FOL_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (NM.ACC_FOL_ID= " & ACC_FOL_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (NM.ACC_FOL_ID = " & ACC_FOL_ID & ") "
'      End If
'   End If
'
'
'     If CUSTOMER_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (NM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (NM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
'      End If
'   End If
'
'
'  OrderByStr = ""
'   OrderTypeStr = ""
'
'   If OrderBy = 1 Then
'      OrderByStr = "ORDER BY ACC_FOL_ID "
'   Else
'      OrderByStr = "ORDER BY CS.CUSTOMER_ID "
'   End If
'
'
'
'
'
'
'
'  If OrderType = 1 Then
'      OrderByStr = OrderByStr & " ASC "
'   Else
'      OrderByStr = OrderByStr & " DESC "
'   End If
'
'
'   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
'   SQL2 = SelectStr2 & FromStr & WhereStr
'
'   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'End Function
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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
      SelectStr1 = "SELECT RCD.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM  RECEIPT_CHEQUE_DOC RCD "
      
      WhereStr = ""
      
   ElseIf Ind = 2 Then
    SelectStr1 = "SELECT RCD.*,CD.*  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM  RECEIPT_CHEQUE_DOC RCD "
       FromStr = FromStr & "LEFT OUTER JOIN CHEQUE_DOC CD ON (RCD.RECEIPT_CHEQUE_DOC_ID= CD.RECEIPT_CHEQUE_DOC_ID) "
      WhereStr = ""
   End If

If RECEIPT_CHEQUE_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RCD.RECEIPT_CHEQUE_DOC_ID= " & RECEIPT_CHEQUE_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RCD.RECEIPT_CHEQUE_DOC_ID  = " & RECEIPT_CHEQUE_DOC_ID & ") "
      End If
   End If

  
     If CHEQUE_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RCD.CHEQUE_DOC_ID= " & CHEQUE_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RCD.CHEQUE_DOC_ID= " & CHEQUE_DOC_ID & ") "
      End If
   End If
   
      
'     If Len(CANCEL_FLAG) > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (AF.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
'      Else
'         WhereStr = WhereStr & "AND (AF.CANCEL_FLAG = '" & CANCEL_FLAG & "') "
'      End If
'   End If
   
  
   
  OrderByStr = ""
   OrderTypeStr = ""
 If Ind = 1 Then
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY  RCD.RECEIPT_CHEQUE_DOC_ID"
   Else
      OrderByStr = "ORDER BY RCD.RECEIPT_CHEQUE_DOC_ID"
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


'Public Sub PopulateFromRS(Rs As ADODB.Recordset)
'  ' If Ind = 1 Then
''      Call SetFieldValue("ACC_FOL_ID", NVLI(Rs("ACC_FOL_ID"), -1))
''      Call SetFieldValue("CUSTOMER_ID", NVLI(Rs("CUSTOMER_ID"), -1))
''      Call SetFieldValue("FOL_NOTE", NVLS(Rs("FOL_NOTE"), ""))
''      Call SetFieldValue("FOL_DATE", InternalDateToDate(NVLS(Rs("FOL_DATE"), "")))
''      Call SetFieldValue("CANCEL_FLAG", NVLS(Rs("CANCEL_FLAG"), ""))
'   ACC_FOL_ID = NVLI(Rs("ACC_FOL_ID"), -1)
'   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
'
'   FOL_NOTE = NVLS(Rs("FOL_NOTE"), "")
'   FOL_DATE = InternalDateToDate(NVLS(Rs("FOL_DATE"), ""))
'   CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
' '  End If
'End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
   RECEIPT_CHEQUE_DOC_ID = NVLI(Rs("RECEIPT_CHEQUE_DOC_ID"), -1)
   RECEIPT_CHEQUE_DOC_NO = NVLS(Rs("RECEIPT_CHEQUE_DOC_NO"), "")
   RECEIPT_CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("RECEIPT_CHEQUE_DOC_DATE"), ""))
   BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
   CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
   
     RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
    
    ElseIf Ind = 2 Then

    End If
   
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double
Dim TempID0 As String
Dim TempID1 As String

TempID0 = CHEQUE_DOC_ID
   If CHEQUE_DOC_ID <= 0 Then
      TempID0 = "NULL"
   End If
   
  TempID1 = BILLING_DOC_ID
   If BILLING_DOC_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("RECEIPT_CHEQUE_DOC_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO  RECEIPT_CHEQUE_DOC" & _
                     "( " & _
                     "   RECEIPT_CHEQUE_DOC_ID, " & _
                     "   RECEIPT_CHEQUE_DOC_NO, " & _
                     "  RECEIPT_CHEQUE_DOC_DATE, " & _
                    "   CHEQUE_DOC_ID, " & _
                    "   BILLING_DOC_ID, " & _
                     "   RECEIPT_ITEM_AMOUNT, " & _
                     "   PAID_AMOUNT, " & _
                     "  CASH_DISCOUNT," & _
                    "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                          "'" & ChangeQuote(Trim(RECEIPT_CHEQUE_DOC_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(RECEIPT_CHEQUE_DOC_DATE))) & "'," & _
                          "" & TempID0 & "," & _
                         "" & TempID1 & "," & _
                          "" & RECEIPT_ITEM_AMOUNT & "," & _
                          "" & PAID_AMOUNT & "," & _
                          "" & CASH_DISCOUNT & "," & _
                          "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         RECEIPT_CHEQUE_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE RECEIPT_CHEQUE_DOC  SET " & _
                     "   RECEIPT_CHEQUE_DOC_NO = '" & ChangeQuote(Trim(RECEIPT_CHEQUE_DOC_NO)) & "', " & _
                     "   RECEIPT_CHEQUE_DOC_DATE = '" & ChangeQuote(Trim(DateToStringInt(RECEIPT_CHEQUE_DOC_DATE))) & "', " & _
                     "   CHEQUE_DOC_ID = " & TempID0 & ", " & _
                      "   BILLING_DOC_ID = " & TempID1 & ", " & _
                     "   RECEIPT_ITEM_AMOUNT = " & RECEIPT_ITEM_AMOUNT & ", " & _
                      "   PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                      "   CASH_DISCOUNT = " & CASH_DISCOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                   " WHERE RECEIPT_CHEQUE_DOC_ID = " & RECEIPT_CHEQUE_DOC_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String


     
      
      SQL1 = "DELETE FROM RECEIPT_CHEQUE_DOC WHERE RECEIPT_CHEQUE_DOC_ID = " & RECEIPT_CHEQUE_DOC_ID
      m_Conn.Execute (SQL1)
End Function



Private Sub Class_Terminate()
 Set ReceiptCheqDocItems = Nothing
  Set m_ChequeDoc_coll = Nothing
End Sub
