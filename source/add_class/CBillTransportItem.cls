VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillTransportItem"
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
Public BILL_TRANSPORT_ITEM_ID As Long
Public BILL_TRANSPORT_ID As Long
Public BILL_TRANSPORT_CODE As String
Public BILL_TYPE_ID As Long
Public BILL_TYPE_CODE  As String
Public BILL_TYPE_NAME As String
Public PACK_AMOUNT As Double
Public WEIGHT_PER_UNIT  As Long
Public RATE_PER_UNIT As Double
Public TOTAL_PRICE As Double
Public CAL_DRIVER As String
Public CAL_CUSTOMER As String
Public CAL_PRICE_IN_PRODUCT As String
Public CAL_DIRECT As String
Public NOTE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public CREATE_DATE As Date




Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String
Dim TempDate As String

   If Ind = 1 Then
      SelectStr1 = "SELECT BTI.*,MR.KEY_NAME,MR.KEY_CODE "
      SelectStr2 = "SELECT COUNT(*)  "
            
      WhereStr = ""
      FromStr = "FROM BILL_TRANSPORT_ITEM  BTI "
      FromStr = FromStr & "LEFT JOIN MASTER_REF MR ON (BTI.BILL_TYPE_ID=MR.KEY_ID)  "
   End If
   
   If BILL_TRANSPORT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BTI.BILL_TRANSPORT_ITEM_ID = " & BILL_TRANSPORT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BTI.BILL_TRANSPORT_ITEM_ID = " & BILL_TRANSPORT_ITEM_ID & ") "
      End If
   End If
   
   If BILL_TRANSPORT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BTI.BILL_TRANSPORT_ID = " & BILL_TRANSPORT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BTI.BILL_TRANSPORT_ID = " & BILL_TRANSPORT_ID & ") "
      End If
   End If
   
   If BILL_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BTI.BILL_TYPE_ID = " & BILL_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BTI.BILL_TYPE_ID = " & BILL_TYPE_ID & ") "
      End If
   End If
   
   If WEIGHT_PER_UNIT > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BTI.WEIGHT_PER_UNIT = " & WEIGHT_PER_UNIT & ") "
      Else
         WhereStr = WhereStr & "AND (BTI.WEIGHT_PER_UNIT = " & WEIGHT_PER_UNIT & ") "
      End If
   End If
  
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BT.LOT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BT.LOT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BT.LOT_DATE  <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BT.LOT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BTI.BILL_TRANSPORT_ITEM_ID "
      End If
   Else
      OrderByStr = "ORDER BY BTI.BILL_TRANSPORT_ITEM_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
         BILL_TRANSPORT_ITEM_ID = NVLI(Rs("BILL_TRANSPORT_ITEM_ID"), -1)
         BILL_TYPE_ID = NVLI(Rs("BILL_TYPE_ID"), -1)
         PACK_AMOUNT = NVLD(Rs("PACK_AMOUNT"), 0)
         WEIGHT_PER_UNIT = NVLI(Rs("WEIGHT_PER_UNIT"), -1)
         RATE_PER_UNIT = NVLD(Rs("RATE_PER_UNIT"), 0)
         CAL_DRIVER = NVLS(Rs("CAL_DRIVER"), "")
         CAL_CUSTOMER = NVLS(Rs("CAL_CUSTOMER"), "")
         CAL_PRICE_IN_PRODUCT = NVLS(Rs("CAL_PRICE_IN_PRODUCT"), "")
         CAL_DIRECT = NVLS(Rs("CAL_DIRECT"), "")
         TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
         NOTE = NVLS(Rs("NOTE"), "")
         BILL_TYPE_NAME = NVLS(Rs("KEY_NAME"), "")
         BILL_TYPE_CODE = NVLS(Rs("KEY_CODE"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = BILL_TRANSPORT_ID
   If BILL_TRANSPORT_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = BILL_TYPE_ID
   If BILL_TYPE_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILL_TRANSPORT_ITEM_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO BILL_TRANSPORT_ITEM " & _
                     "( " & _
                     "BILL_TRANSPORT_ITEM_ID, " & _
                     "BILL_TRANSPORT_ID, " & _
                     "BILL_TYPE_ID, " & _
                     "PACK_AMOUNT, " & _
                     "WEIGHT_PER_UNIT, " & _
                     "RATE_PER_UNIT, " & _
                     "TOTAL_PRICE, " & _
                     "CAL_DRIVER, " & _
                     "CAL_CUSTOMER, " & _
                     "CAL_PRICE_IN_PRODUCT, " & _
                     "CAL_DIRECT," & _
                     "NOTE, " & _
                     "CREATE_DATE, " & _
                    "CREATE_BY, " & _
                    "MODIFY_DATE, " & _
                    "MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                     ID & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     PACK_AMOUNT & "," & _
                     WEIGHT_PER_UNIT & "," & _
                     RATE_PER_UNIT & "," & _
                     TOTAL_PRICE & "," & _
                     "'" & CAL_DRIVER & "'," & _
                     "'" & CAL_CUSTOMER & "'," & _
                     "'" & CAL_PRICE_IN_PRODUCT & "'," & _
                     "'" & CAL_DIRECT & "'," & _
                     "'" & NOTE & "'," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ") "
                     BILL_TRANSPORT_ITEM_ID = ID
   Else 'Show edit
                     SQL1 = "UPDATE BILL_TRANSPORT_ITEM  SET " & _
                     "  BILL_TRANSPORT_ID   = " & TempID1 & ", " & _
                     "  BILL_TYPE_ID   = " & TempID2 & ", " & _
                     "  PACK_AMOUNT   = " & PACK_AMOUNT & ", " & _
                     "  WEIGHT_PER_UNIT   = " & WEIGHT_PER_UNIT & ", " & _
                     "  RATE_PER_UNIT   = " & RATE_PER_UNIT & ", " & _
                     "  TOTAL_PRICE   = " & TOTAL_PRICE & ", " & _
                     "  CAL_DRIVER = '" & CAL_DRIVER & "', " & _
                     "  CAL_CUSTOMER = '" & CAL_CUSTOMER & "', " & _
                     "  CAL_PRICE_IN_PRODUCT = '" & CAL_PRICE_IN_PRODUCT & "', " & _
                     "  CAL_DIRECT = '" & CAL_DIRECT & "', " & _
                     "  NOTE     = '" & NOTE & "', " & _
                     "  MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "  MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "  WHERE BILL_TRANSPORT_ITEM_ID  = " & BILL_TRANSPORT_ITEM_ID
   End If '
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
   SQL1 = "DELETE FROM BILL_TRANSPORT_ITEM WHERE BILL_TRANSPORT_ITEM_ID = " & BILL_TRANSPORT_ITEM_ID
   m_Conn.Execute (SQL1)
End Function


