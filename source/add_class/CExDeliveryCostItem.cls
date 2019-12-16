VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExDeliveryCostItem"
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
' End of generic data member

Public EX_DELIVERY_COST_ITEM_ID As Long
Public EX_DELIVERY_COST_ITEM_CODE As String
Public EX_DELIVERY_COST_ITEM_NAME As String
Public EX_WORKS_PRICE_ID As Long
Public RATE_TYPE As Long
Public RATE_TYPE_CUS As Long
Public RATE_DELIVERY As Double
Public RATE_CUSTOMER As Double
Public WEIGHT_PER_PACK As Long
Public WEIGHT_PER_PACK_CUS  As Long
Public CUSTOMER_ID As Long
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public DELIVERY_CUS_ITEM_ID As Long
Public DELIVERY_CUS_ITEM_CODE As String
Public DELIVERY_CUS_ITEM_NAME As String
Public SocPartType As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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
 
   If Ind = 1 Then
      SelectStr1 = "SELECT EDCI.*,DCI.DELIVERY_CUS_ITEM_CODE,DCI.DELIVERY_CUS_ITEM_NAME,CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME " 'EDCI.RATE_TYPE,EDCI.RATE_TYPE_CUS,
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EX_DELIVERY_COST_ITEM EDCI "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (EDCI.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DELIVERY_CUS_ITEM DCI ON (EDCI.DELIVERY_CUS_ITEM_ID = DCI.DELIVERY_CUS_ITEM_ID) "

      WhereStr = ""
   End If
   
   If EX_DELIVERY_COST_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EDCI.EX_DELIVERY_COST_ITEM_ID = " & EX_DELIVERY_COST_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EDCI.EX_DELIVERY_COST_ITEM_ID = " & EX_DELIVERY_COST_ITEM_ID & ") "
      End If
   End If
   
   If DELIVERY_CUS_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EDCI.DELIVERY_CUS_ITEM_ID = " & DELIVERY_CUS_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EDCI.DELIVERY_CUS_ITEM_ID = " & DELIVERY_CUS_ITEM_ID & ") "
      End If
   End If
   
   If EX_WORKS_PRICE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EDCI.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EDCI.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      End If
   End If
   
 
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EDCI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EDCI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If Len(CUSTOMER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_CODE = '" & CUSTOMER_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_CODE = '" & CUSTOMER_CODE & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY  EDCI.EX_DELIVERY_COST_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY  CM.CUSTOMER_CODE ASC,DCI.DELIVERY_CUS_ITEM_CODE ASC,RATE_TYPE ASC,RATE_TYPE_CUS "
      Else
         OrderByStr = "ORDER BY EDCI.EX_DELIVERY_COST_ITEM_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
      EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      DELIVERY_CUS_ITEM_ID = NVLI(Rs("DELIVERY_CUS_ITEM_ID"), -1)
      DELIVERY_CUS_ITEM_CODE = NVLS(Rs("DELIVERY_CUS_ITEM_CODE"), "")
      DELIVERY_CUS_ITEM_NAME = NVLS(Rs("DELIVERY_CUS_ITEM_NAME"), "")
      RATE_DELIVERY = NVLD(Rs("RATE_DELIVERY"), 0)
      RATE_CUSTOMER = NVLD(Rs("RATE_CUSTOMER"), 0)
      WEIGHT_PER_PACK = NVLI(Rs("WEIGHT_PER_PACK"), -1)
      WEIGHT_PER_PACK_CUS = NVLI(Rs("WEIGHT_PER_PACK_CUS"), -1)
      RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
      RATE_TYPE_CUS = NVLI(Rs("RATE_TYPE_CUS"), -1)
   End If
End Function
Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
   
   TempID1 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = DELIVERY_CUS_ITEM_ID
   If DELIVERY_CUS_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EX_DELIVERY_COST_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EX_DELIVERY_COST_ITEM " & _
                     "( " & _
                     "   EX_DELIVERY_COST_ITEM_ID, " & _
                     "   EX_WORKS_PRICE_ID, " & _
                     "   RATE_TYPE, " & _
                     "   RATE_TYPE_CUS, " & _
                     "   CUSTOMER_ID, " & _
                     "   DELIVERY_CUS_ITEM_ID, " & _
                     "   RATE_DELIVERY, " & _
                     "   RATE_CUSTOMER, " & _
                     "   WEIGHT_PER_PACK, " & _
                     "   WEIGHT_PER_PACK_CUS, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "" & EX_WORKS_PRICE_ID & ", " & _
                        RATE_TYPE & ", " & _
                        RATE_TYPE_CUS & ", " & _
                        TempID1 & ", " & _
                        TempID2 & ", " & _
                        RATE_DELIVERY & ", " & _
                        RATE_CUSTOMER & ", " & _
                        WEIGHT_PER_PACK & ", " & _
                        WEIGHT_PER_PACK_CUS & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               EX_DELIVERY_COST_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EX_DELIVERY_COST_ITEM SET " & _
                     "   EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ", " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   RATE_TYPE_CUS = " & RATE_TYPE_CUS & ", " & _
                     "   CUSTOMER_ID = " & TempID1 & ", " & _
                     "   DELIVERY_CUS_ITEM_ID = " & TempID2 & ", " & _
                     "   RATE_DELIVERY = " & RATE_DELIVERY & ", " & _
                     "   RATE_CUSTOMER = " & RATE_CUSTOMER & ", " & _
                     "   WEIGHT_PER_PACK = " & WEIGHT_PER_PACK & ", " & _
                     "   WEIGHT_PER_PACK_CUS = " & WEIGHT_PER_PACK_CUS & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EX_DELIVERY_COST_ITEM_ID = " & EX_DELIVERY_COST_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EX_DELIVERY_COST_ITEM WHERE EX_DELIVERY_COST_ITEM_ID = " & EX_DELIVERY_COST_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
 Set m_Conn = Nothing
End Sub



