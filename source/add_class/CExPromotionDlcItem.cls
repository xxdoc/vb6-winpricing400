VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExPromotionDlcItem"
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

Public EX_PROMOTION_DLC_ITEM_ID As Long
Public EX_WORKS_PRICE_ID As Long
Public CUSTOMER_ID As Long
Public RATE_TYPE_CUS As Long
Public WEIGHT_PER_PACK_CUS As Long
Public DELIVERY_CUS_ITEM_ID As Long
Public DISCOUNT_AMOUNT As Double

Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public DELIVERY_CUS_ITEM_CODE As String
Public DELIVERY_CUS_ITEM_NAME As String


Public EX_WORKS_PRICE_CODE As String
Public EX_WORKS_PRICE_DESC As String
Public EX_WORKS_PRICE_LEVEL As Long
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
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

   If Ind = 1 Then
      SelectStr1 = "SELECT EWP.*, EPDI.*, DCI.DELIVERY_CUS_ITEM_CODE, DCI.DELIVERY_CUS_ITEM_NAME, CM.CUSTOMER_ID, CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EX_PROMOTION_DLC_ITEM EPDI "
      FromStr = FromStr & "LEFT OUTER JOIN EX_WORKS_PRICE EWP ON (EPDI.EX_WORKS_PRICE_ID = EWP.EX_WORKS_PRICE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DELIVERY_CUS_ITEM DCI ON (EPDI.DELIVERY_CUS_ITEM_ID = DCI.DELIVERY_CUS_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (EPDI.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "

      WhereStr = ""
   End If
   
   If EX_PROMOTION_DLC_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EPDI.EX_PROMOTION_DLC_ITEM_ID = " & EX_PROMOTION_DLC_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EPDI.EX_PROMOTION_DLC_ITEM_ID = " & EX_PROMOTION_DLC_ITEM_ID & ") "
      End If
   End If
   
   If EX_WORKS_PRICE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EPDI.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EPDI.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If Ind = 1 Then
      If OrderBy = 1 Then
        OrderByStr = "ORDER BY EPDI.EX_PROMOTION_DLC_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY  DCI.DELIVERY_CUS_ITEM_CODE  "
      Else
         OrderByStr = "ORDER BY EPDI.EX_PROMOTION_DLC_ITEM_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EX_PROMOTION_DLC_ITEM_ID = NVLI(Rs("EX_PROMOTION_DLC_ITEM_ID"), -1)
      EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      DELIVERY_CUS_ITEM_CODE = NVLS(Rs("DELIVERY_CUS_ITEM_CODE"), "")
      DELIVERY_CUS_ITEM_NAME = NVLS(Rs("DELIVERY_CUS_ITEM_NAME"), "")
      RATE_TYPE_CUS = NVLI(Rs("RATE_TYPE_CUS"), -1)
      WEIGHT_PER_PACK_CUS = NVLI(Rs("WEIGHT_PER_PACK_CUS"), -1)
      DELIVERY_CUS_ITEM_ID = NVLI(Rs("DELIVERY_CUS_ITEM_ID"), -1)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
     
      EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
      EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
      EX_WORKS_PRICE_LEVEL = NVLI(Rs("EX_WORKS_PRICE_LEVEL"), -1)
   End If
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
   
   TempID1 = EX_WORKS_PRICE_ID
   If EX_WORKS_PRICE_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = DELIVERY_CUS_ITEM_ID
   If DELIVERY_CUS_ITEM_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EX_PROMOTION_DLC_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EX_PROMOTION_DLC_ITEM " & _
                     "( " & _
                     "   EX_PROMOTION_DLC_ITEM_ID, " & _
                     "   EX_WORKS_PRICE_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   RATE_TYPE_CUS, " & _
                     "   WEIGHT_PER_PACK_CUS, " & _
                     "   DELIVERY_CUS_ITEM_ID, " & _
                     "   DISCOUNT_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "" & TempID1 & ", " & _
                        TempID2 & ", " & _
                        RATE_TYPE_CUS & ", " & _
                        WEIGHT_PER_PACK_CUS & ", " & _
                        TempID3 & ", " & _
                        DISCOUNT_AMOUNT & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               EX_PROMOTION_DLC_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EX_PROMOTION_DLC_ITEM SET " & _
                     "   EX_WORKS_PRICE_ID = " & TempID1 & ", " & _
                     "   CUSTOMER_ID = " & TempID2 & ", " & _
                     "   RATE_TYPE_CUS = " & RATE_TYPE_CUS & ", " & _
                     "   WEIGHT_PER_PACK_CUS = " & WEIGHT_PER_PACK_CUS & ", " & _
                     "   DELIVERY_CUS_ITEM_ID = " & TempID3 & ", " & _
                     "   DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EX_PROMOTION_DLC_ITEM_ID = " & EX_PROMOTION_DLC_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EX_PROMOTION_DLC_ITEM  WHERE EX_PROMOTION_DLC_ITEM_ID = " & EX_PROMOTION_DLC_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
 Set m_Conn = Nothing
End Sub



