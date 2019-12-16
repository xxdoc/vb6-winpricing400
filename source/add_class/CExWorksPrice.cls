VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExWorksPrice"
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

Public EX_WORKS_PRICE_ID As Long
Public EX_WORKS_PRICE_CODE As String
Public EX_WORKS_PRICE_DATE As Date
Public EX_WORKS_PRICE_DESC As String
Public EX_WORKS_PRICE_STATUS As Long
Public EX_WORKS_PRICE_LEVEL As String
Public EX_WORKS_PRICE_TYPE As Long
Public FROM_ACTIVE_DATE As Date
Public TO_VALID_DATE As Date
Public BETWEEN_DATE As Date

Public EX_WORKS_PRICE_ITEM_ID As Long
Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_NO_SEARCH As String
Public CUSTOMER_CODE_SEARCH As String
Public RATE_TYPE As Long
Public RATE_TYPE_CUS As Long
Public PACKAGE_RATE As Double
Public PART_DESC As String
Public PART_TYPE As Long

Public EX_DELIVERY_COST_ITEM_ID As Long
Public EX_DELIVERY_COST_ITEM_CODE As String
Public EX_DELIVERY_COST_ITEM_NAME As String
Public RATE_DELIVERY As Double
Public RATE_CUSTOMER As Double
Public WEIGHT_PER_PACK As Long
Public WEIGHT_PER_PACK_CUS As Long
Public CUSTOMER_ID As Long
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public DELIVERY_CUS_ITEM_ID As Long
Public DELIVERY_CUS_ITEM_CODE As String
Public DELIVERY_CUS_ITEM_NAME As String

Public EX_PROMOTION_PART_ITEM_ID As Long
Public DISCOUNT_AMOUNT As Double

Public EX_PROMOTION_DLC_ITEM_ID As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String
Public MODIFY_DATE As Date
Public CREATE_BY As Long
Public MODIFY_BY As Long

Public ExWorksPriceItem As Collection
Public ExDeliveryCost As Collection
Public ExPromotionPart As Collection
Public ExPromotionDlc As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ExWorksPriceItem = New Collection
   Set ExDeliveryCost = New Collection
   Set ExPromotionPart = New Collection
   Set ExPromotionDlc = New Collection
   FROM_ACTIVE_DATE = -1
   TO_VALID_DATE = -1
   BETWEEN_DATE = -1
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
Dim TempDate As String

If Ind = 1 Then
   SelectStr1 = "SELECT EWP.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EX_WORKS_PRICE  EWP "
   WhereStr = ""
ElseIf Ind = 2 Then
   SelectStr1 = "SELECT EWP.*,EWPI.EX_WORKS_PRICE_ITEM_ID,EWPI.RATE_TYPE,EWPI.PACKAGE_RATE,PI.PART_ITEM_ID,PI.PART_NO,PI.PART_DESC,PI.PART_TYPE "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EX_WORKS_PRICE  EWP "
   FromStr = FromStr & "LEFT OUTER JOIN EX_WORKS_PRICE_ITEM EWPI ON (EWP.EX_WORKS_PRICE_ID = EWPI.EX_WORKS_PRICE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EWPI.PART_ITEM_ID = PI.PART_ITEM_ID) "
   WhereStr = ""
ElseIf Ind = 3 Then
   SelectStr1 = "SELECT EWP.*,EDCI.EX_DELIVERY_COST_ITEM_ID,EDCI.RATE_TYPE,EDCI.RATE_TYPE_CUS,EDCI.RATE_DELIVERY,EDCI.RATE_CUSTOMER,EDCI.CUSTOMER_ID,EDCI.WEIGHT_PER_PACK,EDCI.WEIGHT_PER_PACK_CUS, "
   SelectStr1 = SelectStr1 & "DCI.DELIVERY_CUS_ITEM_ID,DCI.DELIVERY_CUS_ITEM_CODE,DCI.DELIVERY_CUS_ITEM_NAME,EDCI.RATE_TYPE,CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME "

   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EX_WORKS_PRICE  EWP "
   FromStr = FromStr & "LEFT OUTER JOIN EX_DELIVERY_COST_ITEM EDCI ON (EWP.EX_WORKS_PRICE_ID = EDCI.EX_WORKS_PRICE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (EDCI.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN DELIVERY_CUS_ITEM DCI ON (EDCI.DELIVERY_CUS_ITEM_ID = DCI.DELIVERY_CUS_ITEM_ID) "
   WhereStr = ""
ElseIf Ind = 4 Then
   SelectStr1 = "SELECT EWP.*,EPPI.EX_PROMOTION_PART_ITEM_ID,EPPI.CUSTOMER_ID,EPPI.RATE_TYPE,EPPI.DISCOUNT_AMOUNT,PI.PART_ITEM_ID,PI.PART_NO,PI.PART_DESC,PI.PART_TYPE,CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME "

   SelectStr2 = "SELECT COUNT(*)  "
   FromStr = "FROM EX_WORKS_PRICE  EWP "
   FromStr = FromStr & "LEFT OUTER JOIN EX_PROMOTION_PART_ITEM  EPPI ON (EWP.EX_WORKS_PRICE_ID = EPPI.EX_WORKS_PRICE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (EPPI.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EPPI.PART_ITEM_ID = PI.PART_ITEM_ID) "
   WhereStr = ""
ElseIf Ind = 5 Then
  SelectStr1 = "SELECT EWP.*,EPDI.EX_PROMOTION_DLC_ITEM_ID,EPDI.RATE_TYPE_CUS,EPDI.DISCOUNT_AMOUNT,EPDI.CUSTOMER_ID,EPDI.WEIGHT_PER_PACK_CUS, "
   SelectStr1 = SelectStr1 & "DCI.DELIVERY_CUS_ITEM_ID,DCI.DELIVERY_CUS_ITEM_CODE,DCI.DELIVERY_CUS_ITEM_NAME,CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME "

   SelectStr2 = "SELECT COUNT(*)  "
   FromStr = "FROM EX_WORKS_PRICE  EWP "
   FromStr = FromStr & "LEFT OUTER JOIN EX_PROMOTION_DLC_ITEM  EPDI ON (EWP.EX_WORKS_PRICE_ID = EPDI.EX_WORKS_PRICE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (EPDI.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN DELIVERY_CUS_ITEM DCI ON (EPDI.DELIVERY_CUS_ITEM_ID = DCI.DELIVERY_CUS_ITEM_ID) "
   WhereStr = ""
Else
   SelectStr1 = "SELECT EWP.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EX_WORKS_PRICE  EWP "
   WhereStr = ""
End If
   
   If EX_WORKS_PRICE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWP.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      End If
   End If
      
   If Len(EX_WORKS_PRICE_LEVEL) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWP.EX_WORKS_PRICE_LEVEL = '" & EX_WORKS_PRICE_LEVEL & "') "
      Else
         WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_LEVEL = '" & EX_WORKS_PRICE_LEVEL & "') "
      End If
   End If
   
   If EX_WORKS_PRICE_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWP.EX_WORKS_PRICE_TYPE = " & EX_WORKS_PRICE_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_TYPE = " & EX_WORKS_PRICE_TYPE & ") "
      End If
   End If

      
   If EX_WORKS_PRICE_STATUS > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWP.EX_WORKS_PRICE_STATUS = " & EX_WORKS_PRICE_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_STATUS = " & EX_WORKS_PRICE_STATUS & ") "
      End If
   End If
   
   If Trim(EX_WORKS_PRICE_CODE) <> "" Then
      If WildCard(EX_WORKS_PRICE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EWP.EX_WORKS_PRICE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EWP.EX_WORKS_PRICE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EWP.EX_WORKS_PRICE_CODE = '" & ChangeQuote(Trim(EX_WORKS_PRICE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_CODE = '" & ChangeQuote(Trim(EX_WORKS_PRICE_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(EX_WORKS_PRICE_DESC) <> "" Then
      If WildCard(EX_WORKS_PRICE_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EWP.EX_WORKS_PRICE_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EWP.EX_WORKS_PRICE_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EWP.EX_WORKS_PRICE_DESC = '" & ChangeQuote(Trim(EX_WORKS_PRICE_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_DESC = '" & ChangeQuote(Trim(EX_WORKS_PRICE_DESC)) & "')"
         End If
      End If
   End If
   
   If EX_WORKS_PRICE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(EX_WORKS_PRICE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (EWP.EX_WORKS_PRICE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EWP.EX_WORKS_PRICE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_ACTIVE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_ACTIVE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (EWP.FROM_ACTIVE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EWP.FROM_ACTIVE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_VALID_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(TO_VALID_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (EWP.TO_VALID_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EWP.TO_VALID_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If BETWEEN_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(BETWEEN_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (EWP.TO_VALID_DATE >= '" & ChangeQuote(TempDate) & "')"
         WhereStr = WhereStr & "AND (EWP.FROM_ACTIVE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (EWP.TO_VALID_DATE >= '" & ChangeQuote(TempDate) & "')"
         WhereStr = WhereStr & "AND (EWP.FROM_ACTIVE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
      If Trim(PART_NO) <> "" Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_CODE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_DATE "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY EWP.FROM_ACTIVE_DATE "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY EWP.TO_VALID_DATE "
    Else
         OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_CODE  "
    End If
ElseIf Ind = 3 Then
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_DATE DESC ,EWP.EX_WORKS_PRICE_CODE,CM.CUSTOMER_CODE ASC,DCI.DELIVERY_CUS_ITEM_CODE ASC,EDCI.RATE_TYPE ASC,EDCI.RATE_TYPE_CUS "
   Else
      OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_ID "
   End If
ElseIf Ind = 4 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC,PI.PART_NO  "
    Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC,PI.PART_NO   "
    End If
ElseIf Ind = 5 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DCI.DELIVERY_CUS_ITEM_CODE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY DCI.DELIVERY_CUS_ITEM_CODE ASC,DCI.DELIVERY_CUS_ITEM_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC,DCI.DELIVERY_CUS_ITEM_CODE  "
    Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC,DCI.DELIVERY_CUS_ITEM_CODE   "
    End If
Else
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_DATE DESC ,EWP.EX_WORKS_PRICE_CODE,PI.PART_NO "
   Else
      OrderByStr = "ORDER BY EWP.EX_WORKS_PRICE_ID "
   End If
End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
   EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
   EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
   EX_WORKS_PRICE_DATE = InternalDateToDate(NVLS(Rs("EX_WORKS_PRICE_DATE"), ""))
   EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
   EX_WORKS_PRICE_STATUS = NVLI(Rs("EX_WORKS_PRICE_STATUS"), -1)
   EX_WORKS_PRICE_LEVEL = NVLS(Rs("EX_WORKS_PRICE_LEVEL"), "")
   EX_WORKS_PRICE_TYPE = NVLI(Rs("EX_WORKS_PRICE_TYPE"), -1)
   FROM_ACTIVE_DATE = InternalDateToDate(NVLS(Rs("FROM_ACTIVE_DATE"), ""))
   TO_VALID_DATE = InternalDateToDate(NVLS(Rs("TO_VALID_DATE"), ""))
   MODIFY_DATE = InternalDateToDate(NVLS(Rs("MODIFY_DATE"), ""))
   
   KEY_ID = EX_WORKS_PRICE_ID
   KEY_LOOKUP = EX_WORKS_PRICE_CODE
   
   CREATE_BY = NVLI(Rs("CREATE_BY"), 0)
   MODIFY_BY = NVLI(Rs("MODIFY_BY"), 0)
ElseIf Ind = 2 Then
   EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
   EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
   EX_WORKS_PRICE_DATE = InternalDateToDate(NVLS(Rs("EX_WORKS_PRICE_DATE"), ""))
   EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
   EX_WORKS_PRICE_STATUS = NVLI(Rs("EX_WORKS_PRICE_STATUS"), -1)
   EX_WORKS_PRICE_LEVEL = NVLS(Rs("EX_WORKS_PRICE_LEVEL"), "")
   EX_WORKS_PRICE_TYPE = NVLI(Rs("EX_WORKS_PRICE_TYPE"), -1)
   FROM_ACTIVE_DATE = InternalDateToDate(NVLS(Rs("FROM_ACTIVE_DATE"), ""))
   TO_VALID_DATE = InternalDateToDate(NVLS(Rs("TO_VALID_DATE"), ""))
   
   EX_WORKS_PRICE_ITEM_ID = NVLI(Rs("EX_WORKS_PRICE_ITEM_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   PART_NO = NVLS(Rs("PART_NO"), "")
   RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
   PACKAGE_RATE = NVLD(Rs("PACKAGE_RATE"), 0)
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
ElseIf Ind = 3 Then
   EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
   EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
   EX_WORKS_PRICE_DATE = InternalDateToDate(NVLS(Rs("EX_WORKS_PRICE_DATE"), ""))
   EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
   EX_WORKS_PRICE_STATUS = NVLI(Rs("EX_WORKS_PRICE_STATUS"), -1)
   EX_WORKS_PRICE_LEVEL = NVLS(Rs("EX_WORKS_PRICE_LEVEL"), "")
   EX_WORKS_PRICE_TYPE = NVLI(Rs("EX_WORKS_PRICE_TYPE"), -1)
   FROM_ACTIVE_DATE = InternalDateToDate(NVLS(Rs("FROM_ACTIVE_DATE"), ""))
   TO_VALID_DATE = InternalDateToDate(NVLS(Rs("TO_VALID_DATE"), ""))
   
   EX_DELIVERY_COST_ITEM_ID = NVLI(Rs("EX_DELIVERY_COST_ITEM_ID"), -1)
   RATE_DELIVERY = NVLD(Rs("RATE_DELIVERY"), 0)
   RATE_CUSTOMER = NVLD(Rs("RATE_CUSTOMER"), 0)
   WEIGHT_PER_PACK = NVLI(Rs("WEIGHT_PER_PACK"), -1)
   WEIGHT_PER_PACK_CUS = NVLI(Rs("WEIGHT_PER_PACK_CUS"), -1)
   RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
   RATE_TYPE_CUS = NVLI(Rs("RATE_TYPE_CUS"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   DELIVERY_CUS_ITEM_ID = NVLI(Rs("DELIVERY_CUS_ITEM_ID"), -1)
   DELIVERY_CUS_ITEM_CODE = NVLS(Rs("DELIVERY_CUS_ITEM_CODE"), "")
   DELIVERY_CUS_ITEM_NAME = NVLS(Rs("DELIVERY_CUS_ITEM_NAME"), "")
ElseIf Ind = 4 Then
   EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
   EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
   EX_WORKS_PRICE_DATE = InternalDateToDate(NVLS(Rs("EX_WORKS_PRICE_DATE"), ""))
   EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
   EX_WORKS_PRICE_STATUS = NVLI(Rs("EX_WORKS_PRICE_STATUS"), -1)
   EX_WORKS_PRICE_LEVEL = NVLS(Rs("EX_WORKS_PRICE_LEVEL"), "")
   EX_WORKS_PRICE_TYPE = NVLI(Rs("EX_WORKS_PRICE_TYPE"), -1)
   FROM_ACTIVE_DATE = InternalDateToDate(NVLS(Rs("FROM_ACTIVE_DATE"), ""))
   TO_VALID_DATE = InternalDateToDate(NVLS(Rs("TO_VALID_DATE"), ""))
   
   EX_PROMOTION_PART_ITEM_ID = NVLI(Rs("EX_PROMOTION_PART_ITEM_ID"), -1)
   DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   PART_NO = NVLS(Rs("PART_NO"), "")
   RATE_TYPE = NVLI(Rs("RATE_TYPE"), -1)
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
ElseIf Ind = 5 Then
   EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
   EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
   EX_WORKS_PRICE_DATE = InternalDateToDate(NVLS(Rs("EX_WORKS_PRICE_DATE"), ""))
   EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
   EX_WORKS_PRICE_STATUS = NVLI(Rs("EX_WORKS_PRICE_STATUS"), -1)
   EX_WORKS_PRICE_LEVEL = NVLS(Rs("EX_WORKS_PRICE_LEVEL"), "")
   EX_WORKS_PRICE_TYPE = NVLI(Rs("EX_WORKS_PRICE_TYPE"), -1)
   FROM_ACTIVE_DATE = InternalDateToDate(NVLS(Rs("FROM_ACTIVE_DATE"), ""))
   TO_VALID_DATE = InternalDateToDate(NVLS(Rs("TO_VALID_DATE"), ""))
   
   EX_PROMOTION_DLC_ITEM_ID = NVLI(Rs("EX_PROMOTION_DLC_ITEM_ID"), -1)
   DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   WEIGHT_PER_PACK_CUS = NVLI(Rs("WEIGHT_PER_PACK_CUS"), -1)
   RATE_TYPE_CUS = NVLI(Rs("RATE_TYPE_CUS"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   DELIVERY_CUS_ITEM_ID = NVLI(Rs("DELIVERY_CUS_ITEM_ID"), -1)
   DELIVERY_CUS_ITEM_CODE = NVLS(Rs("DELIVERY_CUS_ITEM_CODE"), "")
   DELIVERY_CUS_ITEM_NAME = NVLS(Rs("DELIVERY_CUS_ITEM_NAME"), "")
End If
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EX_WORKS_PRICE_SEQ", ID, glbErrorLog)
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("EX_WORKS_PRICE_SEQ", ID, glbErrorLog)
      Else
         ID = EX_WORKS_PRICE_ID
      End If
      
      SQL1 = "INSERT INTO EX_WORKS_PRICE " & _
                     "( " & _
                     "   EX_WORKS_PRICE_ID , " & _
                     "   EX_WORKS_PRICE_CODE, " & _
                     "   EX_WORKS_PRICE_DATE, " & _
                     "   EX_WORKS_PRICE_DESC, " & _
                     "   EX_WORKS_PRICE_STATUS, " & _
                     "   EX_WORKS_PRICE_LEVEL, " & _
                     "   EX_WORKS_PRICE_TYPE, " & _
                     "   FROM_ACTIVE_DATE, " & _
                     "   TO_VALID_DATE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(EX_WORKS_PRICE_CODE)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EX_WORKS_PRICE_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(EX_WORKS_PRICE_DESC)) & "'," & _
                         "" & EX_WORKS_PRICE_STATUS & "," & _
                         "'" & EX_WORKS_PRICE_LEVEL & "'," & _
                         "" & EX_WORKS_PRICE_TYPE & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(FROM_ACTIVE_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(TO_VALID_DATE))) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EX_WORKS_PRICE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EX_WORKS_PRICE SET " & _
                     "   EX_WORKS_PRICE_CODE = '" & ChangeQuote(Trim(EX_WORKS_PRICE_CODE)) & "', " & _
                     "   EX_WORKS_PRICE_DATE  = '" & ChangeQuote(Trim(DateToStringInt(EX_WORKS_PRICE_DATE))) & "', " & _
                     "   EX_WORKS_PRICE_DESC  = '" & ChangeQuote(Trim(EX_WORKS_PRICE_DESC)) & "', " & _
                     "   EX_WORKS_PRICE_STATUS  = " & EX_WORKS_PRICE_STATUS & ", " & _
                      "   EX_WORKS_PRICE_LEVEL  = '" & EX_WORKS_PRICE_LEVEL & "', " & _
                      "   EX_WORKS_PRICE_TYPE  = " & EX_WORKS_PRICE_TYPE & ", " & _
                      "   FROM_ACTIVE_DATE   = '" & ChangeQuote(Trim(DateToStringInt(FROM_ACTIVE_DATE))) & "', " & _
                      "   TO_VALID_DATE   = '" & ChangeQuote(Trim(DateToStringInt(TO_VALID_DATE))) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM EX_WORKS_PRICE_ITEM WHERE EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EX_DELIVERY_COST_ITEM WHERE EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EX_PROMOTION_PART_ITEM WHERE EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EX_PROMOTION_DLC_ITEM WHERE EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID
      m_Conn.Execute (SQL1)
            
      SQL1 = "DELETE FROM EX_WORKS_PRICE WHERE EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ExWorksPriceItem = Nothing
   Set ExDeliveryCost = Nothing
   Set ExPromotionPart = Nothing
   Set ExPromotionDlc = Nothing
End Sub

