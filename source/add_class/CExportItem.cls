VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExportItem"
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

Public EXPORT_ITEM_ID As Long
Public INVENTORY_DOC_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public HOUSE_ID As Long
Public PIG_ID As Long
Public PIG_IN_HOUSE As Double
Public EXPORT_AMOUNT As Double
Public EXPORT_AVG_PRICE As Double
Public EXPORT_LAST_PRICE As Double
Public PREVIOUS_AMOUNT As Double
Public CURRENT_AMOUNT As Double
Public TRANSACTION_SEQ As Long
Public GUI_ID As Long
Public CALCULATE_FLAG As String
Public PIG_STATUS As Long
Public TOTAL_WEIGHT As Double
Public LINK_ID As Long

Public TX_TYPE As String
Public PIG_TYPE As String
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long
Public PART_TYPE_ID As String
Public PART_TYPE_NO As String
Public PART_TYPE_NAME As String
Public LOCATION_NAME As String
Public LOCATION_NO As String
Public UNIT_NAME As String
Public OLD_EXPORT_AMOUNT As Double
Public OLD_LOCATION_ID As Long
Public PART_PIG_TYPE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public COMMIT_FLAG As String
Public DOCUMENT_TYPE As Long
Public PIG_FLAG As String
Public PART_GROUP_ID As Long
Public EXPORT_TOTAL_PRICE As Double
Public HOUSE_GROUP_ID As Long
Public STATUS_GROUP_ID As Long
Public FROM_DATE1 As Date
Public TO_DATE1 As Date
Public HOUSE_ID1 As Long
Public COMMIT_FLAG1 As String
Public PIG_NO As String
Public HOUSE_NAME As String
Public SUM_ALL_EXPORT_AMT As Double
Public SUM_ALL_EXPORT As Double
Public SALE_FLAG As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   TX_TYPE = "E"
   FROM_DATE = -1
   TO_DATE = -1
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
Dim GroupByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim FromDate1 As String
Dim ToDate1 As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, PT.*, LC.*, HS.LOCATION_NAME HOUSE_NAME, PI.*, PG.PART_NO PIG_NO, PG.PIG_TYPE PIG_TYPE2, PI.PIG_TYPE PIG_TYPE1, UN.*, IVD.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (UG.HOUSE_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (UG.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (UG.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PTT.PART_TYPE_NO, PTT.PART_TYPE_NAME, PI.PART_NO, PI.PART_DESC, "
      SelectStr1 = SelectStr1 & "IVD.DOCUMENT_NO, IVD.DOCUMENT_DATE, EI.EXPORT_AMOUNT, UN.UNIT_NAME, "
      SelectStr1 = SelectStr1 & "LC1.LOCATION_NAME LC_NAME1, LC1.LOCATION_NO LC_NO1, LC2.LOCATION_NAME LC_NAME2, LC2.LOCATION_NO LC_NO2 "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_TYPE PTT, PART_ITEM PI, UNIT UN, LOCATION LC1, LOCATION LC2, IMPORT_ITEM II "

      WhereStr = "WHERE (Ei.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      WhereStr = WhereStr & "AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      WhereStr = WhereStr & "AND (PI.UNIT_COUNT = UN.UNIT_ID) "
      WhereStr = WhereStr & "AND (EI.GUI_ID = II.GUI_ID) "
      WhereStr = WhereStr & "AND (EI.LOCATION_ID = LC1.LOCATION_ID) "
      WhereStr = WhereStr & "AND (II.LOCATION_ID = LC2.LOCATION_ID) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 3) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT EI.HOUSE_ID, EI.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.HOUSE_ID = LC.LOCATION_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT EI.HOUSE_ID, EI.PIG_ID, PT.PART_GROUP_ID, SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT PT1.*, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM (EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_TYPE PT2 ON (PI.PART_TYPE = PT2.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   WHERE (EI.HOUSE_ID IN (SELECT HGI.LOCATION_ID FROM HGROUP_ITEM HGI WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PT1.PART_TYPE_ID) "
      If Len(COMMIT_FLAG1) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG1 & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) "
      If FROM_DATE1 > 0 Then
         FromDate1 = DateToStringIntLow(Trim(FROM_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      End If
      If TO_DATE1 > 0 Then
         ToDate1 = DateToStringIntLow(Trim(TO_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      End If
      SelectStr1 = SelectStr1 & ") EXPORT_TOTAL_PRICE "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM PART_TYPE PT1 "

      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT PT1.*, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM (EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_TYPE PT2 ON (PI.PART_TYPE = PT2.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   WHERE (EI.HOUSE_ID = " & HOUSE_ID1 & ") "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PT1.PART_TYPE_ID) "
      If Len(COMMIT_FLAG1) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG1 & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) "
      If FROM_DATE1 > 0 Then
         FromDate1 = DateToStringIntLow(Trim(FROM_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      End If
      If TO_DATE1 > 0 Then
         ToDate1 = DateToStringIntLow(Trim(TO_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      End If
      SelectStr1 = SelectStr1 & ") EXPORT_TOTAL_PRICE "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM PART_TYPE PT1 "

      WhereStr = ""
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.HOUSE_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.HOUSE_ID, EI.PIG_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID, EI.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IN (SELECT HGI.LOCATION_ID FROM HGROUP_ITEM HGI WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT PI.PART_NO, PI.PIG_TYPE, EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) SUM_ALL_EXPORT_AMT, SUM(EI.EXPORT_AMOUNT) SUM_ALL_EXPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE >= '" & DateToStringIntLow(FROM_DATE) & "')"
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE < '" & DateToStringIntLow(TO_DATE) & "')"
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, IVD.DOCUMENT_TYPE, IVD.SALE_FLAG, SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) SUM_EXPORT_AMT, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
   ElseIf (Ind = 13) Or (Ind = 17) Then
      SelectStr1 = "SELECT UN.UNIT_NAME, EI.*, PT.PART_TYPE_NO, PT.PART_TYPE_NAME, IVD.DOCUMENT_TYPE, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, PI.PART_NO, PI.PART_DESC, "
      SelectStr1 = SelectStr1 & "PI.PIG_TYPE, PI.PART_TYPE, EI.EXPORT_AMOUNT TX_AMOUNT, EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE TX_PRICE, "
      SelectStr1 = SelectStr1 & "EI.PIG_STATUS, EI.TRANSACTION_SEQ, IVD.SALE_FLAG "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IN (SELECT HGI.LOCATION_ID FROM HGROUP_ITEM HGI WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT DISTINCT EI.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PIG_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT EI.PIG_ID, PI1.PIG_TYPE, PI2.PART_TYPE, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT_AMOUNT, SUM(EI.EXPORT_AVG_PRICE * EI.EXPORT_AMOUNT) SUM_EXPORT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI1 ON (EI.PIG_ID = PI1.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (EI.PART_ITEM_ID = PI2.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 18 Then '17 is already used
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE, EI.TRANSACTION_SEQ, EI.LOCATION_ID, EI.PART_ITEM_ID, EI.CURRENT_AMOUNT, EI.EXPORT_AVG_PRICE , 'E' "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE EI.TRANSACTION_SEQ IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX(EI1.TRANSACTION_SEQ) "
      WhereStr = WhereStr & "   FROM EXPORT_ITEM EI1 "
      WhereStr = WhereStr & "   LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (EI1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "   WHERE (IVD1.DOCUMENT_DATE = IVD.DOCUMENT_DATE) "
      WhereStr = WhereStr & "   AND (EI1.PART_ITEM_ID = EI.PART_ITEM_ID) "
      WhereStr = WhereStr & "   AND (EI1.LOCATION_ID = EI.LOCATION_ID) "
      WhereStr = WhereStr & ") "
   ElseIf Ind = 19 Then
      SelectStr1 = "SELECT EI.LOCATION_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT_AMOUNT, SUM(EI.EXPORT_AVG_PRICE * EI.EXPORT_AMOUNT) SUM_EXPORT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   ElseIf Ind = 20 Then
      SelectStr1 = "SELECT EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'Y') "
      WhereStr = WhereStr & "AND (EI.PIG_STATUS NOT IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 21 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT_AMOUNT, SUM(EI.EXPORT_AVG_PRICE * EI.EXPORT_AMOUNT) SUM_EXPORT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   ElseIf Ind = 22 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'Y') "
      WhereStr = WhereStr & "AND (EI.PIG_STATUS NOT IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 23 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   End If
   
   If EXPORT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      End If
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If HOUSE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (HOUSE_ID = " & HOUSE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (HOUSE_ID = " & HOUSE_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PI.PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
'   If Ind = 2 Then
      If PART_TYPE > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (PI.PART_TYPE = " & PART_TYPE & ") "
         Else
            WhereStr = WhereStr & "AND (PI.PART_TYPE = " & PART_TYPE & ") "
         End If
      End If
      
       If FROM_DATE > -1 Then
         TempDate = DateToStringIntLow(Trim(FROM_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
   
      If TO_DATE > -1 Then
         TempDate = DateToStringIntHi(Trim(TO_DATE))
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
         End If
      End If
'   End If
   
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
   
   GroupByStr = ""
   
   If Ind = 4 Then
      GroupByStr = "GROUP BY EI.HOUSE_ID, EI.PIG_ID, PT.PART_GROUP_ID "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.HOUSE_ID "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.HOUSE_ID, EI.PIG_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY PI.PART_NO, EI.LOCATION_ID, PI.PIG_TYPE "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, IVD.DOCUMENT_TYPE, IVD.SALE_FLAG "
   ElseIf Ind = 16 Then
      GroupByStr = "GROUP BY EI.PIG_ID, PI1.PIG_TYPE, PI2.PART_TYPE "
   ElseIf Ind = 19 Then
      GroupByStr = "GROUP BY EI.LOCATION_ID, EI.PIG_STATUS "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY EI.LOCATION_ID "
   ElseIf Ind = 21 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_STATUS "
   ElseIf Ind = 22 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 23 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY LC.LOCATION_NO "
      Else
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOCATION_NO "
      Else
         OrderByStr = "ORDER BY LOCATION_NAME "
      End If
   ElseIf Ind = 4 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      Else
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      End If
   ElseIf Ind = 5 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      End If
   ElseIf Ind = 6 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      End If
   ElseIf Ind = 7 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      Else
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      End If
   ElseIf Ind = 8 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      End If
   ElseIf Ind = 9 Then
      Dim TempOrderType As String
         
      If OrderType = 1 Then
         TempOrderType = "ASC"
      Else
         TempOrderType = "DESC"
      End If
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO " & TempOrderType & ", PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NO " & TempOrderType & ", PART_NO "
      Else
         OrderByStr = "ORDER BY PART_TYPE_NO " & TempOrderType & ", PART_NO "
      End If
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 13 Then
      If OrderType = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, EI.TRANSACTION_SEQ ASC "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO DESC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, EI.TRANSACTION_SEQ ASC "
      End If
   ElseIf Ind = 14 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      End If
   ElseIf Ind = 15 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      Else
         OrderByStr = "ORDER BY PI.PART_NO  "
      End If
   ElseIf Ind = 16 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EI.PART_ITEM_ID  "
      End If
   ElseIf Ind = 17 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      Else
         OrderByStr = "ORDER BY PI.PART_NO  "
      End If
   ElseIf Ind = 18 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      End If
   ElseIf Ind = 19 Then
      OrderByStr = "ORDER BY EI.LOCATION_ID "
   ElseIf Ind = 20 Then
      OrderByStr = "ORDER BY EI.LOCATION_ID "
   ElseIf Ind = 21 Then
      OrderByStr = "ORDER BY EI.PART_ITEM_ID "
   ElseIf Ind = 22 Then
      OrderByStr = "ORDER BY EI.PART_ITEM_ID "
   ElseIf Ind = 23 Then
      OrderByStr = "ORDER BY EI.PART_ITEM_ID "
   End If
   
   If Ind <> 13 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EXPORT_ITEM_ID = NVLI(Rs("EXPORT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PIG_IN_HOUSE = NVLD(Rs("PIG_IN_HOUSE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      EXPORT_LAST_PRICE = NVLD(Rs("EXPORT_LAST_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      
      OLD_EXPORT_AMOUNT = EXPORT_AMOUNT
      OLD_LOCATION_ID = LOCATION_ID
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_PIG_TYPE = NVLS(Rs("PIG_TYPE2"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE1"), "")
      TX_TYPE = "E" 'Export
      
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      HOUSE_NAME = NVLS(Rs("HOUSE_NAME"), "")
   ElseIf Ind = 4 Then
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 5 Then
      PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 6 Then
      PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 7 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 8 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 10 Then
      PART_NO = NVLS(Rs("PART_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 11 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_ALL_EXPORT_AMT"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_ALL_EXPORT"), 0)
   ElseIf Ind = 12 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_AMT"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT"), 0)
   ElseIf (Ind = 13) Or (Ind = 17) Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      EXPORT_TOTAL_PRICE = EXPORT_AVG_PRICE * EXPORT_AMOUNT
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
   ElseIf Ind = 16 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT_AMOUNT"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_PRICE"), 0)
   ElseIf Ind = 18 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 19 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_PRICE"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT_AMOUNT"), 0)
   ElseIf Ind = 20 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 21 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_PRICE"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT_AMOUNT"), 0)
   ElseIf Ind = 22 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 23 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = PIG_ID
   If PIG_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = HOUSE_ID
   If HOUSE_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PIG_STATUS
   If PIG_STATUS <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXPORT_ITEM_SEQ", ID, glbErrorLog)
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "INSERT INTO EXPORT_ITEM " & _
                     "( " & _
                     "   EXPORT_ITEM_ID, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   HOUSE_ID, " & _
                     "   PIG_ID, " & _
                     "   PIG_IN_HOUSE, " & _
                     "   EXPORT_AMOUNT, " & _
                     "   EXPORT_AVG_PRICE, " & _
                     "   EXPORT_LAST_PRICE, " & _
                     "   PREVIOUS_AMOUNT, " & _
                     "   CURRENT_AMOUNT, "
            SQL1 = SQL1 & _
                     "   TRANSACTION_SEQ, " & _
                     "   GUI_ID, " & _
                     "   CALCULATE_FLAG, " & _
                     "   PIG_STATUS, " & _
                     "   TOTAL_WEIGHT, " & _
                     "   LINK_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & INVENTORY_DOC_ID & "," & _
                         "" & PART_ITEM_ID & "," & _
                         "" & LOCATION_ID & "," & _
                         TempID2 & "," & _
                         TempID1 & "," & _
                         "" & PIG_IN_HOUSE & "," & _
                         "" & EXPORT_AMOUNT & "," & _
                         "" & EXPORT_AVG_PRICE & "," & _
                         "" & EXPORT_LAST_PRICE & "," & _
                         "" & PREVIOUS_AMOUNT & ","
               SQL1 = SQL1 & _
                         "" & CURRENT_AMOUNT & "," & _
                         "" & ID2 & "," & _
                         "" & GUI_ID & "," & _
                         "'" & CALCULATE_FLAG & "'," & _
                         TempID3 & "," & _
                         TOTAL_WEIGHT & "," & _
                         "" & LINK_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EXPORT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXPORT_ITEM SET " & _
                     "   INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   LOCATION_ID = " & LOCATION_ID & ", " & _
                     "   HOUSE_ID = " & TempID2 & ", " & _
                     "   PIG_ID = " & TempID1 & ", " & _
                     "   PIG_IN_HOUSE = " & PIG_IN_HOUSE & ", " & _
                     "   EXPORT_AMOUNT = " & EXPORT_AMOUNT & ", " & _
                     "   EXPORT_AVG_PRICE = " & EXPORT_AVG_PRICE & ", " & _
                     "   EXPORT_LAST_PRICE = " & EXPORT_LAST_PRICE & ", " & _
                     "   PREVIOUS_AMOUNT = " & PREVIOUS_AMOUNT & ", " & _
                     "   CURRENT_AMOUNT = " & CURRENT_AMOUNT & ", " & _
                     "   TRANSACTION_SEQ = " & TRANSACTION_SEQ & ", " & _
                     "   GUI_ID = " & GUI_ID & ", " & _
                     "   CALCULATE_FLAG = '" & CALCULATE_FLAG & "', " & _
                     "   PIG_STATUS = " & TempID3 & ", " & _
                     "   TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EXPORT_ITEM WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

