VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJob"
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

Public JOB_ID As Long
Public JOB_NO As String
Public JOB_NO_LIKE As String
Public JOB_DATE As Date
Public JOB_DESC As String
Public PROCESS_ID As Long
Public PROCESS_ID_IN  As String
Public BATCH_NO As String
Public FROM_BATCH_NO As Long
Public TO_BATCH_NO As Long
Public BATCH_TOTAL As Double
Public BATCH_DETAIL As String
Public LOT_ID As Integer
Public LOT_NO As String
Public BIN_NO As Integer
Public BIN_NAME As String
Public APPROVED_BY As Long
Public RESPONSE_BY As Long
Public START_DATE  As Date
Public FINISH_DATE As Date
Public COMMIT_FLAG As String
Public INVENTORY_DOC_ID As Long
Public INVENTORY_WH_DOC_ID As Long
Public INVENTORY_WH_DOC_ID_INPUT As Long
Public JOB_DOC_TYPE As Long
Public PART_ITEM_ID As Long
Public STD_AMOUNT As Double
Public ACTUAL_AMOUNT As Double
Public FORMULA_ID As Long

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public PARCEL_TYPE As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public LONG_NAMEA As String
Public LAST_NAMEA As String
Public LONG_NAMER As String
Public LAST_NAMER As String
Public DOC_NO As String
Public PROCESS_NAME As String
Public OLD_COMMIT_FLAG As String
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long
Public LOCK_DOC_FLAG As String
Public JOB_ID_REF As Long

Public ProcessSet As String
Public YYYYMM As String

Public SUM_INPUT As Double
Public SUM_OUTPUT As Double

Public CREATE_NAME As String
Public MODIFY_NAME As String
Public CREATE_BY As Long
Public MODIFY_BY As Long
Public VERIFY_FLAG As String
Public VERIFY_NAME As String

Public NOTE As String

Public Inputs As Collection
Public Outputs As Collection
Public Pallets As Collection
Public Peoples As Collection
Public Machines As Collection
Public Parameters As Collection
Public Verifies As Collection
Public InventoryWhDoc As Collection
Public InventoryWhDocInput As Collection
Public tempIWDInput As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   START_DATE = -1
   FINISH_DATE = -1
   COMMIT_FLAG = ""
    JOB_DATE = -1
    
   FROM_DATE = -1
   TO_DATE = -1
   
   Set Inputs = New Collection
   Set Outputs = New Collection
   Set Pallets = New Collection
   Set Peoples = New Collection
   Set Machines = New Collection
   Set Parameters = New Collection
   Set Verifies = New Collection
'   Set InventoryWhDoc = New Collection
'   Set InventoryWhDocInput = New Collection
End Sub

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
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT J.*,NM1.LONG_NAME LONGA,NM1.LAST_NAME LASTA,NM2.LONG_NAME LONGR,NM2.LAST_NAME LASTR,DOC.DOCUMENT_NO,PR.PROCESS_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
       SelectStr1 = SelectStr1 & ",(SELECT SUM(INP.TX_AMOUNT)FROM JOB_INOUT INP WHERE (INP.JOB_ID = J.JOB_ID) AND(INP.TX_TYPE = 'E')) SUM_INPUT, "
       SelectStr1 = SelectStr1 & "(SELECT SUM(INP.TX_AMOUNT)FROM JOB_INOUT INP WHERE (INP.JOB_ID = J.JOB_ID) AND(INP.TX_TYPE = 'I')) SUM_OUTPUT "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PROCESS PR ON (J.PROCESS_ID = PR.PROCESS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC DOC ON (J.INVENTORY_DOC_ID = DOC.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM1 ON (EM1.EMP_ID = J.APPROVED_BY) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM2 ON (EM2.EMP_ID = J.RESPONSE_BY) "
       FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EMN1 ON (EMN1.EMP_ID = EM1.EMP_ID) "
     FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EMN2 ON (EMN2.EMP_ID = EM2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (EMN1.NAME_ID = NM1.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (EMN2.NAME_ID = NM2.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC WH_DOC ON (J.INVENTORY_WH_DOC_ID = WH_DOC.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LTW ON (WH_DOC.INVENTORY_WH_DOC_ID =LTW.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (LTW.LOT_ITEM_WH_ID =LTD.LOT_ITEM_WH_ID) "
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(J.ACTUAL_AMOUNT) ACTUAL_AMOUNT, SUM(J.STD_AMOUNT) STD_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (J.PART_ITEM_ID = PI.PART_ITEM_ID) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(J.ACTUAL_AMOUNT) ACTUAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",PI.PARCEL_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (J.PART_ITEM_ID = PI.PART_ITEM_ID) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (J.PART_ITEM_ID = PI.PART_ITEM_ID) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT J.JOB_NO,J.JOB_DATE,J.FROM_BATCH_NO,J.TO_BATCH_NO,J.BATCH_DETAIL,J.LOCK_DOC_FLAG,J.JOB_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT J.*,LT.LOT_ID,LT.LOT_NO,LTD.BIN_NO,LC.LOCATION_NAME BIN_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_WH_DOC WH_DOC ON (J.INVENTORY_WH_DOC_ID =WH_DOC.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM_WH LTW ON (WH_DOC.INVENTORY_WH_DOC_ID =LTW.INVENTORY_WH_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_DOC LTD ON (LTW.LOT_ITEM_WH_ID = LTD.LOT_ITEM_WH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOT LT ON (LTD.LOT_ID = LT.LOT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LTD.BIN_NO = LC.LOCATION_ID) "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT DISTINCT J.INVENTORY_WH_DOC_ID,J.VERIFY_FLAG "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM JOB J "
   End If
   
   If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
   
    If Trim(PROCESS_ID_IN) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (J.PROCESS_ID IN " & ChangeQuote(Trim(PROCESS_ID_IN)) & ")"
      Else
         WhereStr = WhereStr & "AND (J.PROCESS_ID IN " & ChangeQuote(Trim(PROCESS_ID_IN)) & ")"
      End If
   End If

   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_ID = " & JOB_ID & ") "
      End If
   End If
   
   If INVENTORY_WH_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.INVENTORY_WH_DOC_ID = " & INVENTORY_WH_DOC_ID & ") "
      End If
   End If
   
   If LOT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOT_ID = " & LOT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOT_ID = " & LOT_ID & ") "
      End If
   End If
 
    If Trim(JOB_NO) <> "" Then
      If WildCard(JOB_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         End If
      End If
   End If
   
   If Trim(JOB_NO_LIKE) <> "" Then
         If WhereStr = "" Then
            WhereStr = "WHERE (JOB_NO LIKE '" & ChangeQuote(Trim(JOB_NO_LIKE)) & "%')"
         Else
            WhereStr = WhereStr & "AND (JOB_NO LIKE '" & ChangeQuote(Trim(JOB_NO_LIKE)) & "%')"
         End If
   End If
 
   If Trim(JOB_DESC) <> "" Then
      If WildCard(JOB_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(JOB_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(JOB_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "')"
         End If
      End If
   End If
     
   If JOB_DATE > -1 Then
   TempDate = Trim(DateToStringInt(JOB_DATE)) 'JOB_DATE
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
 If Trim(BATCH_NO) <> "" Then
      If WildCard(BATCH_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(BATCH_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(BATCH_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (J.BATCH_NO = '" & ChangeQuote(Trim(BATCH_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (J.BATCH_NO = '" & ChangeQuote(Trim(BATCH_NO)) & "')"
         End If
      End If
   End If
   
   If APPROVED_BY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.APPROVED_BY = " & APPROVED_BY & ") "
      Else
         WhereStr = WhereStr & "AND (J.APPROVED_BY = " & APPROVED_BY & ") "
      End If
   End If
   
   If RESPONSE_BY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.RESPONSE_BY = " & RESPONSE_BY & ") "
      Else
         WhereStr = WhereStr & "AND (J.RESPONSE_BY = " & RESPONSE_BY & ") "
      End If
   End If
   
  If START_DATE > -1 Then
      TempDate = Trim(DateToStringInt(START_DATE)) 'START_DATE
      If WhereStr = "" Then
         WhereStr = "WHERE (J.START_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (J.START_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
  If FINISH_DATE > -1 Then
      TempDate = Trim(DateToStringInt(FINISH_DATE)) 'FINISH_DATE
      If WhereStr = "" Then
         WhereStr = "WHERE (J.FINISH_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (J.FINISH_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.SUCCESS_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (J.SUCCESS_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(VERIFY_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.VERIFY_FLAG = '" & VERIFY_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (J.VERIFY_FLAG = '" & VERIFY_FLAG & "') "
      End If
   End If
   
   If JOB_DOC_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_DOC_TYPE = " & JOB_DOC_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_DOC_TYPE = " & JOB_DOC_TYPE & ") "
      End If
   End If
   
   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM JOB_INOUT DI LEFT OUTER JOIN PART_ITEM PI ON (DI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (DI.TX_TYPE = 'I') AND (DI.JOB_ID = J.JOB_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')) > 0)"
         End If
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If Len(ProcessSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.PROCESS_ID IN " & ProcessSet & ") "
      Else
         WhereStr = WhereStr & "AND (J.PROCESS_ID IN " & ProcessSet & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY J.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY JOB_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY JOB_DATE "
         ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY  BATCH_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY JOB_DESC "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY START_DATE "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY FINISH_DATE "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY APPROVED_BY "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY RESPONSE_BY "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY PROCESS_NO "
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY DOC_NO "
      Else
         OrderByStr = "ORDER BY JOB_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PART_NO "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PART_NO "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      JOB_ID_REF = NVLI(Rs("JOB_ID_REF"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      JOB_DESC = NVLS(Rs("JOB_DESC"), "")
      PROCESS_ID = NVLI(Rs("PROCESS_ID"), -1)
      APPROVED_BY = NVLI(Rs("APPROVED_BY"), -1)
      RESPONSE_BY = NVLI(Rs("RESPONSE_BY"), -1)
      START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
      FINISH_DATE = InternalDateToDate(NVLS(Rs("FINISH_DATE"), ""))
      COMMIT_FLAG = NVLS(Rs("SUCCESS_FLAG"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      INVENTORY_WH_DOC_ID_INPUT = NVLI(Rs("INVENTORY_WH_DOC_ID_INPUT"), -1)
      BATCH_NO = NVLS(Rs("BATCH_NO"), "")
      FROM_BATCH_NO = NVLI(Rs("FROM_BATCH_NO"), -1)
      TO_BATCH_NO = NVLI(Rs("TO_BATCH_NO"), -1)
      BATCH_TOTAL = NVLD(Rs("BATCH_TOTAL"), 0)
      BATCH_DETAIL = NVLS(Rs("BATCH_DETAIL"), "")
      LONG_NAMEA = NVLS(Rs("LONGA"), "")
      LAST_NAMEA = NVLS(Rs("LASTA"), "")
      LONG_NAMER = NVLS(Rs("LONGR"), "")
      LAST_NAMER = NVLS(Rs("LASTR"), "")
      DOC_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PROCESS_NAME = NVLS(Rs("PROCESS_NAME"), "")
      JOB_DOC_TYPE = NVLI(Rs("JOB_DOC_TYPE"), -1)
      OLD_COMMIT_FLAG = COMMIT_FLAG
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
      FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
      SUM_INPUT = NVLD(Rs("SUM_INPUT"), 0)
      SUM_OUTPUT = NVLD(Rs("SUM_OUTPUT"), 0)
      LOCK_DOC_FLAG = NVLS(Rs("LOCK_DOC_FLAG"), "")
      CREATE_BY = NVLI(Rs("CREATE_BY"), -1)
      MODIFY_BY = NVLI(Rs("MODIFY_BY"), -1)
      
      VERIFY_FLAG = NVLS(Rs("VERIFY_FLAG"), "")
      VERIFY_NAME = NVLS(Rs("VERIFY_NAME"), "")
      
      NOTE = NVLS(Rs("NOTE"), "")
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      STD_AMOUNT = NVLD(Rs("STD_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), -1)
      
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 5 Then
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      JOB_DATE = InternalDateToDate(NVLS(Rs("JOB_DATE"), ""))
      FROM_BATCH_NO = NVLI(Rs("FROM_BATCH_NO"), -1)
      TO_BATCH_NO = NVLI(Rs("TO_BATCH_NO"), -1)
      BATCH_DETAIL = NVLS(Rs("BATCH_DETAIL"), "")
      LOCK_DOC_FLAG = NVLS(Rs("LOCK_DOC_FLAG"), "")
   ElseIf Ind = 6 Then
      JOB_ID = NVLI(Rs("JOB_ID"), -1)
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      LOT_ID = NVLI(Rs("LOT_ID"), -1)
      LOT_NO = NVLS(Rs("LOT_NO"), "")
      BIN_NO = NVLI(Rs("BIN_NO"), -1)
      BIN_NAME = NVLS(Rs("BIN_NAME"), "")
      START_DATE = InternalDateToDate(NVLS(Rs("START_DATE"), ""))
      FROM_BATCH_NO = NVLI(Rs("FROM_BATCH_NO"), -1)
      TO_BATCH_NO = NVLI(Rs("TO_BATCH_NO"), -1)
      BATCH_DETAIL = NVLS(Rs("BATCH_DETAIL"), "")
    ElseIf Ind = 7 Then
      INVENTORY_WH_DOC_ID = NVLI(Rs("INVENTORY_WH_DOC_ID"), -1)
      VERIFY_FLAG = NVLS(Rs("VERIFY_FLAG"), "")
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
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String
Dim TempID8 As String
Dim TempID9 As String

   TempID1 = PROCESS_ID
   If PROCESS_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = APPROVED_BY
   If APPROVED_BY <= 0 Then
      TempID2 = "NULL "
   End If
   
   TempID3 = RESPONSE_BY
   If RESPONSE_BY <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID4 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID4 = "NULL "
   End If
   
   TempID5 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID5 = "NULL "
   End If

   TempID6 = FORMULA_ID
   If FORMULA_ID <= 0 Then
      TempID6 = "NULL "
   End If
   
   TempID7 = INVENTORY_WH_DOC_ID
   If INVENTORY_WH_DOC_ID <= 0 Then
      TempID7 = "NULL "
   End If
   
   TempID8 = INVENTORY_WH_DOC_ID_INPUT
   If INVENTORY_WH_DOC_ID_INPUT <= 0 Then
      TempID8 = "NULL "
   End If
   
   TempID9 = JOB_ID_REF
   If JOB_ID_REF <= 0 Then
      TempID9 = "NULL "
   End If
   
   If Len(VERIFY_FLAG) = 0 Then 'ถ้าเป็นการ เพิ่มข้อมูลที่ไม่ใช้ Bag ยอมให้ปล่อยผ่านไปก่อน
      VERIFY_FLAG = "Y"
   End If

   
   YYYYMM = Mid(DateToStringInt(JOB_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO JOB " & _
                     "( " & _
                     "   JOB_ID, " & _
                     "   JOB_NO, " & _
                     "   JOB_DATE, " & _
                     "   JOB_DESC, " & _
                     "   BATCH_NO, " & _
                     "   PROCESS_ID, " & _
                     "   APPROVED_BY, " & _
                     "   RESPONSE_BY, " & _
                    "   START_DATE, " & _
                     "   FINISH_DATE, " & _
                     "   SUCCESS_FLAG, " & _
                        "   INVENTORY_DOC_ID, " & _
                        "   JOB_DOC_TYPE, " & _
                        "   PART_ITEM_ID, "
            SQL1 = SQL1 & _
                     "   ACTUAL_AMOUNT, " & _
                     "   STD_AMOUNT, " & _
                     "   FORMULA_ID, " & _
                     "   YYYYMM, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY, " & _
                    "    INVENTORY_WH_DOC_ID," & _
                    "    INVENTORY_WH_DOC_ID_INPUT, " & _
                    "    FROM_BATCH_NO," & _
                    "    TO_BATCH_NO," & _
                    "    BATCH_TOTAL," & _
                    "    BATCH_DETAIL,JOB_ID_REF,VERIFY_NAME,VERIFY_FLAG,NOTE " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(JOB_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "'," & _
                        "'" & ChangeQuote(Trim(JOB_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(BATCH_NO)) & "'," & _
                        TempID1 & "," & _
                         TempID2 & "," & _
                          TempID3 & "," & _
                          "'" & ChangeQuote(Trim(DateToStringInt(START_DATE))) & "'," & _
                          "'" & ChangeQuote(Trim(DateToStringInt(FINISH_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                        TempID4 & "," & _
                        JOB_DOC_TYPE & "," & _
                        TempID5 & ","
               SQL1 = SQL1 & _
                        ACTUAL_AMOUNT & "," & _
                        STD_AMOUNT & "," & _
                        TempID6 & "," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         TempID7 & "," & _
                         TempID8 & "," & _
                         FROM_BATCH_NO & "," & _
                         TO_BATCH_NO & "," & BATCH_TOTAL & ",'" & BATCH_DETAIL & "'," & TempID9 & ",'" & VERIFY_NAME & "','" & VERIFY_FLAG & "','" & NOTE & "') "
         JOB_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE JOB SET " & _
                     "   JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "', " & _
                     "   JOB_DATE= '" & ChangeQuote(Trim(DateToStringInt(JOB_DATE))) & "', " & _
                      "  JOB_DESC = '" & ChangeQuote(Trim(JOB_DESC)) & "', " & _
                     "   BATCH_NO= '" & ChangeQuote(Trim(BATCH_NO)) & "', " & _
                     "   FROM_BATCH_NO= '" & FROM_BATCH_NO & "', " & _
                     "   TO_BATCH_NO= '" & TO_BATCH_NO & "', BATCH_TOTAL= '" & BATCH_TOTAL & "',BATCH_DETAIL= '" & BATCH_DETAIL & "', " & _
                    "   PROCESS_ID = " & TempID1 & "," & _
                     "   APPROVED_BY = " & TempID2 & "," & _
                      "   RESPONSE_BY = " & TempID3 & "," & _
                    "   START_DATE = '" & ChangeQuote(Trim(DateToStringInt(START_DATE))) & "', " & _
                    "   FINISH_DATE = '" & ChangeQuote(Trim(DateToStringInt(FINISH_DATE))) & "', " & _
                      "  SUCCESS_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   INVENTORY_DOC_ID = " & TempID4 & "," & _
                     "   INVENTORY_WH_DOC_ID = " & TempID7 & "," & _
                     "   INVENTORY_WH_DOC_ID_INPUT = " & TempID8 & "," & _
                     "   JOB_DOC_TYPE = " & JOB_DOC_TYPE & "," & _
                     "   PART_ITEM_ID = " & TempID5 & "," & _
                     "   ACTUAL_AMOUNT = " & ACTUAL_AMOUNT & "," & _
                     "   STD_AMOUNT = " & STD_AMOUNT & "," & _
                     "   FORMULA_ID = " & TempID6 & ","
   SQL1 = SQL1 & "  YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    VERIFY_NAME = '" & VERIFY_NAME & "', " & _
                     "    VERIFY_FLAG = '" & VERIFY_FLAG & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & "," & _
                    "    NOTE = '" & NOTE & "'" & _
                     " WHERE JOB_ID = " & JOB_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM JOB_PARAMETER WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM JOB_INOUT WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)

     SQL1 = "DELETE FROM JOB_RESOURCE WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
     SQL1 = "DELETE FROM JOB_VERIFY WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM JOB WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Parameters = Nothing
   Set Verifies = Nothing
'   Set InventoryWhDoc = Nothing
'   Set InventoryWhDocInput = Nothing
End Sub
Public Function DeleteDataParameter()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (JOB_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = "DELETE FROM JOB_PARAMETER JP  WHERE JP.JOB_ID IN (SELECT UG.JOB_ID FROM JOB UG " & WhereStr & ")"
   m_Conn.Execute (SQL1)
      
End Function

Public Function UpdateLockDoc() As Boolean
On Error GoTo ErrorHandler
Dim SQL1 As String
Dim TempDate As String
Dim FromStr As String
Dim WhereStr As String
Dim SubLen As Long
Dim NewStr As String

   SQL1 = "UPDATE JOB J SET  J.LOCK_DOC_FLAG = '" & LOCK_DOC_FLAG & "' "
   
   If Trim(JOB_NO) <> "" Then
      If WildCard(JOB_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(JOB_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (J.JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (J.JOB_NO = '" & ChangeQuote(Trim(JOB_NO)) & "')"
         End If
      End If
   End If
   
   If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
   
   If JOB_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(JOB_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (J.JOB_DATE  = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (J.JOB_DATE  = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   SQL1 = SQL1 & WhereStr
   m_Conn.Execute (SQL1)
    UpdateLockDoc = True
   Exit Function
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   UpdateLockDoc = False
End Function
Public Function UpdateJobVerifyFlag(UserName As String) As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE JOB SET " & _
                  "   VERIFY_FLAG = 'Y'," & _
                  "   VERIFY_NAME = '" & UserName & "' " & _
                 "WHERE JOB_ID = " & JOB_ID

    m_Conn.Execute (SQL1)
End Function
Public Function UpdateJobCancelVerifyFlag(UserName As String) As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE JOB SET " & _
                  "   VERIFY_FLAG = 'N'," & _
                  "   VERIFY_NAME = 'CANCEL-" & UserName & "'" & _
                  "WHERE JOB_ID = " & JOB_ID

    m_Conn.Execute (SQL1)
    
End Function


