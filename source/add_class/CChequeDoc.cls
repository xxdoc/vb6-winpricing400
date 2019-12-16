VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CChequeDoc"
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


Public CHEQUE_DOC_ID As Long
Public CHEQUE_DOC_NO As String
Public CHEQUE_DOC_DATE As Date
Public BADCHEQUE_FLAG As String
Public BADCHEQUE_DATE As Date
Public PASSCHEQUE_FLAG As String
Public PASSCHEQUE_DATE As Date
Public AMOUNT_CHEQUE As Double
Public BILLING_DOC_ID As Long

Public CUSTOMER_ID As Long
Public CUSTOMER_NO As String
Public DOCUMENT_DATE As Date

Public ORDER_BY As Long
Public ORDER_TYPE As Long
Public ACCOUNT_ID As Long


Public CUSTOMER_CODE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public LONG_NAME As String 'ชื่อลูกค้า
Public SUM_PAID_AMOUNT As Double
Public PAID_AMOUNT As Double
Public RECEIPT_CHEQUE_DOC_ID As Long
Public RECEIPT_CHEQUE_DOC_NO As String
Public CHEQUE_DOC_ID2 As Long
Public DOCUMENT_NO As String



Public CUSTOMER_TYPE As Long
Public CUSTOMER_GRADE As Long
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String
Public BANK_ID As Long
Public BANK_BRANCH_ID As Long
Public BANK_NAME As String
Public BANK_BRANCH_NAME As String




Private m_ChequeDoc_coll As Collection
Public ChequeDoc As Collection
Public ReceiptCheqDocItems As Collection
Public BankInfo As Collection

Private Sub Class_Initialize()
  Set m_Conn = glbDatabaseMngr.DBConnection
  Set m_ChequeDoc_coll = New Collection
  Set ChequeDoc = New Collection
  Set ReceiptCheqDocItems = New Collection
  Set BankInfo = New Collection
   FROM_DATE = -1
   TO_DATE = -1
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
Dim TempDate As String
Dim SubLen As Long
Dim GroupByStr As String
Dim NewStr As String


   If Ind = 1 Then
      SelectStr1 = "SELECT CD.*, CM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM  CHEQUE_DOC CD "
       FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CD.CUSTOMER_ID = CM.CUSTOMER_ID)"
       FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
       FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID)"
      

      WhereStr = ""
      
      
   ElseIf Ind = 2 Then
    
      SelectStr1 = "SELECT CD.*,CM.CUSTOMER_CODE,NM.LONG_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM  CHEQUE_DOC CD "
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CD.CUSTOMER_ID = CM.CUSTOMER_ID)"
       FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
       FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID)"

      WhereStr = ""
      
   ElseIf Ind = 3 Then
    
      SelectStr1 = "SELECT CD.*,AC.*   "
        SelectStr1 = SelectStr1 & "  , (SELECT SUM(RI.PAID_AMOUNT) FROM RECEIPT_ITEM RI WHERE RI.CHEQUE_DOC_ID = CD.CHEQUE_DOC_ID) SUM_PAID_AMOUNT   "

      SelectStr2 = "SELECT COUNT(*) "

    
      FromStr = "FROM  CHEQUE_DOC CD "
       FromStr = FromStr & "  LEFT OUTER JOIN  CUSTOMER CM ON (CM.CUSTOMER_ID = CD.CUSTOMER_ID)  "
       FromStr = FromStr & " LEFT OUTER JOIN ACCOUNT AC ON (CM.CUSTOMER_ID = AC.CUSTOMER_ID) "
    

'      WhereStr = "WHERE (CD.PASSCHEQUE_FLAG= " & PASSCHEQUE_FLAG & ") """
      WhereStr = ""
      
   ElseIf Ind = 4 Then
    SelectStr1 = "SELECT  CD.*,RCD.*   "
     SelectStr2 = "SELECT COUNT(*) "
     
      FromStr = "   FROM  CHEQUE_DOC CD "
      FromStr = FromStr & "LEFT OUTER JOIN  RECEIPT_CHEQUE_DOC RCD ON(RCD.CHEQUE_DOC_ID =CD.CHEQUE_DOC_ID) "
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = CD.CUSTOMER_ID)   "
        
       WhereStr = "WHERE  (((CD.PASSCHEQUE_FLAG <> 'N')  AND  (CD.BADCHEQUE_FLAG <> 'Y'))  OR" 'เช็คผ่านแล้ว
        WhereStr = WhereStr & "((CD.PASSCHEQUE_FLAG ='N')  AND  (CD.BADCHEQUE_FLAG = 'N')))" 'Key Cheque ไว้เฉยๆยังไม่ถึงวันเช็คเด้ง
    ElseIf Ind = 5 Then
    SelectStr1 = "SELECT  CD.*,RCD.*   "
     SelectStr2 = "SELECT COUNT(*) "
     
      FromStr = "   FROM  CHEQUE_DOC CD "
      FromStr = FromStr & "LEFT OUTER JOIN  RECEIPT_CHEQUE_DOC RCD ON(RCD.CHEQUE_DOC_ID =CD.CHEQUE_DOC_ID) "
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = CD.CUSTOMER_ID)   "
        
       WhereStr = "WHERE  (CD.PASSCHEQUE_FLAG <> 'Y')  AND  (CD.BADCHEQUE_FLAG <> 'N')" ' เช็คเด้ง
   
   ElseIf Ind = 6 Then
    SelectStr1 = "SELECT   CD.*, RT.*, CD.CHEQUE_DOC_ID CHEQUE_DOC_ID2,RT.PAID_AMOUNT,Bd.DOCUMENT_NO  DOCUMENT_NO"
     SelectStr2 = "SELECT COUNT(*) "
     
      FromStr = "   FROM  CHEQUE_DOC CD "
      FromStr = FromStr & "LEFT OUTER JOIN  RECEIPT_ITEM RT ON (RT.CHEQUE_DOC_ID = CD.CHEQUE_DOC_ID)"
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = CD.CUSTOMER_ID)   "
      FromStr = FromStr & " LEFT OUTER JOIN  BILLING_DOC BD ON (BD.BILLING_DOC_ID = RT.DO_ID) "
        
       WhereStr = " WHERE  (RT.CHEQUE_DOC_ID IS NOT NULL) "
       
       
   ElseIf Ind = 7 Then
    SelectStr1 = "SELECT  DOCUMENT_NO, SUM(RT.PAID_AMOUNT) SUM_PAID_AMOUNT ,CD.CUSTOMER_ID   "
     SelectStr2 = "SELECT COUNT(*) "
     
      FromStr = "   FROM  CHEQUE_DOC CD "
      FromStr = FromStr & "LEFT OUTER JOIN  RECEIPT_ITEM RT ON (RT.CHEQUE_DOC_ID = CD.CHEQUE_DOC_ID)"
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = CD.CUSTOMER_ID)   "
      FromStr = FromStr & " LEFT OUTER JOIN  BILLING_DOC BD ON (BD.BILLING_DOC_ID = RT.DO_ID) "
        
       WhereStr = " WHERE  (RT.CHEQUE_DOC_ID IS NOT NULL) "
       
   ElseIf Ind = 8 Then
    
      SelectStr1 = "SELECT CD.*,CM.CUSTOMER_CODE,NM.LONG_NAME "
       SelectStr1 = SelectStr1 & " , B.BANK_NAME    BANK_NAME  ,BB.BBRANCH_NAME    BANK_BRANCH_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM  CHEQUE_DOC CD "
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CD.CUSTOMER_ID = CM.CUSTOMER_ID)"
       FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
       FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID)"
       FromStr = FromStr & " LEFT OUTER JOIN BANK B ON (B.BANK_ID = CD.BANK_ID)"
       FromStr = FromStr & " LEFT OUTER JOIN BANK_BRANCH BB ON (BB.BBRANCH_ID = CD.BANK_BRANCH_ID)"

      WhereStr = ""
   ElseIf Ind = 9 Then 'CReportAR004_10
    SelectStr1 = "SELECT  CD.*,RCD.*   "
        SelectStr1 = SelectStr1 & " , B.BANK_NAME    BANK_NAME  ,BB.BBRANCH_NAME    BANK_BRANCH_NAME "
     SelectStr2 = "SELECT COUNT(*) "
     
      FromStr = "   FROM  CHEQUE_DOC CD "
      FromStr = FromStr & "LEFT OUTER JOIN  RECEIPT_CHEQUE_DOC RCD ON(RCD.CHEQUE_DOC_ID =CD.CHEQUE_DOC_ID) "
      FromStr = FromStr & " LEFT OUTER JOIN CUSTOMER CM ON (CM.CUSTOMER_ID = CD.CUSTOMER_ID)   "
      FromStr = FromStr & " LEFT OUTER JOIN BANK B ON (B.BANK_ID = CD.BANK_ID)   "
      FromStr = FromStr & " LEFT OUTER JOIN BANK_BRANCH BB ON (BB.BBRANCH_ID = CD.BANK_BRANCH_ID)   "
        
       WhereStr = "WHERE  (((CD.PASSCHEQUE_FLAG <> 'N')  AND  (CD.BADCHEQUE_FLAG <> 'Y'))  OR" 'เช็คผ่านแล้ว
        WhereStr = WhereStr & "((CD.PASSCHEQUE_FLAG ='N')  AND  (CD.BADCHEQUE_FLAG = 'N')))" 'Key Cheque ไว้เฉยๆยังไม่ถึงวันเช็คเด้ง
      
   End If

If CHEQUE_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CD.CHEQUE_DOC_ID= " & CHEQUE_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CD.CHEQUE_DOC_ID  = " & CHEQUE_DOC_ID & ") "
      End If
   End If

   
      If Val(ACCOUNT_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AC.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AC.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
   
   
     If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CD.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CD.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   


If Len(Trim(CUSTOMER_CODE)) > 0 Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
  If Len(Trim(FROM_CUSTOMER_CODE)) > 0 Then
      If WildCard(FROM_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") >= '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE >= '" & ChangeQuote(Trim(FROM_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(TO_CUSTOMER_CODE)) > 0 Then
      If WildCard(TO_CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") <= '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE <= '" & ChangeQuote(Trim(TO_CUSTOMER_CODE)) & "')"
         End If
      End If
   End If


   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If CUSTOMER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      End If
   End If



   If Len(Trim(CHEQUE_DOC_NO)) > 0 Then
      If WildCard(CHEQUE_DOC_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CHEQUE_DOC_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CHEQUE_DOC_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CHEQUE_DOC_NO = '" & ChangeQuote(Trim(CHEQUE_DOC_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CHEQUE_DOC_NO = '" & ChangeQuote(Trim(CHEQUE_DOC_NO)) & "')"
         End If
      End If
   End If
   
   
   
         
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CHEQUE_DOC_DATE  >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CHEQUE_DOC_DATE  >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CHEQUE_DOC_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CHEQUE_DOC_DATE  <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   
     If Len(PASSCHEQUE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CD.PASSCHEQUE_FLAG = '" & PASSCHEQUE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CD.PASSCHEQUE_FLAG = '" & PASSCHEQUE_FLAG & "') "
      End If
   End If
   
      If Len(BADCHEQUE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CD.BADCHEQUE_FLAG = '" & BADCHEQUE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CD.BADCHEQUE_FLAG = '" & BADCHEQUE_FLAG & "') "
      End If
   End If
   
  OrderByStr = ""
   OrderTypeStr = ""
   
   
   GroupByStr = ""
   If Ind = 7 Then
      GroupByStr = "GROUP BY   DOCUMENT_NO,CD.CUSTOMER_ID    "
   End If
   
   
   
 If Ind = 1 Then
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   Else
      OrderByStr = "ORDER BY CD.CHEQUE_DOC_ID"
   End If
   
   
   ElseIf Ind = 2 Then

    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_NO"
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY CD.CHEQUE_DOC_DATE"
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY CM.CUSTOMER_CODE"
   Else
   OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   End If
   
   
   ElseIf Ind = 3 Then
    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_DATE"
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_DATE"
   End If
   
   
   ElseIf Ind = 4 Then
    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_DATE"
   Else
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   End If
   
   ElseIf Ind = 5 Then
    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_DATE"
   Else
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   End If
   
   
   ElseIf Ind = 6 Then
    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_DATE"
   Else
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   End If
   
   
   ElseIf Ind = 7 Then
     OrderByStr = " ORDER BY   DOCUMENT_NO , CD.CUSTOMER_ID  "
     
    ElseIf Ind = 8 Then

    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_NO"
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY CD.CHEQUE_DOC_DATE"
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY CM.CUSTOMER_CODE"
   Else
   OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   End If
   
   ElseIf Ind = 9 Then
    If OrderBy = 1 Then
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_DATE"
   Else
      OrderByStr = "ORDER BY  CD.CHEQUE_DOC_ID"
   End If
   End If

 
  If OrderType = 1 Then
          OrderByStr = OrderByStr & " ASC "
   Else
           OrderByStr = OrderByStr & "   DESC  "
   End If
   
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
  If Ind = 1 Then
   CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
   CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
   CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
   BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
    PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
    CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
    LONG_NAME = NVLS(Rs("LONG_NAME"), "")
    AMOUNT_CHEQUE = NVLD(Rs("AMOUNT_CHEQUE"), 0)
    ElseIf Ind = 2 Then
    CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
    CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
    CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
    BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
     PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
  ElseIf Ind = 3 Then
   CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
   CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
   CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
   BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
    PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
'    CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
'    LONG_NAME = NVLS(Rs("LONG_NAME"), "")
    AMOUNT_CHEQUE = NVLD(Rs("AMOUNT_CHEQUE"), 0)
    SUM_PAID_AMOUNT = NVLD(Rs("SUM_PAID_AMOUNT"), 0)
    
     ElseIf Ind = 4 Then
    CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
    CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
    CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
    BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
     PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
    RECEIPT_CHEQUE_DOC_ID = NVLI(Rs("RECEIPT_CHEQUE_DOC_ID"), -1)
    RECEIPT_CHEQUE_DOC_NO = NVLS(Rs("RECEIPT_CHEQUE_DOC_NO"), "")
    PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
    PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    
    ElseIf Ind = 5 Then
    CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
    CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
    CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
    BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
     PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
    RECEIPT_CHEQUE_DOC_ID = NVLI(Rs("RECEIPT_CHEQUE_DOC_ID"), -1)
    RECEIPT_CHEQUE_DOC_NO = NVLS(Rs("RECEIPT_CHEQUE_DOC_NO"), "")
    PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
    
        
     ElseIf Ind = 6 Then
    CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
    CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
    CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
    BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
     PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
    CHEQUE_DOC_ID2 = NVLI(Rs("CHEQUE_DOC_ID2"), -1)
    DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
    PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
    
  ElseIf Ind = 7 Then

   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
    DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
    SUM_PAID_AMOUNT = NVLD(Rs("SUM_PAID_AMOUNT"), 0)
    
    ElseIf Ind = 8 Then
    CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
    CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
    CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
    BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
     PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
     BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BANK_BRANCH_ID = NVLI(Rs("BANK_BRANCH_ID"), -1)
     BANK_NAME = NVLS(Rs("BANK_NAME"), "")
     BANK_BRANCH_NAME = NVLS(Rs("BANK_BRANCH_NAME"), "")
    AMOUNT_CHEQUE = NVLD(Rs("AMOUNT_CHEQUE"), 0)
    
    ElseIf Ind = 9 Then
    CHEQUE_DOC_ID = NVLI(Rs("CHEQUE_DOC_ID"), -1)
    CHEQUE_DOC_NO = NVLS(Rs("CHEQUE_DOC_NO"), "")
    CHEQUE_DOC_DATE = InternalDateToDate(NVLS(Rs("CHEQUE_DOC_DATE"), ""))
    BADCHEQUE_FLAG = NVLS(Rs("BADCHEQUE_FLAG"), "")
     PASSCHEQUE_FLAG = NVLS(Rs("PASSCHEQUE_FLAG"), "")
    BADCHEQUE_DATE = InternalDateToDate(NVLS(Rs("BADCHEQUE_DATE"), ""))
     PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
    CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
    RECEIPT_CHEQUE_DOC_ID = NVLI(Rs("RECEIPT_CHEQUE_DOC_ID"), -1)
    RECEIPT_CHEQUE_DOC_NO = NVLS(Rs("RECEIPT_CHEQUE_DOC_NO"), "")
    PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
    PASSCHEQUE_DATE = InternalDateToDate(NVLS(Rs("PASSCHEQUE_DATE"), ""))
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
     BANK_BRANCH_NAME = NVLS(Rs("BANK_BRANCH_NAME"), "")
    End If
   
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Diff As Double
Dim TempID0 As String
Dim TempID1 As String
Dim TempID2 As String



TempID0 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID0 = "NULL"
   End If
   

TempID1 = BANK_ID
   If BANK_ID <= 0 Then
      TempID1 = "NULL"
   End If

TempID2 = BANK_BRANCH_ID
   If BANK_BRANCH_ID <= 0 Then
      TempID2 = "NULL"
   End If

   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CHEQUE_DOC_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO  CHEQUE_DOC" & _
                     "( " & _
                     "   CHEQUE_DOC_ID, " & _
                     "   CHEQUE_DOC_NO, " & _
                     "  CHEQUE_DOC_DATE, " & _
                     "   BADCHEQUE_FLAG, " & _
                    "   PASSCHEQUE_FLAG, " & _
                    "   BADCHEQUE_DATE, " & _
                    "   PASSCHEQUE_DATE, " & _
                    "   CUSTOMER_ID, " & _
                     "   AMOUNT_CHEQUE, " & _
                     "   BANK_ID, " & _
                    "   BANK_BRANCH_ID, " & _
                    "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
               SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                          "'" & ChangeQuote(Trim(CHEQUE_DOC_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(CHEQUE_DOC_DATE))) & "'," & _
                          "'" & BADCHEQUE_FLAG & "'," & _
                          "'" & PASSCHEQUE_FLAG & "'," & _
                          "'" & ChangeQuote(Trim(DateToStringInt(BADCHEQUE_DATE))) & "'," & _
                          "'" & ChangeQuote(Trim(DateToStringInt(PASSCHEQUE_DATE))) & "'," & _
                          "" & TempID0 & "," & _
                          "" & AMOUNT_CHEQUE & "," & _
                        "" & TempID1 & "," & _
                        "" & TempID2 & "," & _
                          "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CHEQUE_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CHEQUE_DOC  SET " & _
                     "   CHEQUE_DOC_NO = '" & ChangeQuote(Trim(CHEQUE_DOC_NO)) & "', " & _
                     "   CHEQUE_DOC_DATE = '" & ChangeQuote(Trim(DateToStringInt(CHEQUE_DOC_DATE))) & "', " & _
                      "   BADCHEQUE_FLAG  = '" & BADCHEQUE_FLAG & "', " & _
                     "   PASSCHEQUE_FLAG  = '" & PASSCHEQUE_FLAG & "', " & _
                     "   BADCHEQUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(BADCHEQUE_DATE))) & "', " & _
                     "   PASSCHEQUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(PASSCHEQUE_DATE))) & "', " & _
                     "   CUSTOMER_ID = " & TempID0 & ", " & _
                     "  AMOUNT_CHEQUE = " & AMOUNT_CHEQUE & ", " & _
                     "   BANK_ID = " & TempID1 & ", " & _
                     "   BANK_BRANCH_ID = " & TempID2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                   " WHERE CHEQUE_DOC_ID = " & CHEQUE_DOC_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

       SQL1 = "DELETE FROM RECEIPT_CHEQUE_DOC  WHERE CHEQUE_DOC_ID = " & CHEQUE_DOC_ID
      m_Conn.Execute (SQL1)
      
'       SQL1 = "DELETE FROM RECEIPT_ITEM  WHERE CHEQUE_DOC_ID = " & CHEQUE_DOC_ID
'      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CHEQUE_DOC WHERE CHEQUE_DOC_ID = " & CHEQUE_DOC_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
     Set ChequeDoc = Nothing
     Set ReceiptCheqDocItems = Nothing
     Set BankInfo = Nothing
     
End Sub
Public Function UpdateChequeStatus() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

    Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      If AddEditMode = SHOW_ADD Then
      Else 'Show edit
          SQL1 = " UPDATE CHEQUE_DOC SET" & _
                         "   BADCHEQUE_FLAG  = '" & BADCHEQUE_FLAG & "', " & _
                         "   PASSCHEQUE_FLAG  = '" & PASSCHEQUE_FLAG & "', " & _
                         "   BADCHEQUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(BADCHEQUE_DATE))) & "', " & _
                         "   PASSCHEQUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(PASSCHEQUE_DATE))) & "', " & _
                          "MODIFY_DATE = '" & ServerDateTime & "', " & _
                          "MODIFY_BY = " & glbUser.USER_ID & " " & _
                          "WHERE CHEQUE_DOC_ID = " & CHEQUE_DOC_ID
        
      End If
      
      If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
         m_Conn.Execute (SQL1)
      End If


End Function
