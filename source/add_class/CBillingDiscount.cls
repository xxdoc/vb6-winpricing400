VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillingDiscount"
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

Public BILLING_DISCOUNT_ID As Long
Public BILLING_DOC_ID As Double
Public DISCOUNT_NAME As String
Public PART_ITEM_ID As Long
Public FEATURE_ID As Long
Public DSCN_PER_PACK As Double
Public DSCN_PER_WEIGHT As Double
Public DSCN_PER_MONEY As Double
Public DISCOUNT_AMOUNT As Double
Public DISCOUNT_TYPE As Long

Public PART_NO As String
Public PART_DESC As String
Public FEATURE_CODE As String
Public FEATURE_DESC As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public CUSTOMER_ID As Long
Public ACCOUNT_ID As Long

Public ITEM_DESC As String
Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
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
Dim GroupByStr As String
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim FromDate1 As String
Dim ToDate1 As String
Dim FromDate2 As String
Dim ToDate2 As String

   If Ind = 1 Then
      'SelectStr1 = "SELECT LT.*, PI.PART_NO, PI.PART_DESC PART_NAME, FT.FEATURE_CODE, FT.FEATURE_DESC  "
      SelectStr1 = "SELECT LT.*, PI.PART_NO, PI.PART_DESC, FT.FEATURE_CODE, FT.FEATURE_DESC  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DISCOUNT LT "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LT.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FEATURE FT ON (LT.FEATURE_ID = FT.FEATURE_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT LT.BILLING_DOC_ID, SUM(LT.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DISCOUNT LT "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (LT.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      If Len(FROM_CUSTOMER_CODE) > 0 Or Len(TO_CUSTOMER_CODE) > 0 Then
         FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      End If
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT CM.CUSTOMER_ID, SUM(LT.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DISCOUNT LT "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (LT.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT BD.ACCOUNT_ID, SUM(LT.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DISCOUNT LT "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (LT.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      
      WhereStr = ""
   End If
   
   If BILLING_DISCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILLING_DISCOUNT_ID = " & BILLING_DISCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BILLING_DISCOUNT_ID = " & BILLING_DISCOUNT_ID & ") "
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
      
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
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
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY LT.BILLING_DOC_ID "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY CM.CUSTOMER_ID "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY BD.ACCOUNT_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY BILLING_DISCOUNT_ID "
      Else
         OrderByStr = "ORDER BY BILLING_DISCOUNT_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY LT.BILLING_DOC_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY CM.CUSTOMER_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY BD.ACCOUNT_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr & GroupByStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BILLING_DISCOUNT_ID = NVLI(Rs("BILLING_DISCOUNT_ID"), -1)
      BILLING_DOC_ID = NVLD(Rs("BILLING_DOC_ID"), 0)
      DISCOUNT_NAME = NVLS(Rs("DISCOUNT_NAME"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      FEATURE_ID = NVLI(Rs("FEATURE_ID"), -1)
      DSCN_PER_PACK = NVLD(Rs("DSCN_PER_PACK"), 0)
      DSCN_PER_WEIGHT = NVLD(Rs("DSCN_PER_WEIGHT"), 0)
      DSCN_PER_MONEY = NVLD(Rs("DSCN_PER_MONEY"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      DISCOUNT_TYPE = NVLI(Rs("DISCOUNT_TYPE"), -1)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      FEATURE_CODE = NVLS(Rs("FEATURE_CODE"), "")
      FEATURE_DESC = NVLS(Rs("FEATURE_DESC"), "")
      
      If PART_ITEM_ID > 0 Then
         ITEM_DESC = PART_DESC & " (" & PART_NO & ")"
      ElseIf FEATURE_ID > 0 Then
         ITEM_DESC = FEATURE_DESC & " (" & FEATURE_CODE & ")"
      End If
   ElseIf Ind = 2 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = FEATURE_ID
   If FEATURE_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILLING_DISCOUNT_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO BILLING_DISCOUNT " & _
                     "( " & _
                     "   BILLING_DISCOUNT_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   FEATURE_ID, " & _
                     "   DSCN_PER_PACK, " & _
                     "   DSCN_PER_WEIGHT, " & _
                     "   DSCN_PER_MONEY, " & _
                     "   DISCOUNT_NAME, " & _
                     "   DISCOUNT_AMOUNT, " & _
                     "   DISCOUNT_TYPE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                         TempID2 & "," & _
                         TempID1 & "," & _
                         DSCN_PER_PACK & "," & _
                         DSCN_PER_WEIGHT & "," & _
                         DSCN_PER_MONEY & "," & _
                         "'" & ChangeQuote(Trim(DISCOUNT_NAME)) & "'," & _
                         DISCOUNT_AMOUNT & "," & _
                         DISCOUNT_TYPE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               BILLING_DISCOUNT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BILLING_DISCOUNT SET " & _
                     "   BILLING_DOC_ID = " & BILLING_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & TempID2 & ", " & _
                     "   FEATURE_ID = " & TempID1 & ", " & _
                     "   DSCN_PER_PACK = " & DSCN_PER_PACK & ", " & _
                     "   DSCN_PER_WEIGHT = " & DSCN_PER_WEIGHT & ", " & _
                     "   DSCN_PER_MONEY = " & DSCN_PER_MONEY & ", " & _
                     "   DISCOUNT_NAME = '" & ChangeQuote(Trim(DISCOUNT_NAME)) & "', " & _
                     "   DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                     "   DISCOUNT_TYPE = " & DISCOUNT_TYPE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BILLING_DISCOUNT_ID = " & BILLING_DISCOUNT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM BILLING_DISCOUNT WHERE BILLING_DISCOUNT_ID = " & BILLING_DISCOUNT_ID
      m_Conn.Execute (SQL1)
End Function



