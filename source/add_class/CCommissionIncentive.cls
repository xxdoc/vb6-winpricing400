VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionIncentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public Flag As String
' End of generic data member

Public INCENTIVE_ID As Long
Public INCENTIVE_PER_PACK As Double

Public FREELANCE_ID As Long
Public FREELANCE_NAME As String
Public FREELANCE_LASTNAME As String
Public FREELANCE_CODE As String
Public FROM_FREELANCE_CODE As String
Public TO_FREELANCE_CODE As String

Public PART_ITEM_ID As Long
Public PART_ITEM_NAME As String
Public PART_ITEM_CODE As String
Public PART_NO As String

Public COMMISSION_SUBTRACT_ID As Long
Public MONTH_ID As Long
Public YEAR_NO As Long
Public YYYYMM As String

Public EMP_ID As Long
Public EMP_CODE As String
Public EMP_NAME As String

Public CUSTOMER_ID As Long
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public CUSTOMER_LASTNAME As String
Public DOCUMENT_TYPE As Long
Public UNIT_TYPE As Long
Public FROM_AMOUNT As Double
Public TO_AMOUNT As Double
Public AMOUNT_OVER_FLAG As String
Public RATE_OVER As Double


Public COMMISSION_SUBTRACT_AMOUNT As Double
Public COMMISSION_SUBTRACT_DESC As String


Public OrderBy  As Long
Public OrderType As Long
Public collCommissionSubTractSub As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Set collCommissionSubTractSub = New Collection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
Dim GroupByStr  As String

   WhereStr = ""
   
   If Ind = 1 Then
      SelectStr1 = "SELECT ICT.*,EP.EMP_CODE  "
      SelectStr1 = SelectStr1 & ",FL.FREELANCE_NAME,FL.FREELANCE_LASTNAME, PT.PART_DESC PART_ITEM_NAME "
      SelectStr1 = SelectStr1 & ",FL.FREELANCE_CODE,PT.PART_NO PART_ITEM_CODE,CM.CUSTOMER_CODE,NM1.LONG_NAME CUSTOMER_NAME,NM1.LAST_NAME CUSTOMER_LASTNAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INCENTIVE ICT "
      FromStr = FromStr & "LEFT OUTER JOIN FREELANCE FL ON (ICT.FREELANCE_ID = FL.FREELANCE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PT ON (ICT.PART_ITEM_ID = PT.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (ICT.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (CN.NAME_ID = NM1.NAME_ID)"
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EP ON (FL.EMP_ID = EP.EMP_ID)"

   End If
   

   
   If INCENTIVE_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ICT.INCENTIVE_ID = " & INCENTIVE_ID & ")"
      Else
         WhereStr = WhereStr & "AND (ICT.INCENTIVE_ID = " & INCENTIVE_ID & ")"
      End If
   End If
   
   If Trim(FREELANCE_CODE) <> "" Then
      If WildCard(FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FREELANCE_CODE = '" & ChangeQuote(Trim(FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
    If Trim(FROM_FREELANCE_CODE) <> "" Then
      If WildCard(FROM_FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FROM_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FROM_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FREELANCE_CODE >= '" & ChangeQuote(Trim(FROM_FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FREELANCE_CODE >= '" & ChangeQuote(Trim(FROM_FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(TO_FREELANCE_CODE) <> "" Then
      If WildCard(TO_FREELANCE_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(TO_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(TO_FREELANCE_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FREELANCE_CODE <= '" & ChangeQuote(Trim(TO_FREELANCE_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (FREELANCE_CODE <= '" & ChangeQuote(Trim(TO_FREELANCE_CODE)) & "')"
         End If
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ICT.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ")"
      Else
         WhereStr = WhereStr & "AND (ICT.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ")"
      End If
   End If
'   DOCUMENT_TYPE
   
   If Trim(PART_NO) <> "" Then
      If WildCard(PART_ITEM_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & Trim(PART_NO) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & Trim(PART_NO) & "')"
         End If
      End If
   End If
   
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY FL.FREELANCE_NAME  "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PT.PART_NO   "
      Else
         OrderByStr = "ORDER BY ICT.FREELANCE_ID "
         OrderType = 2
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
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      
      INCENTIVE_ID = NVLI(Rs("INCENTIVE_ID"), -1)
      
      FREELANCE_ID = NVLI(Rs("FREELANCE_ID"), -1)
      FREELANCE_NAME = NVLS(Rs("FREELANCE_NAME"), "")
      FREELANCE_LASTNAME = NVLS(Rs("FREELANCE_LASTNAME"), "")
      FREELANCE_CODE = NVLS(Rs("FREELANCE_CODE"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_ITEM_NAME = NVLS(Rs("PART_ITEM_NAME"), "")
      PART_ITEM_CODE = NVLS(Rs("PART_ITEM_CODE"), "")
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      CUSTOMER_LASTNAME = NVLS(Rs("CUSTOMER_LASTNAME"), "")
      
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      UNIT_TYPE = NVLI(Rs("UNIT_TYPE"), -1)
      FROM_AMOUNT = NVLD(Rs("FROM_AMOUNT"), 0)
      TO_AMOUNT = NVLD(Rs("TO_AMOUNT"), 0)
      AMOUNT_OVER_FLAG = NVLS(Rs("AMOUNT_OVER_FLAG"), "")
      RATE_OVER = NVLD(Rs("RATE_OVER"), 0)
      
      INCENTIVE_PER_PACK = NVLD(Rs("INCENTIVE_PER_PACK"), 0)
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      
   End If
End Sub
Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = FREELANCE_ID
   If FREELANCE_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("INCENTIVE_SEQ", ID, glbErrorLog)
      Else
         ID = INCENTIVE_ID
      End If

      SQL1 = "INSERT INTO INCENTIVE " & _
                     "( " & _
                     "   INCENTIVE_ID, " & _
                     "   FREELANCE_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   INCENTIVE_PER_PACK, " & _
                     "   UNIT_TYPE, " & _
                     "   FROM_AMOUNT, " & _
                     "   TO_AMOUNT, " & _
                     "   AMOUNT_OVER_FLAG, " & _
                     "   RATE_OVER, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     ID & "," & _
                     TempID1 & ", " & _
                     TempID2 & ", " & _
                     TempID3 & ", " & _
                     DOCUMENT_TYPE & ", " & _
                     INCENTIVE_PER_PACK & ", " & _
                     UNIT_TYPE & ", " & _
                     FROM_AMOUNT & ", " & _
                     TO_AMOUNT & ", " & _
                     "'" & AMOUNT_OVER_FLAG & "'," & _
                     RATE_OVER & ", " & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & "," & _
                     "'" & ServerDateTime & "'," & _
                     glbUser.USER_ID & _
                     ")  "
            INCENTIVE_ID = ID
            

   Else
            SQL1 = "UPDATE INCENTIVE SET " & _
                        "   FREELANCE_ID = " & TempID1 & ", " & _
                        "   PART_ITEM_ID = " & TempID2 & ", " & _
                        "   CUSTOMER_ID = " & TempID3 & ", " & _
                        "   DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                        "   INCENTIVE_PER_PACK = " & INCENTIVE_PER_PACK & ", " & _
                        "   UNIT_TYPE = " & UNIT_TYPE & ", " & _
                        "   FROM_AMOUNT = " & FROM_AMOUNT & ", " & _
                        "   TO_AMOUNT = " & TO_AMOUNT & ", " & _
                        "    AMOUNT_OVER_FLAG = '" & AMOUNT_OVER_FLAG & "', " & _
                        "   RATE_OVER = " & RATE_OVER & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_BY = " & glbUser.USER_ID & "" & _
                        " WHERE INCENTIVE_ID = " & INCENTIVE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM INCENTIVE WHERE INCENTIVE_ID = " & INCENTIVE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set collCommissionSubTractSub = Nothing
End Sub
