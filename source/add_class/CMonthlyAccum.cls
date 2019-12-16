VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMonthlyAccum"
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

Public MONTHLY_ACCUM_ID As Long

Public YYYYMM  As String                  'เก็บเป็นเดือน

Public PART_ITEM_ID As Long
Public BALANCE_AMOUNT As Double
Public TOTAL_INCLUDE_PRICE As Double
Public LOCATION_ID As Long
Public AVG_PRICE As Double

Public FROM_YYYYMM As String
Public TO_YYYYMM As String

Public LOCATION_NO As String
Public LOCATION_NAME As String
Public PART_NO  As String
Public PART_DESC As String
Public PART_TYPE As Long
Public PART_TYPE_NO As String
Public PART_TYPE_NAME As String
Public PART_GROUP As Long
Public PARCEL_TYPE  As Long

Public FROM_PART_NO As String
Public TO_PART_NO As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, MustCount As Boolean)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then               ' Normal Use
      SelectStr1 = "SELECT MA.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM MA "
      
      WhereStr = ""
   ElseIf Ind = 2 Then        'LoadMonthlyBalance
      SelectStr1 = "SELECT MA.LOCATION_ID, MA.PART_ITEM_ID, MA.BALANCE_AMOUNT, MA.TOTAL_INCLUDE_PRICE, MA.AVG_PRICE "
      SelectStr1 = SelectStr1 & ", PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM MA "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MA.PART_ITEM_ID = PI.PART_ITEM_ID) "
      WhereStr = ""
   ElseIf Ind = 3 Then              'CReportInventory004_5,CReportInventory004_4
      SelectStr1 = "SELECT DISTINCT MA.LOCATION_ID, LC.LOCATION_NO, LC.LOCATION_NAME, MA.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",PT.PART_TYPE_NO,PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM MA "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MA.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (MA.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then              'CReportInventory015,CReportInventory004_3
      SelectStr1 = "SELECT DISTINCT MA.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM MA "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MA.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (MA.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 5 Then           'LoadMonthlyBalancePartItem
      SelectStr1 = "SELECT MA.PART_ITEM_ID,SUM(MA.BALANCE_AMOUNT) BALANCE_AMOUNT, SUM(MA.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM MA "
      WhereStr = ""
   ElseIf Ind = 6 Then              'CReportInventory030
      SelectStr1 = "SELECT DISTINCT MA.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM MA "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MA.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (MA.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   End If
   
   If MONTHLY_ACCUM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MA.MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MA.MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MA.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MA.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If PART_GROUP > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MA.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MA.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Len(YYYYMM) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MA.YYYYMM = '" & YYYYMM & "') "
      Else
         WhereStr = WhereStr & "AND (MA.YYYYMM = '" & YYYYMM & "') "
      End If
   End If
   
   If Len(Trim(FROM_PART_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (PI.PART_NO >= '" & ChangeQuote(Trim(FROM_PART_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (PI.PART_NO >= '" & ChangeQuote(Trim(FROM_PART_NO)) & "')"
      End If
   End If
   
   If Len(Trim(TO_PART_NO)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (PI.PART_NO <= '" & ChangeQuote(Trim(TO_PART_NO)) & "')"
      Else
         WhereStr = WhereStr & "AND (PI.PART_NO <= '" & ChangeQuote(Trim(TO_PART_NO)) & "')"
      End If
   End If
   
   If Ind = 5 Then
      GroupByStr = "GROUP BY MA.PART_ITEM_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY MA.MONTHLY_ACCUM_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC,LC.LOCATION_NO ASC,PI.PART_NO "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   End If
   
   If Len(OrderByStr) > 0 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      MONTHLY_ACCUM_ID = NVLI(Rs("MONTHLY_ACCUM_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 3 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 6 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), 0)
   End If
End Function
Public Function AddEditData()
On Error Resume Next
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MONTHLY_ACCUM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO MONTHLY_ACCUM " & _
                     "( " & _
                     "   MONTHLY_ACCUM_ID, " & _
                     "   YYYYMM, " & _
                     "   PART_ITEM_ID, " & _
                     "   BALANCE_AMOUNT, " & _
                     "   TOTAL_INCLUDE_PRICE, " & _
                     "   LOCATION_ID, " & _
                     "   AVG_PRICE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        PART_ITEM_ID & "," & _
                        BALANCE_AMOUNT & "," & _
                        TOTAL_INCLUDE_PRICE & "," & _
                        LOCATION_ID & "," & _
                        AVG_PRICE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         MONTHLY_ACCUM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MONTHLY_ACCUM SET " & _
                     "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    BALANCE_AMOUNT = " & BALANCE_AMOUNT & ", " & _
                     "    TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "    LOCATION_ID = " & LOCATION_ID & ", " & _
                     "    AVG_PRICE = " & AVG_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM MONTHLY_ACCUM WHERE MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID
      m_Conn.Execute (SQL1)
End Function
Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String

   SQL1 = "DELETE FROM MONTHLY_ACCUM MA "
   
   If Len(FROM_YYYYMM) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MA.YYYYMM >= '" & FROM_YYYYMM & "') "
      Else
         WhereStr = WhereStr & "AND (MA.YYYYMM >= '" & FROM_YYYYMM & "') "
      End If
   End If
   
   If Len(TO_YYYYMM) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MA.YYYYMM <= '" & TO_YYYYMM & "') "
      Else
         WhereStr = WhereStr & "AND (MA.YYYYMM <= '" & TO_YYYYMM & "') "
      End If
   End If
   
   SQL1 = SQL1 & WhereStr
   m_Conn.Execute (SQL1)
End Function


