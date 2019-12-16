VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionBgChart"
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

Public COMMISSION_BUDGET_CHART_ID As Long
Public MASTER_VALID_ID As Long
Public PARENT_ID As Long
Public EMP_ID As String

Public OLD_PK As Long

Public EMP_CODE As String
Public EMP_NAME As String
Public EMP_LNAME As String

Public OrderBy  As Long
Public OrderType As Long

Public VALID_FROM As Date
Public VALID_TO As Date
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long, Optional MustCount As Boolean = False)
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
   
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CMSBGC.*,EM.EMP_CODE  "
      SelectStr1 = SelectStr1 & ",NM1.LONG_NAME EMP_NAME,NM1.LAST_NAME EMP_LNAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_BUDGET_CHART CMSBGC "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (CMSBGC.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EM.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (EN.NAME_ID = NM1.NAME_ID) "
   
   ElseIf Ind = 2 Then        'LoadCommissionChartValidFromTo        ,LoadCommissionChartValidFromToEmp
      SelectStr1 = "SELECT CMSBGC.*,EM.EMP_CODE  "
      SelectStr1 = SelectStr1 & ",NM1.LONG_NAME EMP_NAME,NM1.LAST_NAME EMP_LNAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_BUDGET_CHART CMSBGC "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_VALID MTVD ON (CMSBGC.MASTER_VALID_ID = MTVD.MASTER_VALID_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (CMSBGC.EMP_ID = EM.EMP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EM.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (EN.NAME_ID = NM1.NAME_ID) "
      
   End If
   WhereStr = ""
   
   If COMMISSION_BUDGET_CHART_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSBGC.COMMISSION_BUDGET_CHART_ID = " & COMMISSION_BUDGET_CHART_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSBGC.COMMISSION_BUDGET_CHART_ID = " & COMMISSION_BUDGET_CHART_ID & ")"
      End If
   End If
   
   If MASTER_VALID_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSBGC.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSBGC.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      End If
   End If
      
   If VALID_FROM > 0 Then
      TempDate = DateToStringIntLow(Trim(VALID_FROM))
      If WhereStr = "" Then
         WhereStr = " WHERE (MTVD.VALID_FROM <=  '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (MTVD.VALID_FROM<= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If VALID_TO > 0 Then
      TempDate = DateToStringIntLow(Trim(VALID_TO))
      If WhereStr = "" Then
         WhereStr = " WHERE (MTVD.VALID_TO>= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & " AND (MTVD.VALID_TO>= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CMSBGC.COMMISSION_BUDGET_CHART_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CMSBGC.COMMISSION_BUDGET_CHART_ID "
      Else
         OrderByStr = "ORDER BY CMSBGC.COMMISSION_BUDGET_CHART_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY CMSBGC.PARENT_ID DESC,EM.EMP_CODE "
      OrderType = 1
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COMMISSION_BUDGET_CHART_ID = NVLI(Rs("COMMISSION_BUDGET_CHART_ID"), -1)
      OLD_PK = COMMISSION_BUDGET_CHART_ID
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
   ElseIf Ind = 2 Then
      COMMISSION_BUDGET_CHART_ID = NVLI(Rs("COMMISSION_BUDGET_CHART_ID"), -1)
      OLD_PK = COMMISSION_BUDGET_CHART_ID
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
      
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String


   TempID1 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = EMP_ID
   If EMP_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COMMISSION_BUDGET_CHART " & _
                     "( " & _
                     "   MASTER_VALID_ID, " & _
                     "   PARENT_ID, " & _
                     "   EMP_ID, " & _
                     "   CREATE_DATE, " & " CREATE_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     MASTER_VALID_ID & "," & _
                     TempID1 & "," & _
                     TempID2 & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING COMMISSION_BUDGET_CHART_ID "
   Else
            SQL1 = "UPDATE COMMISSION_BUDGET_CHART SET " & _
                        "   PARENT_ID = " & TempID1 & ", " & _
                        "   EMP_ID = " & TempID2 & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_NAME = '" & glbUser.USER_NAME & "' " & _
                        " WHERE COMMISSION_BUDGET_CHART_ID = " & COMMISSION_BUDGET_CHART_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      COMMISSION_BUDGET_CHART_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_BUDGET_CHART WHERE COMMISSION_BUDGET_CHART_ID = " & COMMISSION_BUDGET_CHART_ID
      m_Conn.Execute (SQL1)
End Function
