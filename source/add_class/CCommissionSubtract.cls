VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionSubtract"
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

Public COMMISSION_SUBTRACT_AMOUNT As Double
Public COMMISSION_SUBTRACT_DESC As String

Public OrderBy  As Long
Public OrderType As Long
Public collCommissionSubTractSub As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Set collCommissionSubTractSub = New Collection
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

   WhereStr = ""
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CMSST.*  "
      SelectStr1 = SelectStr1 & ",CM.CUSTOMER_CODE, NM2.LONG_NAME CUSTOMER_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE,NM.LONG_NAME SALE_NAME, NM.LAST_NAME SALE_LNAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_SUBTRACT CMSST "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (CMSST.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EM.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = EN.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CMSST.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (CN.NAME_ID = NM2.NAME_ID) "
   End If
   
   
   
   If COMMISSION_SUBTRACT_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSST.COMMISSION_SUBTRACT_ID = " & COMMISSION_SUBTRACT_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSST.COMMISSION_SUBTRACT_ID = " & COMMISSION_SUBTRACT_ID & ")"
      End If
   End If
   
   If MONTH_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSST.MONTH_ID = " & MONTH_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSST.MONTH_ID = " & MONTH_ID & ")"
      End If
   End If
   
   If YEAR_NO > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSST.YEAR_NO = " & YEAR_NO & ")"
      Else
         WhereStr = WhereStr & "AND (CMSST.YEAR_NO = " & YEAR_NO & ")"
      End If
   End If
   
   If Len(YYYYMM) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CMSST.YYYYMM = '" & YYYYMM & "') "
      Else
         WhereStr = WhereStr & "AND (CMSST.YYYYMM = '" & YYYYMM & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CMSST.COMMISSION_SUBTRACT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CMSST.COMMISSION_SUBTRACT_ID "
      Else
         OrderByStr = "ORDER BY CMSST.YYYYMM DESC,CMSST.COMMISSION_SUBTRACT_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COMMISSION_SUBTRACT_ID = NVLI(Rs("COMMISSION_SUBTRACT_ID"), -1)
      
      MONTH_ID = NVLI(Rs("MONTH_ID"), -1)
      YEAR_NO = NVLI(Rs("YEAR_NO"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("SALE_NAME"), "") & " " & NVLS(Rs("SALE_LNAME"), "")
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      COMMISSION_SUBTRACT_AMOUNT = NVLD(Rs("COMMISSION_SUBTRACT_AMOUNT"), 0)
      COMMISSION_SUBTRACT_DESC = NVLS(Rs("COMMISSION_SUBTRACT_DESC"), "")
      
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = EMP_ID
   If EMP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   YYYYMM = YEAR_NO & "-" & Format(MONTH_ID, "00")
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COMMISSION_SUBTRACT " & _
                     "( " & _
                     "   MONTH_ID, " & _
                     "   YEAR_NO, " & _
                     "   YYYYMM, " & _
                     "   EMP_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   COMMISSION_SUBTRACT_AMOUNT, " & _
                     "   COMMISSION_SUBTRACT_DESC, " & _
                     "   CREATE_DATE, " & " CREATE_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     MONTH_ID & "," & _
                     YEAR_NO & "," & _
                     "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                     TempID1 & ", " & _
                     TempID2 & ", " & _
                     COMMISSION_SUBTRACT_AMOUNT & ", " & _
                     "'" & ChangeQuote(Trim(COMMISSION_SUBTRACT_DESC)) & "'," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING COMMISSION_SUBTRACT_ID "
   Else
            SQL1 = "UPDATE COMMISSION_SUBTRACT SET " & _
                        "   MONTH_ID = " & MONTH_ID & ", " & _
                        "   YEAR_NO = " & YEAR_NO & ", " & _
                        "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                        "   EMP_ID = " & TempID1 & ", " & _
                        "   CUSTOMER_ID = " & TempID2 & ", " & _
                        "   COMMISSION_SUBTRACT_AMOUNT = " & COMMISSION_SUBTRACT_AMOUNT & ", " & _
                        "   COMMISSION_SUBTRACT_DESC = '" & ChangeQuote(Trim(COMMISSION_SUBTRACT_DESC)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_NAME = '" & glbUser.USER_NAME & "' " & _
                        " WHERE COMMISSION_SUBTRACT_ID = " & COMMISSION_SUBTRACT_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      COMMISSION_SUBTRACT_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_SUBTRACT WHERE COMMISSION_SUBTRACT_ID = " & COMMISSION_SUBTRACT_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set collCommissionSubTractSub = Nothing
End Sub
