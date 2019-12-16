VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionCredit"
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
   
Public COMMISSION_CREDIT_ID As Long

Public CUSTOMER_ID As Long
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String

Public COMMISSION_CREDIT_AMOUNT As Long
Public COMMISSION_CREDIT_DESC As String

Public OrderBy  As Long
Public OrderType As Long
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

   WhereStr = ""
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CMSCD.*  "
      SelectStr1 = SelectStr1 & ",CM.CUSTOMER_CODE, NM2.LONG_NAME CUSTOMER_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_CREDIT CMSCD "
      
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (CMSCD.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM2 ON (CN.NAME_ID = NM2.NAME_ID) "
   ElseIf Ind = 2 Then  'LoadCommissionCredit
      SelectStr1 = "SELECT CMSCD.*  "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_CREDIT CMSCD "
   End If
   
   
   
   If COMMISSION_CREDIT_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSCD.COMMISSION_CREDIT_ID = " & COMMISSION_CREDIT_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSCD.COMMISSION_CREDIT_ID = " & COMMISSION_CREDIT_ID & ")"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CMSCD.COMMISSION_CREDIT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CMSCD.COMMISSION_CREDIT_ID "
      Else
         OrderByStr = "ORDER BY CMSCD.COMMISSION_CREDIT_ID "
         OrderType = 2
      End If
   End If
   
   If Len(OrderByStr) > 0 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COMMISSION_CREDIT_ID = NVLI(Rs("COMMISSION_CREDIT_ID"), -1)
      
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      COMMISSION_CREDIT_AMOUNT = NVLI(Rs("COMMISSION_CREDIT_AMOUNT"), 0)
      COMMISSION_CREDIT_DESC = NVLS(Rs("COMMISSION_CREDIT_DESC"), "")
   ElseIf Ind = 2 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      
      COMMISSION_CREDIT_AMOUNT = NVLI(Rs("COMMISSION_CREDIT_AMOUNT"), 0)
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
   
   TempID1 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COMMISSION_CREDIT " & _
                     "( " & _
                     "   CUSTOMER_ID, " & _
                     "   COMMISSION_CREDIT_AMOUNT, " & _
                     "   COMMISSION_CREDIT_DESC, " & _
                     "   CREATE_DATE, " & " CREATE_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     TempID1 & ", " & _
                     COMMISSION_CREDIT_AMOUNT & ", " & _
                     "'" & ChangeQuote(Trim(COMMISSION_CREDIT_DESC)) & "'," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING COMMISSION_CREDIT_ID "
   Else
            SQL1 = "UPDATE COMMISSION_CREDIT SET " & _
                        "   CUSTOMER_ID = " & TempID1 & ", " & _
                        "   COMMISSION_CREDIT_AMOUNT = " & COMMISSION_CREDIT_AMOUNT & ", " & _
                        "   COMMISSION_CREDIT_DESC = '" & ChangeQuote(Trim(COMMISSION_CREDIT_DESC)) & "', " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_NAME = '" & glbUser.USER_NAME & "' " & _
                        " WHERE COMMISSION_CREDIT_ID = " & COMMISSION_CREDIT_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      COMMISSION_CREDIT_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_CREDIT WHERE COMMISSION_CREDIT_ID = " & COMMISSION_CREDIT_ID
      m_Conn.Execute (SQL1)
End Function
