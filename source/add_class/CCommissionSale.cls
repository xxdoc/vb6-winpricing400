VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionSale"
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

Public COMMISSION_SALE_ID As Long
Public MASTER_VALID_ID As Long
Public COMMISSION_SALE_AREA As Long       'AREA =1 คือ คอมได้ตอนเก็บเงิน         AREA = 2 คือคอมได้เงินหักเงินทันที           AREA = 3 คือคอมได้เงินหักเงินทันที MANAGER
Public COMMISSION_SALE_TYPE As Long       'TYPE =1 คือรับจ่ายเป็น % , 2 รับจ่ายเป็นบาท

Public SELL_FROM  As Double
Public SELL_TO As Double
Public COMMISSION_SALE_AMOUNT As Double

Public COMMISSION_SALE_TYPE_NAME As String

Public VALID_FROM As Date
Public VALID_TO As Date

Public OrderBy  As Long
Public OrderType As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CMSSALE.*  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM COMMISSION_SALE CMSSALE "
   ElseIf Ind = 2 Then                          'LoadMasterValidCommissionSale
      SelectStr1 = "SELECT CMSSALE.*,MTVD.VALID_FROM,MTVD.VALID_TO  "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM COMMISSION_SALE CMSSALE "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_VALID MTVD ON (CMSSALE.MASTER_VALID_ID = MTVD.MASTER_VALID_ID) "
      
   End If
   WhereStr = ""
   
   If COMMISSION_SALE_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSSALE.COMMISSION_SALE_ID = " & COMMISSION_SALE_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSSALE.COMMISSION_SALE_ID = " & COMMISSION_SALE_ID & ")"
      End If
   End If
   
   If MASTER_VALID_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSSALE.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSSALE.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CMSSALE.COMMISSION_SALE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY CMSSALE.COMMISSION_SALE_ID "
   Else
      OrderByStr = "ORDER BY CMSSALE.SELL_FROM "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COMMISSION_SALE_ID = NVLI(Rs("COMMISSION_SALE_ID"), -1)
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      COMMISSION_SALE_AREA = NVLI(Rs("COMMISSION_SALE_AREA"), -1)
      COMMISSION_SALE_TYPE = NVLI(Rs("COMMISSION_SALE_TYPE"), -1)
      
      SELL_FROM = NVLD(Rs("SELL_FROM"), 0)
      SELL_TO = NVLD(Rs("SELL_TO"), 0)
      
      COMMISSION_SALE_AMOUNT = NVLD(Rs("COMMISSION_SALE_AMOUNT"), 0)
   ElseIf Ind = 2 Then
      COMMISSION_SALE_ID = NVLI(Rs("COMMISSION_SALE_ID"), -1)
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      COMMISSION_SALE_AREA = NVLI(Rs("COMMISSION_SALE_AREA"), -1)
      COMMISSION_SALE_TYPE = NVLI(Rs("COMMISSION_SALE_TYPE"), -1)
      
      SELL_FROM = NVLD(Rs("SELL_FROM"), 0)
      SELL_TO = NVLD(Rs("SELL_TO"), 0)
      
      COMMISSION_SALE_AMOUNT = NVLD(Rs("COMMISSION_SALE_AMOUNT"), 0)
      
      VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
      VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
'Dim TempID1 As String

'   TempID1 = COMMISSION_SALE_AREA
'   If COMMISSION_SALE_AREA <= 0 Then
'      TempID1 = "NULL"
'   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COMMISSION_SALE " & _
                     "( " & _
                     "   MASTER_VALID_ID, " & _
                     "   COMMISSION_SALE_AREA, " & _
                     "   COMMISSION_SALE_TYPE, " & _
                     "   SELL_FROM, " & _
                     "   SELL_TO, " & _
                     "   COMMISSION_SALE_AMOUNT, " & _
                     "   CREATE_DATE, " & " CREATE_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     MASTER_VALID_ID & "," & _
                     COMMISSION_SALE_AREA & "," & _
                     COMMISSION_SALE_TYPE & "," & _
                     SELL_FROM & "," & _
                     SELL_TO & "," & _
                     COMMISSION_SALE_AMOUNT & "," & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING COMMISSION_SALE_ID "
   Else
            SQL1 = "UPDATE COMMISSION_SALE SET " & _
                        "   COMMISSION_SALE_AREA = " & COMMISSION_SALE_AREA & ", " & _
                        "   COMMISSION_SALE_TYPE = " & COMMISSION_SALE_TYPE & ", " & _
                        "   SELL_FROM = " & SELL_FROM & ", " & _
                        "   SELL_TO = " & SELL_TO & ", " & _
                        "   COMMISSION_SALE_AMOUNT = " & COMMISSION_SALE_AMOUNT & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_NAME = '" & glbUser.USER_NAME & "' " & _
                        " WHERE COMMISSION_SALE_ID = " & COMMISSION_SALE_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      COMMISSION_SALE_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_SALE WHERE COMMISSION_SALE_ID = " & COMMISSION_SALE_ID
      m_Conn.Execute (SQL1)
End Function
