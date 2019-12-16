VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommissionCost"
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

Public COMMISSION_COST_ID As Long
Public MASTER_VALID_ID As Long

Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_DESC As String

Public COST_AMOUNT As Double

Public VALID_FROM As Date
Public VALID_TO As Date

Public VALID_FROM2 As Date
Public VALID_TO2 As Date


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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr  As String

   WhereStr = ""
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CMSC.*  "
      SelectStr1 = SelectStr1 & ",PI.PART_NO,PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM COMMISSION_COST CMSC "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CMSC.PART_ITEM_ID = PI.PART_ITEM_ID) "
   ElseIf Ind = 2 Then                       'LoadSumCommissionGp
      SelectStr1 = "SELECT CMSC.*,MTVD.VALID_FROM,MTVD.VALID_TO  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COMMISSION_COST CMSC "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_VALID MTVD ON (CMSC.MASTER_VALID_ID = MTVD.MASTER_VALID_ID) "
      
      WhereStr = " WHERE ((MTVD.VALID_FROM <=  '" & DateToStringIntLow(Trim(VALID_FROM2)) & "'" & " AND " & " MTVD.VALID_TO >= '" & DateToStringIntLow(Trim(VALID_FROM2)) & "')" & " OR " & "(MTVD.VALID_FROM <=  '" & DateToStringIntLow(Trim(VALID_TO2)) & "'" & " AND " & " MTVD.VALID_TO >= '" & DateToStringIntLow(Trim(VALID_TO2)) & "')" & " OR " & "(MTVD.VALID_FROM >=  '" & DateToStringIntLow(Trim(VALID_FROM2)) & "'" & " AND " & " MTVD.VALID_TO <= '" & DateToStringIntLow(Trim(VALID_TO2)) & "')" & ")"
   End If
   
   
   
   If COMMISSION_COST_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSC.COMMISSION_COST_ID = " & COMMISSION_COST_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSC.COMMISSION_COST_ID = " & COMMISSION_COST_ID & ")"
      End If
   End If
   
   If MASTER_VALID_ID > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (CMSC.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
      Else
         WhereStr = WhereStr & "AND (CMSC.MASTER_VALID_ID = " & MASTER_VALID_ID & ")"
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
         OrderByStr = "ORDER BY CMSC.COMMISSION_COST_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CMSC.COMMISSION_COST_ID "
      Else
         OrderByStr = "ORDER BY PI.PART_NO "
         OrderType = 1
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY MTVD.VALID_FROM "
      OrderType = 2
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
      COMMISSION_COST_ID = NVLI(Rs("COMMISSION_COST_ID"), -1)
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      COST_AMOUNT = NVLD(Rs("COST_AMOUNT"), 0)
      
   ElseIf Ind = 2 Then
      COMMISSION_COST_ID = NVLI(Rs("COMMISSION_COST_ID"), -1)
      
      MASTER_VALID_ID = NVLI(Rs("MASTER_VALID_ID"), -1)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      
      COST_AMOUNT = NVLD(Rs("COST_AMOUNT"), 0)
      
      VALID_FROM = InternalDateToDate(NVLS(Rs("VALID_FROM"), ""))
      VALID_TO = InternalDateToDate(NVLS(Rs("VALID_TO"), ""))
   End If
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO COMMISSION_COST " & _
                     "( " & _
                     "   MASTER_VALID_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   COST_AMOUNT, " & _
                     "   CREATE_DATE, " & " CREATE_NAME " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                     MASTER_VALID_ID & "," & _
                     TempID1 & ", " & _
                     COST_AMOUNT & ", " & _
                     "'" & ServerDateTime & "','" & glbUser.USER_NAME & "'" & _
                     ")  RETURNING COMMISSION_COST_ID "
   Else
            SQL1 = "UPDATE COMMISSION_COST SET " & _
                        "   PART_ITEM_ID = " & TempID1 & ", " & _
                        "   COST_AMOUNT = " & COST_AMOUNT & ", " & _
                        "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                        "    MODIFY_NAME = '" & glbUser.USER_NAME & "' " & _
                        " WHERE COMMISSION_COST_ID = " & COMMISSION_COST_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      COMMISSION_COST_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COMMISSION_COST WHERE COMMISSION_COST_ID = " & COMMISSION_COST_ID
      m_Conn.Execute (SQL1)
End Function
