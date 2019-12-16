VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTargetDetail"
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

Public TARGET_DETAIL_ID As Long
Public TARGET_ID As Long

Public EMP_ID As Long

Public TARGET_PRICE1 As Double
Public TARGET_PRICE2 As Double
Public TARGET_PRICE3 As Double
Public TARGET_PRICE4 As Double
Public TARGET_PRICE5 As Double
Public TARGET_PRICE6 As Double
Public TARGET_PRICE7 As Double
Public TARGET_PRICE8 As Double
Public TARGET_PRICE9 As Double
Public TARGET_PRICE10 As Double
Public TARGET_PRICE11 As Double
Public TARGET_PRICE12 As Double

Public EMP_CODE As String
Public EMP_NAME As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public YEAR_NO As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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
      SelectStr1 = "SELECT TGDT.* "
      SelectStr1 = SelectStr1 & ",EMP.EMP_CODE "
      SelectStr1 = SelectStr1 & ",NM.LONG_NAME, NM.LAST_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM TARGET_DETAIL TGDT "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EMP ON (TGDT.EMP_ID = EMP.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EMP.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = EN.NAME_ID) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT TGDT.*,TG.YEAR_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM TARGET_DETAIL TGDT "
      FromStr = FromStr & "LEFT OUTER JOIN TARGET TG ON (TGDT.TARGET_ID = TG.TARGET_ID) "
   End If
   
   If TARGET_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TGDT.TARGET_DETAIL_ID = " & TARGET_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TGDT.TARGET_DETAIL_ID = " & TARGET_DETAIL_ID & ") "
      End If
   End If
   
   If TARGET_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TGDT.TARGET_ID = " & TARGET_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TGDT.TARGET_ID = " & TARGET_ID & ") "
      End If
   End If
   
   If YEAR_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TG.YEAR_NO = " & YEAR_NO & ") "
      Else
         WhereStr = WhereStr & "AND (TG.YEAR_NO = " & YEAR_NO & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY TGDT.EMP_ID, EMP.EMP_CODE, EMP.EMP_NAME "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY TGDT.EMP_ID, EMP.EMP_CODE, EMP.EMP_NAME,EMP.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY EMP.EMP_CODE "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TARGET_DETAIL_ID = NVLI(Rs("TARGET_DETAIL_ID"), -1)
      TARGET_ID = NVLI(Rs("TARGET_ID"), -1)
      
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      
      TARGET_PRICE1 = NVLD(Rs("TARGET_PRICE1"), 0)
      TARGET_PRICE2 = NVLD(Rs("TARGET_PRICE2"), 0)
      TARGET_PRICE3 = NVLD(Rs("TARGET_PRICE3"), 0)
      TARGET_PRICE4 = NVLD(Rs("TARGET_PRICE4"), 0)
      TARGET_PRICE5 = NVLD(Rs("TARGET_PRICE5"), 0)
      TARGET_PRICE6 = NVLD(Rs("TARGET_PRICE6"), 0)
      TARGET_PRICE7 = NVLD(Rs("TARGET_PRICE7"), 0)
      TARGET_PRICE8 = NVLD(Rs("TARGET_PRICE8"), 0)
      TARGET_PRICE9 = NVLD(Rs("TARGET_PRICE9"), 0)
      TARGET_PRICE10 = NVLD(Rs("TARGET_PRICE10"), 0)
      TARGET_PRICE11 = NVLD(Rs("TARGET_PRICE11"), 0)
      TARGET_PRICE12 = NVLD(Rs("TARGET_PRICE12"), 0)
            
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      EMP_NAME = NVLS(Rs("LONG_NAME"), "") & " " & NVLS(Rs("LAST_NAME"), "")
   ElseIf Ind = 2 Then
      EMP_ID = NVLI(Rs("EMP_ID"), -1)
      
      YEAR_NO = NVLS(Rs("YEAR_NO"), "")
      
      TARGET_PRICE1 = NVLD(Rs("TARGET_PRICE1"), 0)
      TARGET_PRICE2 = NVLD(Rs("TARGET_PRICE2"), 0)
      TARGET_PRICE3 = NVLD(Rs("TARGET_PRICE3"), 0)
      TARGET_PRICE4 = NVLD(Rs("TARGET_PRICE4"), 0)
      TARGET_PRICE5 = NVLD(Rs("TARGET_PRICE5"), 0)
      TARGET_PRICE6 = NVLD(Rs("TARGET_PRICE6"), 0)
      TARGET_PRICE7 = NVLD(Rs("TARGET_PRICE7"), 0)
      TARGET_PRICE8 = NVLD(Rs("TARGET_PRICE8"), 0)
      TARGET_PRICE9 = NVLD(Rs("TARGET_PRICE9"), 0)
      TARGET_PRICE10 = NVLD(Rs("TARGET_PRICE10"), 0)
      TARGET_PRICE11 = NVLD(Rs("TARGET_PRICE11"), 0)
      TARGET_PRICE12 = NVLD(Rs("TARGET_PRICE12"), 0)
            
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TemEMPD1 As String
Dim TemEMPD2 As String
Dim TemEMPD3 As String
Dim TemEMPD4 As String
Dim TemEMPD5 As String
Dim TemEMPD6 As String

   TemEMPD1 = EMP_ID
   If EMP_ID <= 0 Then
      TemEMPD1 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO TARGET_DETAIL "
      SQL1 = SQL1 & "( "
      SQL1 = SQL1 & "   TARGET_ID, "
      SQL1 = SQL1 & "   EMP_ID, "
      SQL1 = SQL1 & "   TARGET_PRICE1, "
      SQL1 = SQL1 & "   TARGET_PRICE2, "
      SQL1 = SQL1 & "   TARGET_PRICE3, "
      SQL1 = SQL1 & "   TARGET_PRICE4, "
      SQL1 = SQL1 & "   TARGET_PRICE5, "
      SQL1 = SQL1 & "   TARGET_PRICE6, "
      SQL1 = SQL1 & "   TARGET_PRICE7, "
      SQL1 = SQL1 & "   TARGET_PRICE8, "
      SQL1 = SQL1 & "   TARGET_PRICE9, "
      SQL1 = SQL1 & "   TARGET_PRICE10, "
      SQL1 = SQL1 & "   TARGET_PRICE11, "
      SQL1 = SQL1 & "   TARGET_PRICE12, "
      SQL1 = SQL1 & "   CREATE_DATE, "
      SQL1 = SQL1 & "   CREATE_NAME "
      SQL1 = SQL1 & "   ) "
      SQL1 = SQL1 & " VALUES "
      SQL1 = SQL1 & " ( "
      SQL1 = SQL1 & TARGET_ID & ","
      SQL1 = SQL1 & TemEMPD1 & ","
      SQL1 = SQL1 & TARGET_PRICE1 & ","
      SQL1 = SQL1 & TARGET_PRICE2 & ","
      SQL1 = SQL1 & TARGET_PRICE3 & ","
      SQL1 = SQL1 & TARGET_PRICE4 & ","
      SQL1 = SQL1 & TARGET_PRICE5 & ","
      SQL1 = SQL1 & TARGET_PRICE6 & ","
      SQL1 = SQL1 & TARGET_PRICE7 & ","
      SQL1 = SQL1 & TARGET_PRICE8 & ","
      SQL1 = SQL1 & TARGET_PRICE9 & ","
      SQL1 = SQL1 & TARGET_PRICE10 & ","
      SQL1 = SQL1 & TARGET_PRICE11 & ","
      SQL1 = SQL1 & TARGET_PRICE12 & ","
      SQL1 = SQL1 & "'" & ServerDateTime & "',"
      SQL1 = SQL1 & "'" & ChangeQuote(Trim(glbUser.USER_NAME)) & "'"
      SQL1 = SQL1 & ")  RETURNING TARGET_DETAIL_ID "
      
   Else 'Show edit
      SQL1 = "UPDATE TARGET_DETAIL SET "
      SQL1 = SQL1 & "  EMP_ID = " & TemEMPD1 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE1 = " & TARGET_PRICE1 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE2 = " & TARGET_PRICE2 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE3 = " & TARGET_PRICE3 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE4 = " & TARGET_PRICE4 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE5 = " & TARGET_PRICE5 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE6 = " & TARGET_PRICE6 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE7 = " & TARGET_PRICE7 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE8 = " & TARGET_PRICE8 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE9 = " & TARGET_PRICE9 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE10 = " & TARGET_PRICE10 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE11 = " & TARGET_PRICE11 & ", "
      SQL1 = SQL1 & "  TARGET_PRICE12 = " & TARGET_PRICE12 & ", "
      SQL1 = SQL1 & "    MODIFY_DATE = '" & ServerDateTime & "', "
      SQL1 = SQL1 & "    MODIFY_NAME = '" & glbUser.USER_NAME & "'"
      SQL1 = SQL1 & "WHERE TARGET_DETAIL_ID = " & TARGET_DETAIL_ID
      
   End If
   
   If AddEditMode = SHOW_ADD Then
      TARGET_DETAIL_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function
Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM TARGET_DETAIL WHERE TARGET_DETAIL_ID = " & TARGET_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function

