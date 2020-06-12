VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExWorksPriceItem"
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

Public EX_WORKS_PRICE_ITEM_ID As Long
Public EX_WORKS_PRICE_ID As Long
Public PART_ITEM_ID As Long
Public RATE_TYPE As Long
Public PACKAGE_RATE As Double
Public SocPartType As Long

Public EX_WORKS_PRICE_CODE As String
Public EX_WORKS_PRICE_DESC As String
Public EX_WORKS_PRICE_LEVEL As Long
Public PART_TYPE As Long
Public PART_MASTER_NAME As String
Public PART_NO As String
Public PART_DESC As String
Public GP_VALUE As Double

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

Public DECLARE_NEW_FLAG As String
Public LAST_EDIT_FLAG As String
Public APPROVED_FLAG As String
Public APPROVED_NAME As String
Public VERIFY_FLAG As String
Public VERIFY_NAME As String
Public DECLARE_COUNT As Long

Public Modify_by As Long
Public SUBSCRIBER_ID As Long
Public ACCOUNT_ID As Long
Public ENABLE_FLAG As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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
 
   If Ind = 1 Then
      SelectStr1 = "SELECT EWP.*, EWPI.*,PM.PART_MASTER_NAME, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EX_WORKS_PRICE_ITEM EWPI "
      FromStr = FromStr & "LEFT OUTER JOIN EX_WORKS_PRICE EWP ON (EWPI.EX_WORKS_PRICE_ID = EWP.EX_WORKS_PRICE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EWPI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_MASTER PM ON (PI.PART_MASTER_ID = PM.PART_MASTER_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT MAX(EWPI.DECLARE_COUNT) DECLARE_COUNT "
      SelectStr2 = "SELECT COUNT(*)  "
     
     FromStr = "FROM EX_WORKS_PRICE_ITEM EWPI "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT EWPI.DECLARE_COUNT  "
      SelectStr2 = "SELECT COUNT(*)  "
     
     FromStr = "FROM EX_WORKS_PRICE_ITEM EWPI "
   End If
   
   If EX_WORKS_PRICE_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWPI.EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EWPI.EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID & ") "
      End If
   End If
   
   If EX_WORKS_PRICE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWPI.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EWPI.EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ") "
      End If
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   If Len(ENABLE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AC.ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (AC.ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      End If
   End If
   
   If Trim(VERIFY_FLAG) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (EWPI.VERIFY_FLAG = '" & ChangeQuote(Trim(VERIFY_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (EWPI.VERIFY_FLAG = '" & ChangeQuote(Trim(VERIFY_FLAG)) & "')"
      End If
   End If
   
   If Trim(APPROVED_FLAG) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (EWPI.APPROVED_FLAG = '" & ChangeQuote(Trim(APPROVED_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (EWPI.APPROVED_FLAG = '" & ChangeQuote(Trim(APPROVED_FLAG)) & "')"
      End If
   End If
   
   If Trim(DECLARE_NEW_FLAG) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (EWPI.DECLARE_NEW_FLAG = '" & ChangeQuote(Trim(DECLARE_NEW_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (EWPI.DECLARE_NEW_FLAG = '" & ChangeQuote(Trim(DECLARE_NEW_FLAG)) & "')"
      End If
   End If
   
   If DECLARE_COUNT > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EWPI.DECLARE_COUNT = " & DECLARE_COUNT & ") "
      Else
         WhereStr = WhereStr & "AND (EWPI.DECLARE_COUNT = " & DECLARE_COUNT & ") "
      End If
   End If
     
   OrderByStr = ""
   OrderTypeStr = ""
      
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY  PI.PART_TYPE ASC,PM.PART_MASTER_NAME,PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EX_WORKS_PRICE_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EX_WORKS_PRICE_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY DECLARE_COUNT DESC "
   End If
   
If Ind = 1 Then
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EX_WORKS_PRICE_ITEM_ID = NVLI(Rs("EX_WORKS_PRICE_ITEM_ID"), -1)
      EX_WORKS_PRICE_ID = NVLI(Rs("EX_WORKS_PRICE_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PACKAGE_RATE = NVLD(Rs("PACKAGE_RATE"), 0)
      PART_MASTER_NAME = NVLS(Rs("PART_MASTER_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      EX_WORKS_PRICE_CODE = NVLS(Rs("EX_WORKS_PRICE_CODE"), "")
      EX_WORKS_PRICE_DESC = NVLS(Rs("EX_WORKS_PRICE_DESC"), "")
      
      DECLARE_COUNT = NVLI(Rs("DECLARE_COUNT"), -1)
      DECLARE_NEW_FLAG = NVLS(Rs("DECLARE_NEW_FLAG"), "")
      LAST_EDIT_FLAG = NVLS(Rs("LAST_EDIT_FLAG"), "")
      APPROVED_FLAG = NVLS(Rs("APPROVED_FLAG"), "")
      APPROVED_NAME = NVLS(Rs("APPROVED_NAME"), "")
      VERIFY_FLAG = NVLS(Rs("VERIFY_FLAG"), "")
      VERIFY_NAME = NVLS(Rs("VERIFY_NAME"), "")
      
      Modify_by = NVLI(Rs("MODIFY_BY"), -1)
      
      GP_VALUE = NVLD(Rs("GP_VALUE"), 0)
   ElseIf Ind = 2 Then
      DECLARE_COUNT = NVLI(Rs("DECLARE_COUNT"), -1)
   ElseIf Ind = 3 Then
      DECLARE_COUNT = NVLI(Rs("DECLARE_COUNT"), -1)
   End If
End Function
Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String

Dim TempID01 As String
Dim TempID02 As String
Dim TempID03 As String
Dim TempID04 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID01 = VERIFY_FLAG
   If Len(VERIFY_FLAG) <= 0 Then
      TempID01 = "N"
   End If
   
  TempID02 = APPROVED_FLAG
   If Len(APPROVED_FLAG) <= 0 Then
      TempID02 = "N"
   End If
   
   TempID03 = LAST_EDIT_FLAG
   If Len(LAST_EDIT_FLAG) <= 0 Then
      TempID03 = "N"
   End If
   
   TempID04 = DECLARE_NEW_FLAG
   If Len(DECLARE_NEW_FLAG) <= 0 Then
      TempID04 = "N"
   End If

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SOC_FEATURE_SEQ", id, glbErrorLog)
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("SOC_FEATURE_SEQ", id, glbErrorLog)
      Else
         id = EX_WORKS_PRICE_ITEM_ID
      End If
      
      SQL1 = "INSERT INTO EX_WORKS_PRICE_ITEM " & _
                     "( " & _
                     "   EX_WORKS_PRICE_ITEM_ID, " & _
                     "   EX_WORKS_PRICE_ID, " & _
                     "   RATE_TYPE, " & _
                     "   PART_ITEM_ID, " & _
                     "   PACKAGE_RATE, " & _
                     "   DECLARE_COUNT," & _
                     "   DECLARE_NEW_FLAG," & _
                     "   LAST_EDIT_FLAG, " & _
                     "   APPROVED_FLAG, " & _
                     "   APPROVED_NAME, " & _
                     "   VERIFY_FLAG, " & _
                     "   VERIFY_NAME, " & _
                     "   GP_VALUE, " & _
                     "   CREATE_DATE, " & _
                     "   CREATE_BY, " & _
                     "   MODIFY_DATE, " & _
                     "   MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         id & "," & _
                        "" & EX_WORKS_PRICE_ID & ", " & _
                        RATE_TYPE & ", " & _
                        TempID1 & ", " & _
                        PACKAGE_RATE & ", " & _
                        DECLARE_COUNT & ", " & _
                        "'" & TempID04 & "'," & _
                        "'" & TempID03 & "'," & _
                        "'" & TempID02 & "'," & _
                        "'" & APPROVED_NAME & "'," & _
                        "'" & TempID01 & "'," & _
                        "'" & VERIFY_NAME & "'," & _
                        GP_VALUE & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               EX_WORKS_PRICE_ITEM_ID = id
   Else 'Show edit
      SQL1 = "UPDATE EX_WORKS_PRICE_ITEM SET " & _
                     "   EX_WORKS_PRICE_ID = " & EX_WORKS_PRICE_ID & ", " & _
                     "   RATE_TYPE = " & RATE_TYPE & ", " & _
                     "   PART_ITEM_ID = " & TempID1 & ", " & _
                     "   PACKAGE_RATE = " & PACKAGE_RATE & ", " & _
                     "   DECLARE_COUNT = " & DECLARE_COUNT & ", " & _
                     "   DECLARE_NEW_FLAG= '" & TempID04 & "', " & _
                     "   LAST_EDIT_FLAG= '" & TempID03 & "', " & _
                     "   APPROVED_FLAG  = '" & TempID02 & "', " & _
                     "   APPROVED_NAME  = '" & APPROVED_NAME & "', " & _
                     "   VERIFY_FLAG  = '" & TempID01 & "', " & _
                     "   VERIFY_NAME  = '" & VERIFY_NAME & "', " & _
                     "    GP_VALUE = " & GP_VALUE & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                     "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EX_WORKS_PRICE_ITEM WHERE EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID
      m_Conn.Execute (SQL1)
End Function
Public Sub UpdateApprovedFlag(ProcessType As Long, CountUpdate As Long)
Dim ServerDateTime As String
Dim SQL1 As String
Dim strTable As String
Dim WhereStr As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE EX_WORKS_PRICE_ITEM SET "
 
   If ProcessType = 1 Then
      SQL1 = SQL1 & "   VERIFY_FLAG = '" & VERIFY_FLAG & "', "
      SQL1 = SQL1 & "   VERIFY_NAME = '" & VERIFY_NAME & "'"
   ElseIf ProcessType = 3 Then
      SQL1 = SQL1 & "   APPROVED_FLAG = '" & APPROVED_FLAG & "', "
      SQL1 = SQL1 & "   APPROVED_NAME = '" & APPROVED_NAME & "', "
      SQL1 = SQL1 & "   DECLARE_COUNT = '" & DECLARE_COUNT & "', "
      SQL1 = SQL1 & "   LAST_EDIT_FLAG = '" & LAST_EDIT_FLAG & "'"
   End If
      
      SQL1 = SQL1 & "WHERE (EX_WORKS_PRICE_ITEM_ID = " & EX_WORKS_PRICE_ITEM_ID & ")"
          
            If ProcessType = 1 Then
               SQL1 = SQL1 & " AND (APPROVED_FLAG='N')" 'ถ้าจะตรวจสอบต้องไม่มีการอนุมัติมาก่อน
            ElseIf ProcessType = 3 Then
               SQL1 = SQL1 & " AND (VERIFY_FLAG='Y') " 'ถ้าจะอนุมัติต้องมีการตรวจสอบก่อน
            End If
     SQL1 = SQL1 & " AND (DECLARE_NEW_FLAG='Y') " 'ต้องเป็นที่ตัวประกาศใหม่เท่านั้น
   
    Call m_Conn.Execute(SQL1, CountUpdate)
End Sub
Private Sub Class_Terminate()
 Set m_Conn = Nothing
End Sub



