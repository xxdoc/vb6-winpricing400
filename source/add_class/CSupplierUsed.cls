VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupplierUsed"
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

Public SUPPLIER_USED_ID As Long
Public SUPPLIER_ID As Long
Public PART_ITEM_ID As Long
Public USED_FLAG As String

Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long

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

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUPPLIER_USED UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If SUPPLIER_USED_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_USED_ID = " & SUPPLIER_USED_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_USED_ID = " & SUPPLIER_USED_ID & ") "
      End If
   End If
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SUPPLIER_USED_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SUPPLIER_USED_ID "
      Else
         OrderByStr = "ORDER BY SUPPLIER_USED_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      SUPPLIER_USED_ID = NVLI(Rs("SUPPLIER_USED_ID"), -1)
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      USED_FLAG = NVLS(Rs("USED_FLAG"), "")
      
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SUPPLIER_USED_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SUPPLIER_USED " & _
                     "( " & _
                     "   SUPPLIER_USED_ID, " & _
                     "   SUPPLIER_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   USED_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & SUPPLIER_ID & "," & _
                         "" & PART_ITEM_ID & "," & _
                         "'" & USED_FLAG & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SUPPLIER_USED_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUPPLIER_USED SET " & _
                     "   SUPPLIER_ID = " & SUPPLIER_ID & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   USED_FLAG = '" & USED_FLAG & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUPPLIER_USED_ID = " & SUPPLIER_USED_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUPPLIER_USED WHERE SUPPLIER_USED_ID = " & SUPPLIER_USED_ID
      m_Conn.Execute (SQL1)
End Function

