VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartItemSpec"
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

Public PARTITEM_SPEC_ID As Long
Public PART_ITEM_ID As Long
Public FROM_RATE As Double
Public TO_RATE As Double
Public HUMIDITY_WEIGHT As Double

Public PART_TYPE As Long
Public PART_NO As String
Public PART_DESC As String

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
      
      FromStr = "FROM PARTITEM_SPEC UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If PARTITEM_SPEC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PARTITEM_SPEC_ID = " & PARTITEM_SPEC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PARTITEM_SPEC_ID = " & PARTITEM_SPEC_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PARTITEM_SPEC_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PARTITEM_SPEC_ID "
      Else
         OrderByStr = "ORDER BY FROM_RATE ASC, TO_RATE "
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
      PARTITEM_SPEC_ID = NVLI(Rs("PARTITEM_SPEC_ID"), -1)
      FROM_RATE = NVLD(Rs("FROM_RATE"), 0)
      TO_RATE = NVLD(Rs("TO_RATE"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HUMIDITY_WEIGHT = NVLD(Rs("HUMIDITY_WEIGHT"), 0)
      
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
      Call glbDatabaseMngr.GetSeqID("PARTITEM_SPEC_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PARTITEM_SPEC " & _
                     "( " & _
                     "   PARTITEM_SPEC_ID, " & _
                     "   FROM_RATE, " & _
                     "   TO_RATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   HUMIDITY_WEIGHT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & FROM_RATE & "," & _
                         "" & TO_RATE & "," & _
                         "" & PART_ITEM_ID & "," & _
                         HUMIDITY_WEIGHT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PARTITEM_SPEC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PARTITEM_SPEC SET " & _
                     "   FROM_RATE = " & FROM_RATE & ", " & _
                     "   TO_RATE = " & TO_RATE & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   HUMIDITY_WEIGHT = " & HUMIDITY_WEIGHT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PARTITEM_SPEC_ID = " & PARTITEM_SPEC_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PARTITEM_SPEC WHERE PARTITEM_SPEC_ID = " & PARTITEM_SPEC_ID
      m_Conn.Execute (SQL1)
End Function



