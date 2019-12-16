VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPromotional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public PROMOTIONAL_ITEM_ID As Long
Public PROMOTIONAL_DETAIL_ID As Long
Public PROMOTIONAL_DETAIL_NAME As String
Public UNIT_TYPE As Long
Public PROMOTIONAL_RATE As Double
Public CUSTOMER_ID As Long

Public OrderBy As Long
Public OrderType As Long
Public QueryFlag As Long
Public AddEditMode As SHOW_MODE_TYPE
Public Flag As String
Public KEY_ID As Long
Public KEY_LOOKUP As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   QueryFlag = -1
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Temp1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DELIVERY_CUS_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PROMOTIONAL_ITEM  " & _
                     "( " & _
                     "   PROMOTIONAL_ITEM_ID, " & _
                     "   PROMOTIONAL_DETAIL_ID ," & _
                     "   UNIT_TYPE, " & _
                     "   PROMOTIONAL_RATE, " & _
                     "   CUSTOMER_ID, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         PROMOTIONAL_DETAIL_ID & "," & _
                         UNIT_TYPE & "," & _
                         PROMOTIONAL_RATE & "," & _
                         CUSTOMER_ID & "," & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ") "
                  PROMOTIONAL_ITEM_ID = ID

   Else 'Show edit
      SQL1 = "UPDATE PROMOTIONAL_ITEM  SET " & _
                     "   PROMOTIONAL_DETAIL_ID = " & PROMOTIONAL_DETAIL_ID & ", " & _
                     "   UNIT_TYPE = " & UNIT_TYPE & ", " & _
                     "   PROMOTIONAL_RATE = " & PROMOTIONAL_RATE & ", " & _
                     "   CUSTOMER_ID = '" & ChangeQuote(Trim(CUSTOMER_ID)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & "" & _
                    "WHERE PROMOTIONAL_ITEM_ID = " & PROMOTIONAL_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PROMOTIONAL_ITEM  WHERE PROMOTIONAL_ITEM_ID = " & PROMOTIONAL_ITEM_ID
      m_Conn.Execute (SQL1)

End Function

Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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


   SelectStr1 = "SELECT PTI.*,MR1.KEY_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PROMOTIONAL_ITEM  PTI "
   FromStr = FromStr & "LEFT JOIN MASTER_REF MR1 ON (PTI.PROMOTIONAL_DETAIL_ID=MR1.KEY_ID)  "
   WhereStr = ""
   
   If PROMOTIONAL_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTI.PROMOTIONAL_ITEM_ID = " & PROMOTIONAL_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PTI.PROMOTIONAL_ITEM_ID = " & PROMOTIONAL_ITEM_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PTI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If PROMOTIONAL_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PTI.PROMOTIONAL_DETAIL_ID = " & PROMOTIONAL_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PTI.PROMOTIONAL_DETAIL_ID = " & PROMOTIONAL_DETAIL_ID & ") "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PTI.PROMOTIONAL_ITEM_ID  "
   Else
      OrderByStr = "ORDER BY PTI.PROMOTIONAL_ITEM_ID  "
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
      PROMOTIONAL_ITEM_ID = NVLI(Rs("PROMOTIONAL_ITEM_ID"), -1)
      PROMOTIONAL_DETAIL_ID = NVLI(Rs("KEY_ID"), -1)
      PROMOTIONAL_DETAIL_NAME = NVLS(Rs("KEY_NAME"), "")
      UNIT_TYPE = NVLI(Rs("UNIT_TYPE"), -1)
      PROMOTIONAL_RATE = NVLD(Rs("PROMOTIONAL_RATE"), 0)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      
      KEY_LOOKUP = PROMOTIONAL_DETAIL_NAME
      KEY_ID = PROMOTIONAL_ITEM_ID
   End If
End Function
