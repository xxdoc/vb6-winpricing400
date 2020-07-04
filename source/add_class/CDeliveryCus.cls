VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDeliveryCus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public DELIVERY_CUS_ITEM_ID As Long
Public DELIVERY_CUS_ITEM_CODE As String
Public DELIVERY_CUS_ITEM_NAME As String
Public HIDE_FLAG As String
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
Dim id As Long
Dim Temp1 As String
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   TempID1 = HIDE_FLAG
   If Len(HIDE_FLAG) <= 0 Then
      TempID1 = "N"
   End If
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DELIVERY_CUS_ITEM_SEQ", id, glbErrorLog)
      
      SQL1 = "INSERT INTO DELIVERY_CUS_ITEM  " & _
                     "( " & _
                     "   DELIVERY_CUS_ITEM_ID, " & _
                     "   DELIVERY_CUS_ITEM_CODE ," & _
                     "   DELIVERY_CUS_ITEM_NAME, " & _
                     "   CUSTOMER_ID, " & _
                     "   HIDE_FLAG, " & _
                     "   CREATE_DATE, " & _
                     "    CREATE_BY, " & _
                     "    MODIFY_DATE, " & _
                     "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         id & "," & _
                        "'" & ChangeQuote(Trim(DELIVERY_CUS_ITEM_CODE)) & "', " & _
                        "'" & ChangeQuote(Trim(DELIVERY_CUS_ITEM_NAME)) & "', " & _
                        "" & ChangeQuote(Trim(CUSTOMER_ID)) & ", " & _
                        "'" & ChangeQuote(Trim(TempID1)) & "', " & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                        "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "" & _
                     ") "
                  DELIVERY_CUS_ITEM_ID = id
   Else 'Show edit
      SQL1 = "UPDATE DELIVERY_CUS_ITEM  SET " & _
                     "   DELIVERY_CUS_ITEM_CODE = '" & ChangeQuote(Trim(DELIVERY_CUS_ITEM_CODE)) & "', " & _
                     "   DELIVERY_CUS_ITEM_NAME = '" & ChangeQuote(Trim(DELIVERY_CUS_ITEM_NAME)) & "', " & _
                     "   CUSTOMER_ID = " & ChangeQuote(Trim(CUSTOMER_ID)) & ", " & _
                     "   HIDE_FLAG = '" & ChangeQuote(Trim(TempID1)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & "" & _
                    "WHERE DELIVERY_CUS_ITEM_ID = " & DELIVERY_CUS_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM DELIVERY_CUS_ITEM  WHERE DELIVERY_CUS_ITEM_ID = " & DELIVERY_CUS_ITEM_ID
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

   SelectStr1 = "SELECT DCI.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DELIVERY_CUS_ITEM  DCI "
   WhereStr = ""
   
   If DELIVERY_CUS_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DCI.DELIVERY_CUS_ITEM_ID = " & DELIVERY_CUS_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DCI.DELIVERY_CUS_ITEM_ID = " & DELIVERY_CUS_ITEM_ID & ") "
      End If
   End If
   
   If Len(DELIVERY_CUS_ITEM_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DCI.DELIVERY_CUS_ITEM_CODE = '" & DELIVERY_CUS_ITEM_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (DCI.DELIVERY_CUS_ITEM_CODE = '" & DELIVERY_CUS_ITEM_CODE & "') "
      End If
   End If
   
     If Len(HIDE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DCI.HIDE_FLAG = '" & HIDE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (DCI.HIDE_FLAG = '" & HIDE_FLAG & "') "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DCI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DCI.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY DCI.DELIVERY_CUS_ITEM_ID  "
   Else
      OrderByStr = "ORDER BY DCI.DELIVERY_CUS_ITEM_ID  "
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
      DELIVERY_CUS_ITEM_ID = NVLI(Rs("DELIVERY_CUS_ITEM_ID"), -1)
      DELIVERY_CUS_ITEM_CODE = NVLS(Rs("DELIVERY_CUS_ITEM_CODE"), "")
      DELIVERY_CUS_ITEM_NAME = NVLS(Rs("DELIVERY_CUS_ITEM_NAME"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      HIDE_FLAG = NVLS(Rs("HIDE_FLAG"), "")
      
      KEY_LOOKUP = DELIVERY_CUS_ITEM_CODE
      KEY_ID = DELIVERY_CUS_ITEM_ID
   End If
End Function
