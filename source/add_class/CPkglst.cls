VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPkglst"
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

Public PKGLST_ITEM_ID As Long
Public BILLING_DOC_ID As Long
Public PKG_NUMBER As String
Public DESCRIPTION As String
Public QUANTITY As Double
Public MEASURE As Double
Public NET_WEIGHT As Double
Public GROSS_WEIGHT As Double
Public SET_FLAG As String




Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   SET_FLAG = "N"
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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PKGLST_ITEM UG "
      'FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
   
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PKGLST_ITEM_ID "
       Else
         OrderByStr = "ORDER BY PKGLST_ITEM_ID "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      
      PKGLST_ITEM_ID = NVLI(Rs("PKGLST_ITEM_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      PKG_NUMBER = NVLS(Rs("PKG_NUMBER"), "")
      DESCRIPTION = NVLS(Rs("DESCRIPTION"), "")
      QUANTITY = NVLD(Rs("QUANTITY"), 0)
      MEASURE = NVLD(Rs("MEASURE"), 0)
      NET_WEIGHT = NVLD(Rs("NET_WEIGHT"), 0)
      GROSS_WEIGHT = NVLD(Rs("GROSS_WEIGHT"), 0)
      SET_FLAG = NVLS(Rs("SET_FLAG"), "")
         End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim TempID1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PKGLST_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PKGLST_ITEM " & _
                     "( " & _
                     "   PKGLST_ITEM_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   PKG_NUMBER, " & _
                     "   DESCRIPTION, " & _
                     "   QUANTITY, " & _
                     "   MEASURE, " & _
                     "   NET_WEIGHT, " & _
                     "   GROSS_WEIGHT, " & _
                     "   SET_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
            SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                           "'" & ChangeQuote(Trim(PKG_NUMBER)) & "', " & _
                           "'" & ChangeQuote(Trim(DESCRIPTION)) & "', " & _
                         "" & QUANTITY & "," & _
                         "" & MEASURE & "," & _
                         "" & NET_WEIGHT & "," & _
                         "" & GROSS_WEIGHT & "," & _
                           "'" & ChangeQuote(Trim(SET_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PKGLST_ITEM_ID = ID
   Else
      SQL1 = "UPDATE PKGLST_ITEM SET " & _
                    "   PKG_NUMBER = '" & ChangeQuote(Trim(PKG_NUMBER)) & "', " & _
                    "   DESCRIPTION = '" & ChangeQuote(Trim(DESCRIPTION)) & "', " & _
                     "   QUANTITY = " & QUANTITY & ", " & _
                     "   MEASURE = " & MEASURE & ", " & _
                     "   NET_WEIGHT = " & NET_WEIGHT & ", " & _
                     "   GROSS_WEIGHT = " & GROSS_WEIGHT & ", " & _
                     "   SET_FLAG = '" & ChangeQuote(Trim(SET_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PKGLST_ITEM_ID = " & PKGLST_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PKGLST_ITEM WHERE PKGLST_ITEM_ID = " & PKGLST_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

