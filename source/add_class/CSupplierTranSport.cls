VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupplierTranSport"
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

Public SUPPLIER_TRANSPORT_ID As Long
Public SUPPLIER_TRANSPORT_CODE As String
Public SUPPLIER_TRANSPORT_DETAIL As String
Public SUPPLIER_ID As Long
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

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
      SelectStr1 = "SELECT UG.*,SP.*,NM.* "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUPPLIER_TRANSPORT UG "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = UG.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (SN.NAME_ID = NM.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT(SP.SUPPLIER_CODE),NM.LONG_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUPPLIER_TRANSPORT UG "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = UG.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (SN.NAME_ID = NM.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT(UG.SUPPLIER_TRANSPORT_CODE),NM.LONG_NAME,SP.SUPPLIER_CODE  "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUPPLIER_TRANSPORT UG "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = UG.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (SN.NAME_ID = NM.NAME_ID) "

      WhereStr = ""
   End If

   If SUPPLIER_TRANSPORT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.SUPPLIER_TRANSPORT_ID = " & SUPPLIER_TRANSPORT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.SUPPLIER_TRANSPORT_ID = " & SUPPLIER_TRANSPORT_ID & ") "
      End If
   End If

   If Len(SUPPLIER_TRANSPORT_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.SUPPLIER_TRANSPORT_CODE = '" & SUPPLIER_TRANSPORT_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (UG.SUPPLIER_TRANSPORT_CODE = '" & SUPPLIER_TRANSPORT_CODE & "') "
      End If
   End If

   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
   If Len(SUPPLIER_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SP.SUPPLIER_CODE = '" & SUPPLIER_CODE & "') "
      Else
         WhereStr = WhereStr & "AND (SP.SUPPLIER_CODE = '" & SUPPLIER_CODE & "') "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY UG.SUPPLIER_TRANSPORT_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY UG.SUPPLIER_TRANSPORT_CODE "
      Else
         OrderByStr = "ORDER BY UG.SUPPLIER_TRANSPORT_ID "
      End If
   Else
      OrderByStr = "ORDER BY UG.SUPPLIER_TRANSPORT_ID "
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
      SUPPLIER_TRANSPORT_ID = NVLI(Rs("SUPPLIER_TRANSPORT_ID"), -1)
      SUPPLIER_TRANSPORT_CODE = NVLS(Rs("SUPPLIER_TRANSPORT_CODE"), "")
      SUPPLIER_TRANSPORT_DETAIL = NVLS(Rs("SUPPLIER_TRANSPORT_DETAIL"), "")
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("LONG_NAME"), "")
      
      KEY_ID = SUPPLIER_TRANSPORT_ID
      KEY_LOOKUP = SUPPLIER_TRANSPORT_CODE
   ElseIf Ind = 2 Then
       SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("LONG_NAME"), "")
   ElseIf Ind = 3 Then
      SUPPLIER_TRANSPORT_CODE = NVLS(Rs("SUPPLIER_TRANSPORT_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("LONG_NAME"), "")
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SUPPLIER_TRANSPORT_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO SUPPLIER_TRANSPORT " & _
                     "( " & _
                     "   SUPPLIER_TRANSPORT_ID, " & _
                     "   SUPPLIER_TRANSPORT_CODE, " & _
                     "   SUPPLIER_TRANSPORT_DETAIL, " & _
                     "   SUPPLIER_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & SUPPLIER_TRANSPORT_CODE & "'," & _
                         "'" & SUPPLIER_TRANSPORT_DETAIL & "'," & _
                         "" & SUPPLIER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SUPPLIER_TRANSPORT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUPPLIER_TRANSPORT SET " & _
                     "   SUPPLIER_TRANSPORT_CODE = '" & SUPPLIER_TRANSPORT_CODE & "', " & _
                     "   SUPPLIER_TRANSPORT_DETAIL = '" & SUPPLIER_TRANSPORT_DETAIL & "', " & _
                     "   SUPPLIER_ID = " & SUPPLIER_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUPPLIER_TRANSPORT_ID = " & SUPPLIER_TRANSPORT_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUPPLIER_TRANSPORT WHERE SUPPLIER_TRANSPORT_ID = " & SUPPLIER_TRANSPORT_ID
      m_Conn.Execute (SQL1)
End Function

