VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStpTierVol"
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

Public STPTIER_VOL_ID As Long
Public SOC_FEATURE_ID As Long
Public FROM_QUANTITY As Double
Public TO_QUANTITY As Double
Public Width As Long
Public RATE_AMOUNT As Double

Public UC_RATE_ID As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any tranSOC_tion
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*, UR.RATE_AMOUNT, UR.UC_RATE_ID  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM STPTIER_VOL UG, UC_RATE UR "

   WhereStr = "WHERE (UG.STPTIER_VOL_ID = UR.STPTIER_VOL_ID) "
   
   If STPTIER_VOL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STPTIER_VOL_ID = " & STPTIER_VOL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (STPTIER_VOL_ID = " & STPTIER_VOL_ID & ") "
      End If
   End If
   
   If SOC_FEATURE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.SOC_FEATURE_ID = " & SOC_FEATURE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY STPTIER_VOL_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY STPTIER_VOL_ID "
   Else
      OrderByStr = "ORDER BY FROM_QUANTITY, TO_QUANTITY "
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   STPTIER_VOL_ID = NVLI(Rs("STPTIER_VOL_ID"), -1)
   SOC_FEATURE_ID = NVLI(Rs("SOC_FEATURE_ID"), -1)
   FROM_QUANTITY = NVLD(Rs("FROM_QUANTITY"), -1)
   TO_QUANTITY = NVLD(Rs("TO_QUANTITY"), -1)
   RATE_AMOUNT = NVLD(Rs("RATE_AMOUNT"), -1)
   Width = TO_QUANTITY - FROM_QUANTITY
   
   UC_RATE_ID = NVLI(Rs("UC_RATE_ID"), -1)
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("STPTIER_VOL_SEQ", ID, glbErrorLog)
      Else
         ID = STPTIER_VOL_ID
      End If
      
      SQL1 = "INSERT INTO STPTIER_VOL " & _
                     "( " & _
                     "   STPTIER_VOL_ID, " & _
                     "   SOC_FEATURE_ID, " & _
                     "   FROM_QUANTITY, " & _
                     "   TO_QUANTITY, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & SOC_FEATURE_ID & "," & _
                         "" & FROM_QUANTITY & "," & _
                         "" & TO_QUANTITY & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         STPTIER_VOL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STPTIER_VOL SET " & _
                     "   SOC_FEATURE_ID = " & SOC_FEATURE_ID & ", " & _
                     "   FROM_QUANTITY = " & FROM_QUANTITY & ", " & _
                     "   TO_QUANTITY = " & TO_QUANTITY & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STPTIER_VOL_ID = " & STPTIER_VOL_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM STPTIER_VOL WHERE STPTIER_VOL_ID = " & STPTIER_VOL_ID
      m_Conn.Execute (SQL1)
End Function

Public Function CopyObject(C As CStpTierVol)
   C.STPTIER_VOL_ID = STPTIER_VOL_ID
   C.FROM_QUANTITY = FROM_QUANTITY
   C.TO_QUANTITY = TO_QUANTITY
   C.RATE_AMOUNT = RATE_AMOUNT
   C.SOC_FEATURE_ID = SOC_FEATURE_ID
   C.Width = Width
   
   C.UC_RATE_ID = UC_RATE_ID
End Function




