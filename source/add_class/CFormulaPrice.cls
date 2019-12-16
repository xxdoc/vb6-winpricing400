VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormulaPrice"
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
Public ExtraFlag As Long
' End of generic data member

Public FORMULA_PRICE_ID As Long
Public FORMULA_ID As Long
Public PRICE_DATE As Date
Public RMC As Double
Public PMC As Double

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
Dim TempDate As String

   SelectStr1 = "SELECT BL.* "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FORMULA_PRICE BL "
   
   WhereStr = ""
   
   If FORMULA_PRICE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FORMULA_PRICE_ID = " & FORMULA_PRICE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FORMULA_PRICE_ID = " & FORMULA_PRICE_ID & ") "
      End If
   End If
         
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
   
   '===
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRICE_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PRICE_DATE "
   Else
      OrderByStr = "ORDER BY FORMULA_PRICE_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   FORMULA_PRICE_ID = NVLI(Rs("FORMULA_PRICE_ID"), -1)
   PRICE_DATE = InternalDateToDate(NVLS(Rs("PRICE_DATE"), ""))
   FORMULA_ID = NVLI(Rs("FORMULA_ID"), 0)
   RMC = NVLD(Rs("RMC"), 0)
   PMC = NVLD(Rs("PMC"), 0)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = FORMULA_ID
   If FORMULA_ID <= 0 Then
      TempID1 = "NULL"
   End If
         
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FORMULA_PRICE_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO FORMULA_PRICE " & _
                     "( " & _
                     "   FORMULA_PRICE_ID, " & _
                     "   FORMULA_ID, " & _
                     "   PRICE_DATE, " & _
                     "   RMC, " & _
                     "   PMC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         TempID1 & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PRICE_DATE))) & "'," & _
                         RMC & "," & _
                         PMC & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         FORMULA_PRICE_ID = ID

   Else 'Show edit
      SQL1 = "UPDATE FORMULA_PRICE SET " & _
                     "    FORMULA_ID = " & TempID1 & ", " & _
                     "   PRICE_DATE = '" & ChangeQuote(Trim(DateToStringInt(PRICE_DATE))) & "', " & _
                     "    RMC = " & RMC & ", " & _
                     "    PMC = " & PMC & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FORMULA_PRICE_ID = " & FORMULA_PRICE_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM FORMULA_PRICE WHERE FORMULA_PRICE_ID = " & FORMULA_PRICE_ID
      m_Conn.Execute (SQL1)
End Function


