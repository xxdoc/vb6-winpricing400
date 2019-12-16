VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormulaExt"
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

Public FORMULA_ID As Long
Public FORMULA_NO As String
Public FORMULA_NAME As String
Public FORMULA_CODE As String
Public FORMULA_DATE As Date
Public FORMULA_TYPE As Long
Public FORMULA_DESC As String
Public APPROVED_BY As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public CAN_UNIT_CHARGE As Double
Public UP_MARK_CHARGE As Double
Public LOST_PERCENT As Double
Public PMC As Double
Public RMC As Double

Public PART_TYPE_ID As Long
Public PART_TYPE_NAME As String
Public PART_NO As String
Public PART_ITEM_NAME As String
Public LONG_NAME As String
Public LAST_NAME As String
Public FORMULA_TYPE_NAME As String
Public LOCATION_NO As String
Public LOCATION_NAME As String
Public INTERMEDIAT_FLAG As String

Public SUM_REAL_AMOUNT As Double
Public KEY_ID As Long
Public KEY_LOOKUP As String

Public FROM_DATE As Date
Public TO_DATE As Date

Public Inputs As Collection
Public FormulaItems As Collection
Public FormulaPrices As Collection
Public FormulaVariables As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set FormulaItems = New Collection
   Set FormulaPrices = New Collection
   Set Inputs = New Collection
   Set FormulaVariables = New Collection
   
   FROM_DATE = -1
   TO_DATE = -1
   FORMULA_DATE = -1
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

   If Ind = 1 Then
      SelectStr1 = "SELECT BL.* "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM FORMULA_EXT BL "
      
      WhereStr = ""
   ElseIf Ind = 2 Then

   End If
   
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
   
     If Trim(FORMULA_NO) <> "" Then
      If WildCard(FORMULA_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FORMULA_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FORMULA_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FORMULA_NO = '" & ChangeQuote(Trim(FORMULA_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (FORMULA_NO = '" & ChangeQuote(Trim(FORMULA_NO)) & "')"
         End If
      End If
   End If

     If Trim(FORMULA_DESC) <> "" Then
      If WildCard(FORMULA_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(FORMULA_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(FORMULA_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (FORMULA_DESC = '" & ChangeQuote(Trim(FORMULA_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (FORMULA_DESC = '" & ChangeQuote(Trim(FORMULA_DESC)) & "')"
         End If
      End If
   End If

   If FORMULA_DATE > -1 Then
      TempDate = DateToStringIntLow(FORMULA_DATE)
      If WhereStr = "" Then
         WhereStr = "WHERE (FORMULA_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (FORMULA_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If

 If FORMULA_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FORMULA_TYPE = " & FORMULA_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (FORMULA_TYPE = " & FORMULA_TYPE & ") "
      End If
   End If
   
    If APPROVED_BY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (APPROVED_BY = " & APPROVED_BY & ") "
      Else
         WhereStr = WhereStr & "AND (APPROVED_BY = " & APPROVED_BY & ") "
      End If
   End If
   
    If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If

     If Trim(PART_NO) <> "" Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If

    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BL.FORMULA_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BL.FORMULA_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BL.FORMULA_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BL.FORMULA_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   '===
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FORMULA_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY FORMULA_DESC "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY FORMULA_DATE "
ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY FORMULA_TYPE "
ElseIf OrderBy = 5 Then
      OrderByStr = "ORDER BY APPROVED_BY "
ElseIf OrderBy = 6 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   Else
      OrderByStr = "ORDER BY FORMULA_ID "
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
'''Debug.Print (SQL1)
      Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
   FORMULA_NO = NVLS(Rs("FORMULA_NO"), "")
   FORMULA_DESC = NVLS(Rs("FORMULA_DESC"), "")
   FORMULA_DATE = InternalDateToDate(NVLS(Rs("FORMULA_DATE"), ""))
   FORMULA_TYPE = NVLI(Rs("FORMULA_TYPE"), 0)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   APPROVED_BY = NVLI(Rs("APPROVED_BY"), -1)
   LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   RMC = NVLD(Rs("RMC"), 0)
   PMC = NVLD(Rs("PMC"), 0)
   
   KEY_ID = FORMULA_ID
   KEY_LOOKUP = FORMULA_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = FORMULA_TYPE
   If FORMULA_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
         
TempID2 = APPROVED_BY
   If APPROVED_BY <= 0 Then
      TempID2 = "NULL"
   End If
TempID3 = PART_ITEM_ID
   If FORMULA_TYPE <= 0 Then
      TempID3 = "NULL"
   End If
         
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FORMULA_SEQ", id, glbErrorLog)
               
      SQL1 = "INSERT INTO FORMULA " & _
                     "( " & _
                     "   FORMULA_ID, " & _
                     "   FORMULA_NO, " & _
                     "   FORMULA_TYPE, " & _
                     "   FORMULA_DATE, " & _
                     "   FORMULA_DESC, " & _
                     "   APPROVED_BY, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
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
                         id & "," & _
                         "'" & ChangeQuote(Trim(FORMULA_NO)) & "'," & _
                         TempID1 & "," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(FORMULA_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(FORMULA_DESC)) & "'," & _
                         TempID2 & "," & _
                         TempID3 & "," & _
                         LOCATION_ID & "," & _
                         RMC & "," & _
                         PMC & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         FORMULA_ID = id
         
   Else 'Show edit
      SQL1 = "UPDATE FORMULA SET " & _
                     "   FORMULA_NO = '" & ChangeQuote(Trim(FORMULA_NO)) & "', " & _
                    "    FORMULA_TYPE = " & TempID1 & ", " & _
                    "   FORMULA_DATE = '" & ChangeQuote(Trim(DateToStringInt(FORMULA_DATE))) & "', " & _
                     "   FORMULA_DESC = '" & ChangeQuote(Trim(FORMULA_DESC)) & "', " & _
                    "    APPROVED_BY = " & TempID2 & ", " & _
                    "    PART_ITEM_ID = " & TempID3 & ", " & _
                    "    LOCATION_ID = " & LOCATION_ID & ", " & _
                    "    RMC = " & RMC & ", " & _
                    "    PMC = " & PMC & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FORMULA_ID = " & FORMULA_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String


      SQL1 = "DELETE FROM FORMULA_VARIABLE WHERE FORMULA_ID = " & FORMULA_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM FORMULA_ITEM WHERE FORMULA_ID = " & FORMULA_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM FORMULA WHERE FORMULA_ID = " & FORMULA_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Inputs = Nothing
   Set FormulaItems = Nothing
   Set FormulaPrices = Nothing
   Set FormulaVariables = Nothing
End Sub


