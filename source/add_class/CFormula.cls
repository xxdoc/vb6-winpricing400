VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormula"
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
Public CANCEL_FLAG As String

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
      SelectStr1 = "SELECT BL.*,NM.LONG_NAME,NM.LAST_NAME,FT.INTERMEDIAT_FLAG, FT.FORMULA_TYPE_NAME, PI.PART_DESC, PI.PART_TYPE, PI.PART_NO, PI.PART_DESC, PT.PART_TYPE_NAME, LC.LOCATION_NO, LC.LOCATION_NAME "
      SelectStr1 = SelectStr1 & ", (SELECT SUM(FI.REAL_AMOUNT) FROM FORMULA_ITEM FI WHERE (FI.FORMULA_ID = BL.FORMULA_ID)) SUM_REAL_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM FORMULA BL "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EM.EMP_ID = BL.APPROVED_BY) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EM.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN FORMULA_TYPE FT ON (BL.FORMULA_TYPE = FT.FORMULA_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BL.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BL.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT BL.FORMULA_NO "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM FORMULA BL "
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
   
   If Len(Trim(CANCEL_FLAG)) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (BL.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
      Else
         WhereStr = WhereStr & "AND (BL.CANCEL_FLAG = '" & ChangeQuote(Trim(CANCEL_FLAG)) & "')"
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
         WhereStr = "WHERE (PI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
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
      CANCEL_FLAG = NVLS(Rs("CANCEL_FLAG"), "")
      
      FORMULA_TYPE_NAME = NVLS(Rs("FORMULA_TYPE_NAME"), "")
      INTERMEDIAT_FLAG = NVLS(Rs("INTERMEDIAT_FLAG"), "")
      LONG_NAME = NVLS(Rs("LONG_NAME"), "")
      LAST_NAME = NVLS(Rs("LAST_NAME"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PART_TYPE_ID = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_ITEM_NAME = NVLS(Rs("PART_DESC"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      SUM_REAL_AMOUNT = NVLD(Rs("SUM_REAL_AMOUNT"), 0)
      
      KEY_ID = FORMULA_ID
      KEY_LOOKUP = FORMULA_NO
   ElseIf Ind = 2 Then
      FORMULA_NO = NVLS(Rs("FORMULA_NO"), "")
   End If
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
                     "   CANCEL_FLAG, " & _
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
                        "'" & CANCEL_FLAG & "'," & _
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
                    "    CANCEL_FLAG = '" & CANCEL_FLAG & "', " & _
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
