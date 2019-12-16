VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormulaItem"
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

Public FORMULA_ITEM_ID As Long
Public FORMULA_ID As Long
Public PART_ITEM_ID As Long
Public ITEM_PERCENT As Double
Public FROM_FORMULA As Long
Public LOCATION_ID As Long
Public AVG_PRICE As Double
Public GROUP_NO As Long
Public REAL_AMOUNT As Double

Public PART_TYPE_ID As Long
Public PART_NO As String
Public PART_ITEM_NAME As String
Public PART_TYPE_NO As String
Public PART_TYPE_NAME As String
Public LOCATION_NO As String
Public LOCATION_NAME As String
Public FROM_FORMULA_NO As String
Public UNIT_NAME As String

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

   SelectStr1 = "SELECT FM2.FORMULA_NO FROM_FORMULA_NO, UN.UNIT_NAME, BL.*,PT.PART_TYPE_NAME, PT.PART_TYPE_NO, PART_TYPE_ID,PI.PART_NO,PI.PART_DESC, LC.LOCATION_NO, LC.LOCATION_NAME "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FORMULA_ITEM BL "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BL.PART_ITEM_ID = PI.PART_ITEM_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PT.PART_TYPE_ID = PI.PART_TYPE) "
   FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BL.LOCATION_ID = LC.LOCATION_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN FORMULA FM2 ON (BL.FROM_FORMULA = FM2.FORMULA_ID) "

   WhereStr = ""
   
   If FORMULA_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FORMULA_ITEM_ID = " & FORMULA_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FORMULA_ITEM_ID = " & FORMULA_ITEM_ID & ") "
      End If
   End If
         
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BL.FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BL.FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
   
   '===
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ITEM_TYPE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ITEM_TYPE "
   Else
      OrderByStr = "ORDER BY GROUP_NO ASC, FORMULA_ITEM_ID "
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
   FORMULA_ITEM_ID = NVLI(Rs("FORMULA_ITEM_ID"), -1)
   FORMULA_ID = NVLI(Rs("FORMULA_ID"), 0)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   ITEM_PERCENT = NVLD(Rs("ITEM_PERCENT"), 0)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   GROUP_NO = NVLI(Rs("GROUP_NO"), 0)
   REAL_AMOUNT = NVLD(Rs("REAL_AMOUNT"), 0)
   
   PART_NO = NVLS(Rs("PART_NO"), "")
   PART_ITEM_NAME = NVLS(Rs("PART_DESC"), "")
   PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
   PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
   FROM_FORMULA = NVLI(Rs("FROM_FORMULA"), -1)
   LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
   LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   FROM_FORMULA_NO = NVLS(Rs("FROM_FORMULA_NO"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

  TempID1 = FORMULA_ID
   If FORMULA_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
  TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
         
  TempID3 = FROM_FORMULA
   If FROM_FORMULA <= 0 Then
      TempID3 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FORMULA_ITEM_SEQ", id, glbErrorLog)
               
      SQL1 = "INSERT INTO FORMULA_ITEM " & _
                     "( " & _
                     "   FORMULA_ITEM_ID, " & _
                     "   FORMULA_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   ITEM_PERCENT, " & _
                     "   LOCATION_ID, " & _
                     "   FROM_FORMULA, " & _
                     "   AVG_PRICE, " & _
                     "   GROUP_NO, " & _
                     "   REAL_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         id & "," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         ITEM_PERCENT & "," & _
                         LOCATION_ID & "," & _
                         TempID3 & "," & _
                         AVG_PRICE & "," & _
                         GROUP_NO & "," & _
                         REAL_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         FORMULA_ITEM_ID = id
         
   Else 'Show edit
      SQL1 = "UPDATE FORMULA_ITEM SET " & _
                     "    FORMULA_ID = " & TempID1 & ", " & _
                     "    PART_ITEM_ID = " & TempID2 & ", " & _
                     "    ITEM_PERCENT = " & ITEM_PERCENT & ", " & _
                     "    LOCATION_ID = " & LOCATION_ID & ", " & _
                     "    FROM_FORMULA = " & TempID3 & ", " & _
                     "    AVG_PRICE = " & AVG_PRICE & ", " & _
                     "    GROUP_NO = " & GROUP_NO & ", " & _
                     "    REAL_AMOUNT = " & REAL_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FORMULA_ITEM_ID = " & FORMULA_ITEM_ID
   End If
''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM FORMULA_ITEM WHERE FORMULA_ITEM_ID = " & FORMULA_ITEM_ID
       m_Conn.Execute (SQL1)
End Function

