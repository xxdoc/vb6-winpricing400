VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobOutputEn"
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
'
Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_DESC As String
Public UNIT_COUNT As Long
Public PART_TYPE As Long
Public MINIMUM_ALLOW As Double
Public PIG_FLAG As String
Public PIG_TYPE As String
Public UNIT_WEIGHT As Double
Public BARCODE_NO As String

Public LOCATION_ID As Long
Public UNIT_NAME As String
Public PART_TYPE_NAME As String
Public PART_TYPE_NO As String
Public FROM_DATE As Date
Public TO_DATE As Date

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public SUCCESS_FLAG As String
Public PART_GROUP_ID  As Long
Public JOB_DATE As String
Public TX_AMOUNT As Double



Public PartLocations As Collection
Public Suppliers As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartLocations = New Collection
   Set Suppliers = New Collection
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
Dim GroupByStr As String
Dim TempDate As String

Dim SubLen As Long
Dim NewStr As String

    SelectStr1 = "SELECT  DISTINCT  PTY.PART_TYPE_NAME , SUM(JIO.TX_AMOUNT)  TX_AMOUNT "
   SelectStr2 = "SELECT COUNT(*)  "
   
    FromStr = "  FROM JOB J  "
    FromStr = FromStr & "LEFT OUTER JOIN JOB_INOUT JIO ON (J.JOB_ID = JIO.JOB_ID) "
    FromStr = FromStr & "LEFT OUTER JOIN FORMULA FM ON (FM.FORMULA_ID = JIO.FROM_FORMULA)"
    FromStr = FromStr & "LEFT OUTER JOIN FORMULA_TYPE FMTY ON (FM.FORMULA_TYPE = FMTY.FORMULA_TYPE_ID)"
    FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PT ON (FM.PART_ITEM_ID = PT.PART_ITEM_ID)"
    FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PTY ON (PT.PART_TYPE = PTY.PART_TYPE_ID)"
    FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PTY.PART_GROUP_ID = PG.PART_GROUP_ID)"
   
   WhereStr = ""
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
      
   If UNIT_COUNT > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UNIT_COUNT = " & UNIT_COUNT & ") "
      Else
         WhereStr = WhereStr & "AND (UNIT_COUNT = " & UNIT_COUNT & ") "
      End If
   End If
      
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If Len(PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_TYPE = '" & PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_TYPE = '" & PIG_TYPE & "') "
      End If
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
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

   If Len(Trim(PART_DESC)) > 0 Then
      If WildCard(PART_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(BARCODE_NO)) > 0 Then
      If WildCard(BARCODE_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(BARCODE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(BARCODE_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "')"
         End If
      End If
   End If
   
   If Len(SUCCESS_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUCCESS_FLAG = '" & SUCCESS_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SUCCESS_FLAG = '" & SUCCESS_FLAG & "') "
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_GROUP_ID = " & PART_GROUP_ID & ")  "
      Else
         WhereStr = WhereStr & "AND (PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   If WhereStr = "" Then
        WhereStr = " WHERE ( PART_TYPE_NAME IS NOT NULL ) AND (PART_TYPE_NAME  LIKE '%FINISH%' ) "
      Else
         WhereStr = WhereStr & " AND (PART_TYPE_NAME IS NOT NULL )  AND (PART_TYPE_NAME  LIKE '%FINISH%' ) "
      End If
      
    
    
   OrderByStr = ""
   OrderTypeStr = ""
   GroupByStr = ""

    If GroupByStr = "" Then
        GroupByStr = "  GROUP BY    PTY.PART_TYPE_NAME  "
   Else
   End If

   OrderByStr = "ORDER BY PTY.PART_TYPE_NAME "
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
      
   
   
   
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
If Ind = 1 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      UNIT_COUNT = NVLI(Rs("UNIT_COUNT"), -1)
      MINIMUM_ALLOW = NVLS(Rs("MINIMUM_ALLOW"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PIG_FLAG = NVLS(Rs("PIG_FLAG"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      UNIT_WEIGHT = NVLD(Rs("UNIT_WEIGHT"), 0)
      BARCODE_NO = NVLS(Rs("BARCODE_NO"), "")
   
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
   
      KEY_ID = PART_ITEM_ID
      KEY_LOOKUP = PART_NO
ElseIf Ind = 2 Then
 'SUCCESS_FLAG = NVLS(Rs("SUCCESS_FLAG"), "")
 'PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
' JOB_DATE = NVLS(Rs("JOB_DATE"), "")
   PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   TX_AMOUNT = NVLI(Rs("TX_AMOUNT"), 0)
End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = PART_TYPE
   If PART_TYPE <= 0 Then
      TempID1 = "NULL "
   End If
      
   TempID2 = UNIT_COUNT
   If UNIT_COUNT <= 0 Then
      TempID2 = "NULL "
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PART_ITEM " & _
                     "( " & _
                     "   PART_ITEM_ID, " & _
                     "   PART_NO, " & _
                     "   UNIT_COUNT, " & _
                     "   MINIMUM_ALLOW, " & _
                     "   PART_TYPE, " & _
                     "   PIG_FLAG, " & _
                     "   PART_DESC, " & _
                     "   PIG_TYPE, " & _
                     "   UNIT_WEIGHT, " & _
                     "   BARCODE_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PART_NO)) & "'," & _
                         "" & TempID2 & "," & _
                         "'" & ChangeQuote(Trim(MINIMUM_ALLOW)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PIG_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(PIG_TYPE)) & "'," & _
                         UNIT_WEIGHT & "," & _
                         "'" & ChangeQuote(Trim(BARCODE_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_ITEM SET " & _
                     "   PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "', " & _
                     "   UNIT_COUNT = " & TempID2 & ", " & _
                     "   MINIMUM_ALLOW = '" & MINIMUM_ALLOW & "', " & _
                     "   PART_TYPE = " & TempID1 & ", " & _
                     "   PIG_FLAG = '" & ChangeQuote(Trim(PIG_FLAG)) & "', " & _
                     "   PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "', " & _
                     "   PIG_TYPE = '" & ChangeQuote(Trim(PIG_TYPE)) & "', " & _
                     "   UNIT_WEIGHT = " & UNIT_WEIGHT & ", " & _
                     "   BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_ITEM_ID = " & PART_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
   ' LOT_ITEM  LOT_INOUT
      
      SQL1 = "DELETE FROM PART_LOCATION WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)

End Function

Private Sub Class_Terminate()
   Set PartLocations = Nothing
   Set Suppliers = Nothing
End Sub






