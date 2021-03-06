VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartItemExt"
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
Public MAXIMUM_ALLOW As Double
Public PIG_FLAG As String
Public PIG_TYPE As String
Public UNIT_WEIGHT As Double
Public BARCODE_NO As String
Public BILL_DESC As String
Public WEIGHT_PER_PACK As Double
Public PARCEL_TYPE As Long
Public OLD_PART_ID As Long

Public LOCATION_ID As Long
Public UNIT_NAME As String

Public PART_TYPE_NAME As String
Public PART_TYPE_NO As String

Public PART_GROUP_ID As Long
Public PART_GROUP_NO As String
Public PART_GROUP_NAME As String

Public FROM_DATE As Date
Public TO_DATE As Date

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public SUCCESS_FLAG As String
Public JOB_DATE As String
Public JOB_NO  As String
Public MODIFY_DATE As Date
Public TX_AMOUNT As Long

Public PartLocations As Collection
Public Suppliers As Collection
Public HumidRates As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartLocations = New Collection
   Set Suppliers = New Collection
   Set HumidRates = New Collection
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
   SelectStr1 = "SELECT ADR.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM_EXT ADR  "
ElseIf Ind = 2 Then
ElseIf Ind = 3 Then
   SelectStr1 = "SELECT DISTINCT  JIO.TX_AMOUNT , ADR.*, PT.*, UN.* , JV.MODIFY_DATE  MD,J.JOB_NO   "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM ADR  "
   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.UNIT_COUNT = UN.UNIT_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN JOB_VERIFY JV ON ( JV.PART_ITEM_ID = ADR.PART_ITEM_ID ) "
   FromStr = FromStr & " LEFT OUTER JOIN JOB_INOUT JIO ON ( JIO.PART_ITEM_ID = JV.PART_ITEM_ID ) "
   FromStr = FromStr & " LEFT OUTER JOIN JOB  J  ON ( JIO.JOB_ID = J.JOB_ID ) "
   
ElseIf Ind = 4 Then
   SelectStr1 = "SELECT ADR.*, PT.*,PG.*, UN.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_ITEM ADR  "
   FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PG.PART_GROUP_ID = PT.PART_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.UNIT_COUNT = UN.UNIT_ID) "
End If
  
   
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
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   
   If PART_TYPE_NO <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE_NO = '" & PART_TYPE_NO & "') "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE_NO = '" & PART_TYPE_NO & "') "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.CREATE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (ADR.CREATE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (ADR.CREATE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Ind = 1 Then
   ElseIf Ind = 2 Then
   ElseIf Ind = 3 Then
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 4 Then
      OrderByStr = "ORDER BY  PART_GROUP_NO,PART_TYPE_NO,PART_NO "
   Else
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY  PART_ITEM_ID  "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY  PART_DESC  "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY  PART_TYPE_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY  PART_TYPE_NAME "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY  PART_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY  PART_DESC "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY  BARCODE_NO "
      Else
         OrderByStr = "ORDER BY  PART_ITEM_ID "
      End If
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

   KEY_ID = PART_ITEM_ID
   KEY_LOOKUP = PART_NO

 ' SUCCESS_FLAG = NVLS(Rs("SUCCESS_FLAG"), "")
 ' PART_TYPE_NO = NVLI(Rs("PART_TYPE_NO"), -1)
   If Ind = 1 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      UNIT_COUNT = NVLI(Rs("UNIT_COUNT"), -1)
      MINIMUM_ALLOW = NVLD(Rs("MINIMUM_ALLOW"), 0)
      MAXIMUM_ALLOW = NVLD(Rs("MAXIMUM_ALLOW"), 0)
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PIG_FLAG = NVLS(Rs("PIG_FLAG"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      UNIT_WEIGHT = NVLD(Rs("UNIT_WEIGHT"), 0)
      BARCODE_NO = NVLS(Rs("BARCODE_NO"), "")
      BILL_DESC = NVLS(Rs("BILL_DESC"), "")
      WEIGHT_PER_PACK = NVLD(Rs("WEIGHT_PER_PACK"), 0)
      PARCEL_TYPE = NVLI(Rs("PARCEL_TYPE"), -1)
'      OLD_PART_ID = NVLI(Rs("OLD_PART_ID"), -1)
   ElseIf Ind = 2 Then
      JOB_DATE = NVLS(Rs("JOB_DATE"), "")
   ElseIf Ind = 3 Then
      MODIFY_DATE = InternalDateToDate(NVLS(Rs("MD"), ""))
      JOB_NO = NVLS(Rs("JOB_NO"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), -1)
   ElseIf Ind = 4 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
      PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
      PART_GROUP_NO = NVLS(Rs("PART_GROUP_NO"), "")
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
                     "   MAXIMUM_ALLOW, " & _
                     "   PART_TYPE, " & _
                     "   PIG_FLAG, " & _
                     "   PART_DESC, " & _
                     "   PIG_TYPE, " & _
                     "   UNIT_WEIGHT, " & _
                     "   BARCODE_NO, " & _
                     "   BILL_DESC, " & _
                     "   WEIGHT_PER_PACK, " & _
                     "   PARCEL_TYPE, " & _
                     "   OLD_PART_ID, " & _
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
                         "'" & ChangeQuote(Trim(MAXIMUM_ALLOW)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PIG_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(PIG_TYPE)) & "'," & _
                         UNIT_WEIGHT & "," & _
                         "'" & ChangeQuote(Trim(BARCODE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(BILL_DESC)) & "'," & _
                         WEIGHT_PER_PACK & "," & _
                         PARCEL_TYPE & "," & _
                         OLD_PART_ID & "," & _
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
                     "   MAXIMUM_ALLOW = '" & MAXIMUM_ALLOW & "', " & _
                     "   PART_TYPE = " & TempID1 & ", " & _
                     "   PIG_FLAG = '" & ChangeQuote(Trim(PIG_FLAG)) & "', " & _
                     "   PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "', " & _
                     "   PIG_TYPE = '" & ChangeQuote(Trim(PIG_TYPE)) & "', " & _
                     "   UNIT_WEIGHT = " & UNIT_WEIGHT & ", " & _
                     "   BARCODE_NO = '" & ChangeQuote(Trim(BARCODE_NO)) & "', " & _
                     "   BILL_DESC = '" & ChangeQuote(Trim(BILL_DESC)) & "', " & _
                     "   WEIGHT_PER_PACK = " & WEIGHT_PER_PACK & ", " & _
                     "   PARCEL_TYPE = " & PARCEL_TYPE & ", " & _
                     "   OLD_PART_ID = " & OLD_PART_ID & ", " & _
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
   Set HumidRates = Nothing
End Sub
Public Sub GeneratePIHeader(FileID As Long)
Dim TempStr As String
   
   TempStr = "PI;"
   TempStr = TempStr & PART_NO & ";"
   TempStr = TempStr & UNIT_COUNT & ";"
   TempStr = TempStr & MINIMUM_ALLOW & ";"
   
   TempStr = TempStr & MAXIMUM_ALLOW & ";"
   TempStr = TempStr & PART_TYPE & ";"
   TempStr = TempStr & PIG_FLAG & ";"
   TempStr = TempStr & PART_DESC & ";"
   
   TempStr = TempStr & UNIT_WEIGHT & ";"
   TempStr = TempStr & BARCODE_NO & ";"
   TempStr = TempStr & BILL_DESC & ";"
   TempStr = TempStr & WEIGHT_PER_PACK & ";"
   TempStr = TempStr & PARCEL_TYPE & ";"
   
   Print #FileID, TempStr
End Sub





