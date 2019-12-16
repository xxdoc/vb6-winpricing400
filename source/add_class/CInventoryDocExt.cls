VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryDocExt"
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

Public INVENTORY_DOC_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_DESC As String
Public BILL_NO As String
Public DO_NO As String
Public TRUCK_NO As String
Public SUPPLIER_ID As Long
Public DELIVERY_ID As Long
Public DELIVERY_FEE As Double
Public SENDER_NAME As String
Public RECEIVE_NAME As String
Public DOCUMENT_TYPE As Long
Public EMP_ID As Long
Public COMMIT_FLAG As String
Public SALE_FLAG As String
Public REASON_ID As Long
Public QUE_NO As String
Public ENTRY_DATE As Date
Public EXIT_DATE As Date
Public EXCEPTION_FLAG As String
Public DEPARTMENT_ID As Long
Public ADJUST_FLAG As String
Public Credit As Double


Public OLD_COMMIT_FLAG As String
Public PART_NO As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public RESPONSE_NAME As String
Public RESPONSE_LNAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DEPARTMENT_NAME As String
Public PR_NO As String

Public ImportItems As Collection
Public ExportItems As Collection
Public ImportExports As Collection
Public TransferItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   DOCUMENT_DATE = -1
   COMMIT_FLAG = "N"
   SALE_FLAG = "N"
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   Set ImportExports = New Collection
   Set TransferItems = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT ADR.*, NM1.LONG_NAME SUPPLIER_NAME, NM3.LONG_NAME RESPONSE_NAME, NM3.LAST_NAME RESPONSE_LNAME, SP.SUPPLIER_CODE, LO.LAY_OUT_NAME DEPARTMENT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM INVENTORY_DOC_EXT ADR "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = ADR.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (ADR.DEPARTMENT_ID = LO.LAY_OUT_ID) "
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
      
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Trim(DOCUMENT_NO) <> "" Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
      
   If DOCUMENT_DATE > -1 Then
      TempDate = Trim(DateToStringInt(DOCUMENT_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If Len(PART_NO) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((SELECT COUNT(*) FROM LOT_ITEM LI, PART_ITEM PI1 WHERE (LI.PART_ITEM_ID = PI1.PART_ITEM_ID) AND (LI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI1.PART_NO = '" & PART_NO & "')) > 0) "
      Else
         WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM LOT_ITEM LI, PART_ITEM PI1 WHERE (LI.PART_ITEM_ID = PI1.PART_ITEM_ID) AND (LI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI1.PART_NO = '" & PART_NO & "')) > 0) "
      End If
   End If
            
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY DOCUMENT_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY DOCUMENT_DATE "
   Else
      OrderByStr = "ORDER BY INVENTORY_DOC_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
   DELIVERY_ID = NVLI(Rs("DELIVERY_ID"), 0)
   DELIVERY_FEE = NVLD(Rs("DELIVERY_FEE"), 0)
   DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
   RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
   BILL_NO = NVLS(Rs("BILL_NO"), "")
   DO_NO = NVLS(Rs("DO_NO"), "")
   DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
   DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
   OLD_COMMIT_FLAG = COMMIT_FLAG
   SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
   REASON_ID = NVLI(Rs("REASON_ID"), -1)
   QUE_NO = NVLS(Rs("QUE_NO"), "")
   ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
   EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
   EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
   DEPARTMENT_ID = NVLI(Rs("DEPARTMENT_ID"), -1)
   ADJUST_FLAG = NVLS(Rs("ADJUST_FLAG"), "")
   Credit = NVLD(Rs("CREDIT"), 0)
   PR_NO = NVLS(Rs("PR_NO"), "")
   
   SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
   SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
   RESPONSE_NAME = NVLS(Rs("RESPONSE_NAME"), "")
   RESPONSE_LNAME = NVLS(Rs("RESPONSE_LNAME"), "")
   DEPARTMENT_NAME = NVLS(Rs("DEPARTMENT_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String

   If ADJUST_FLAG = "" Then
      ADJUST_FLAG = "N"
   End If
   
   TempID1 = SUPPLIER_ID
   If SUPPLIER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = DELIVERY_ID
   If DELIVERY_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = EMP_ID
   If EMP_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = REASON_ID
   If REASON_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = DEPARTMENT_ID
   If DEPARTMENT_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("INVENTORY_DOC_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO INVENTORY_DOC " & _
                     "( " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   SUPPLIER_ID, " & _
                     "   DELIVERY_ID, " & _
                     "   DOCUMENT_DESC, " & _
                     "   RECEIVE_NAME, " & _
                     "   BILL_NO, " & _
                     "   DO_NO, " & _
                     "   TRUCK_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   SENDER_NAME, " & _
                     "   DELIVERY_FEE, " & _
                     "   EMP_ID, " & _
                     "   COMMIT_FLAG, " & _
                     "   SALE_FLAG, " & _
                     "   REASON_ID, "
         SQL1 = SQL1 & _
                     "   QUE_NO, " & _
                     "   ENTRY_DATE, " & _
                     "   EXIT_DATE, " & _
                     "   EXCEPTION_FLAG, " & _
                     "   DEPARTMENT_ID, " & _
                     "   ADJUST_FLAG, " & _
                     "   CREDIT, " & _
                     "   PR_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(RECEIVE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(BILL_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DO_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(TRUCK_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         DOCUMENT_TYPE & "," & _
                         "'" & ChangeQuote(Trim(SENDER_NAME)) & "'," & _
                         DELIVERY_FEE & "," & _
                         TempID3 & "," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(SALE_FLAG)) & "'," & _
                         TempID4 & "," & _
                         "'" & ChangeQuote(Trim(QUE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "',"
            SQL1 = SQL1 & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EXIT_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                         TempID5 & "," & _
                         "'" & ChangeQuote(Trim(ADJUST_FLAG)) & "'," & _
                         Credit & "," & _
                         "'" & ChangeQuote(Trim(PR_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         INVENTORY_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_DOC SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   SUPPLIER_ID = " & TempID1 & ", " & _
                     "   DELIVERY_ID = " & TempID2 & ", " & _
                     "   DOCUMENT_DESC = '" & ChangeQuote(Trim(DOCUMENT_DESC)) & "', " & _
                     "   RECEIVE_NAME = '" & ChangeQuote(Trim(RECEIVE_NAME)) & "', " & _
                     "   BILL_NO = '" & ChangeQuote(Trim(BILL_NO)) & "', " & _
                     "   DO_NO = '" & ChangeQuote(Trim(DO_NO)) & "', " & _
                     "   TRUCK_NO = '" & ChangeQuote(Trim(TRUCK_NO)) & "', " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "   SENDER_NAME = '" & ChangeQuote(Trim(SENDER_NAME)) & "', " & _
                     "   DELIVERY_FEE = '" & DELIVERY_FEE & "', " & _
                     "   EMP_ID = " & TempID3 & ", " & _
                     "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   SALE_FLAG = '" & ChangeQuote(Trim(SALE_FLAG)) & "', " & _
                     "   REASON_ID = " & TempID4 & ", " & _
                     "   QUE_NO = '" & ChangeQuote(Trim(QUE_NO)) & "', " & _
                     "   ENTRY_DATE = '" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "', " & _
                     "   EXIT_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXIT_DATE))) & "', " & _
                     "   EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "   DEPARTMENT_ID = " & TempID5 & ", "
            SQL1 = SQL1 & _
                     "   ADJUST_FLAG = '" & ChangeQuote(Trim(ADJUST_FLAG)) & "', " & _
                     "   CREDIT = " & Credit & ", " & _
                     "   PR_NO = '" & ChangeQuote(Trim(PR_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUBLOT_ITEM WHERE LOT_ITEM_ID IN (SELECT LOT_ITEM_ID FROM LOT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM LOT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
            
      SQL1 = "DELETE FROM INVENTORY_DOC WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
End Function

Public Function PatchDocDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set ImportExports = Nothing
   Set TransferItems = Nothing
End Sub


