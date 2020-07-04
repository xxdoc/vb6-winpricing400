VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Conn As ADODB.Connection
Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function IsPatch(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim ItemCount As Long
Dim Result As Boolean

   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERSION = '" & VersionStr & "' "
   Call glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, TempRs, glbErrorLog)

   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch = Result
   Exit Function
   
ErrorHandler:
   IsPatch = False
End Function

Public Function Patch_1_0_12_2() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_2"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD TOTAL_ACTUAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE IMPORT_ITEM ADD TOTAL_INCLUDE_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE IMPORT_ITEM SET TOTAL_ACTUAL_PRICE = TOTAL_PRICE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE IMPORT_ITEM SET TOTAL_INCLUDE_PRICE = TOTAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE IMPORT_ITEM DROP TOTAL_PRICE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_2 = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_3() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_3"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE IMPORT_ITEM ADD TOTAL_ACTUAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_3 = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_9_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_9_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE DO_ITEM ADD FROM_PERIOD DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD TO_PERIOD DATE_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_9_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_10_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_10_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE PROCESS "
   SQL = SQL & "( "
   SQL = SQL & "PROCESS_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "PROCESS_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "PROCESS_NAME          CODE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PROCESS ADD CONSTRAINT PROCESS_ID_PK PRIMARY KEY (PROCESS_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR PROCESS_SEQ "
   m_Conn.Execute (SQL)


   SQL = "CREATE TABLE MACHINE "
   SQL = SQL & "( "
   SQL = SQL & "MACHINE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "MACHINE_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "MACHINE_NAME          CODE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE MACHINE ADD CONSTRAINT MACHINE_ID_PK PRIMARY KEY (MACHINE_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR MACHINE_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_10_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_11_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_11_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_11_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_12_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_12_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_12_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_13_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_13_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE JOB "
   SQL = SQL & "( "
   SQL = SQL & "JOB_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "JOB_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "JOB_DATE          DATE_TYPE NOT NULL, "
   SQL = SQL & "JOB_DESC          DESC_TYPE , "
   SQL = SQL & "BATCH_NO          CODE_TYPE , "
   SQL = SQL & "PROCESS_ID      ID_TYPE , "
   SQL = SQL & "APPROVED_BY      ID_TYPE , "
   SQL = SQL & "RESPONSE_BY      ID_TYPE , "
   SQL = SQL & "START_DATE      DATE_TYPE , "
   SQL = SQL & "FINISH_DATE      DATE_TYPE , "
   SQL = SQL & "SUCCESS_FLAG      FLAG_TYPE , "
   SQL = SQL & "INVENTORY_DOC_ID      ID_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE JOB_INOUT "
   SQL = SQL & "( "
   SQL = SQL & "JOB_INOUT_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "JOB_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "LOCATION_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "LINK_ID           ID_TYPE , "
   SQL = SQL & "TX_TYPE          CODE_TYPE , "
   SQL = SQL & "TX_AMOUNT          MONEY_TYPE , "
   SQL = SQL & "SERIAL_NUMBER          NAME_TYPE , "
   SQL = SQL & "INOUT_REF          CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE JOB_RESOURCE "
   SQL = SQL & "( "
   SQL = SQL & "JOB_RESOURCE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "JOB_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "MACHINE_ID           ID_TYPE , "
   SQL = SQL & "MACHINE_ID_HOUR           TIME_TYPE , "
   SQL = SQL & "MACHINE_ID_HOURN           TIME_TYPE , "
   SQL = SQL & "EMP_ID          ID_TYPE , "
   SQL = SQL & "EMP_ID_HOUR          TIME_TYPE , "
   SQL = SQL & "EMP_ID_HOURN          TIME_TYPE , "
   SQL = SQL & "OCCUPY_DATE          DATE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   
   m_Conn.Execute (SQL)
   SQL = "CREATE GENERATOR JOB_SEQ "
   m_Conn.Execute (SQL)
  SQL = "CREATE GENERATOR JOB_INOUT_SEQ "
   m_Conn.Execute (SQL)
 SQL = "CREATE GENERATOR JOB_RESOURCE_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB ADD CONSTRAINT JOB_ID_PK PRIMARY KEY (JOB_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD CONSTRAINT JOB_INOUT_ID_PK PRIMARY KEY (JOB_INOUT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_RESOURCE ADD CONSTRAINT JOB_RESOURCE_ID_PK PRIMARY KEY (JOB_RESOURCE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE JOB ADD CONSTRAINT PROCESS_ID_FK5 FOREIGN KEY (PROCESS_ID) REFERENCES PROCESS"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD CONSTRAINT APPROVED_BY_FK FOREIGN KEY (APPROVED_BY) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB ADD CONSTRAINT RESPONSE_BY_FK FOREIGN KEY (RESPONSE_BY) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB ADD CONSTRAINT INVENTORY_DOC_ID_FK FOREIGN KEY (INVENTORY_DOC_ID) REFERENCES INVENTORY_DOC"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD CONSTRAINT JOB_ID_FK FOREIGN KEY (JOB_ID) REFERENCES JOB"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD CONSTRAINT PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD CONSTRAINT LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION"
   m_Conn.Execute (SQL)

   SQL = "   ALTER TABLE JOB_RESOURCE ADD CONSTRAINT JOB_ID_FK2 FOREIGN KEY (JOB_ID) REFERENCES JOB"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_RESOURCE ADD CONSTRAINT MACHINE_ID_FK FOREIGN KEY (MACHINE_ID) REFERENCES MACHINE"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_RESOURCE ADD CONSTRAINT EMP_ID_FK2 FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)
   
   SQL = "DROP TABLE FORMULA_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "DROP TABLE FORMULA_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "DROP TABLE FORMULA_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "DROP TABLE FORMULA "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE FORMULA "
   SQL = SQL & "( "
   SQL = SQL & "FORMULA_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "FORMULA_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "FORMULA_TYPE           ID_TYPE , "
   SQL = SQL & "FORMULA_DATE           DATE_TYPE , "
   SQL = SQL & "FORMULA_DESC           DESC_TYPE , "
   SQL = SQL & "APPROVED_BY          ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID          ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

 SQL = "CREATE TABLE FORMULA_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "FORMULA_ITEM_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "FORMULA_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID           ID_TYPE , "
   SQL = SQL & "ITEM_PERCENT           MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE FORMULA_TYPE "
   SQL = SQL & "( "
   SQL = SQL & "FORMULA_TYPE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "FORMULA_TYPE_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "FORMULA_TYPE_NAME           CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE FORMULA ADD CONSTRAINT FORMULA_ID_PK PRIMARY KEY (FORMULA_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_ID_PK PRIMARY KEY (FORMULA_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA_TYPE ADD CONSTRAINT FORMULA_TYPE_ID_PK PRIMARY KEY (FORMULA_TYPE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA ADD CONSTRAINT FORMULA_FMR_TYPE_FK FOREIGN KEY (FORMULA_TYPE) REFERENCES FORMULA_TYPE"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA ADD CONSTRAINT FORMULA_EMP_ID_FK FOREIGN KEY (APPROVED_BY) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE FORMULA ADD CONSTRAINT FORMULA_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_FORMULA_ID_FK FOREIGN KEY (FORMULA_ID) REFERENCES FORMULA"
   m_Conn.Execute (SQL)
  
  SQL = "ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_PART_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM"
 m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_11_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_14_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_14_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB ADD JOB_DOC_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE JOB_INOUT ADD FROM_FORMULA ID_TYPE "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE JOB_INOUT ADD CONSTRAINT JOB_INOUT_FML_FK1 FOREIGN KEY (FROM_FORMULA) REFERENCES FORMULA "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE JOB_RESOURCE ADD OCCUPY_INTERVAL MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE FORMULA_ITEM ADD FROM_FORMULA ID_TYPE "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE FORMULA_ITEM ADD LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_FML_FK1 FOREIGN KEY (FROM_FORMULA) REFERENCES FORMULA "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE FORMULA_ITEM ADD CONSTRAINT FORMULA_ITEM_LOCATION_FK1 FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_14_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_15_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_15_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB_INOUT ADD AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA ADD LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA ADD CONSTRAINT FORMULA_LOCATION_FK1 FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR QUOATATION_NUMBER_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_15_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function


Public Function Patch_1_0_12_15_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_15_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE CURRENCY "
   SQL = SQL & "( "
   SQL = SQL & "CURRENCY_EXC_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "CURRENCY_EXC_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "EXCHANGE_DATE          DATE_TYPE NOT NULL, "
   SQL = SQL & "COUNTRY_ID1      ID_TYPE NOT NULL, "
   SQL = SQL & "COUNTRY_ID2      ID_TYPE NOT NULL, "
   SQL = SQL & "COEFFICIENT      MONEY_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE PKGLST_ITEM"
   SQL = SQL & "( "
   SQL = SQL & "PKGLST_ITEM_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "BILLING_DOC_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "PKG_NUMBER           CODE_TYPE NOT NULL, "
   SQL = SQL & "DESCRIPTION           DESC_TYPE , "
   SQL = SQL & "QUANTITY           MONEY_TYPE , "
   SQL = SQL & "MEASURE          MONEY_TYPE , "
   SQL = SQL & "NET_WEIGHT          MONEY_TYPE , "
   SQL = SQL & "GROSS_WEIGHT          MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CURRENCY_SEQ "
   m_Conn.Execute (SQL)
  SQL = "CREATE GENERATOR PKGLST_ITEM_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CURRENCY ADD CONSTRAINT CURRENCY_ID_PK PRIMARY KEY (CURRENCY_EXC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PKGLST_ITEM ADD CONSTRAINT PKGLST_ITEM_ID_PK PRIMARY KEY (PKGLST_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   
   SQL = " ALTER TABLE CURRENCY ADD CONSTRAINT CURRENCY_COUNTRY_ID_FK FOREIGN KEY (COUNTRY_ID1) REFERENCES COUNTRY"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CURRENCY ADD CONSTRAINT CURRENCY_COUNTRY_ID_FK2 FOREIGN KEY (COUNTRY_ID2) REFERENCES COUNTRY"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PKGLST_ITEM ADD CONSTRAINT PKGLST_ITEM_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_15_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_16_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_16_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE  EMP_CHILD ADD PHONE PHONE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EMP_DEPENDENCY ADD PHONE PHONE_TYPE "
   m_Conn.Execute (SQL)
   
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_16_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function


Public Function Patch_1_0_12_17_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_17_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

     SQL = "DROP TABLE CURRENCY "
     m_Conn.Execute (SQL)

    SQL = "CREATE TABLE MONEY_FAMILY "
    SQL = SQL & "( "
    SQL = SQL & "MONEY_FAMILY_ID         ID_TYPE NOT NULL, "
    SQL = SQL & "MONEY_FAMILY_NO       CODE_TYPE NOT NULL, "
    SQL = SQL & "MONEY_FAMILY_NAME       CODE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
    SQL = SQL & ") "
    m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR MONEY_FAMILY_SEQ "
   m_Conn.Execute (SQL)

    SQL = "ALTER TABLE MONEY_FAMILY ADD CONSTRAINT MONEY_FAMILY_ID_PK PRIMARY KEY (MONEY_FAMILY_ID) "
   m_Conn.Execute (SQL)

    SQL = "CREATE TABLE CURRENCY "
    SQL = SQL & "( "
    SQL = SQL & "CURRENCY_EXC_ID      ID_TYPE NOT NULL, "
    SQL = SQL & "CURRENCY_EXC_NO           CODE_TYPE NOT NULL, "
    SQL = SQL & "EXCHANGE_DATE          DATE_TYPE NOT NULL, "
    SQL = SQL & "MONEY_FAMILY_ID1      ID_TYPE NOT NULL, "
    SQL = SQL & "MONEY_FAMILY_ID2      ID_TYPE NOT NULL, "
    SQL = SQL & "COEFFICIENT      MONEY_TYPE NOT NULL, "
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
    SQL = SQL & ") "
    m_Conn.Execute (SQL)

 SQL = "ALTER TABLE CURRENCY ADD CONSTRAINT CURRENCY_ID_PK PRIMARY KEY (CURRENCY_EXC_ID) "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE CURRENCY ADD CONSTRAINT MONEY_FAMILY_ID_FK FOREIGN KEY (MONEY_FAMILY_ID1) REFERENCES MONEY_FAMILY "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE CURRENCY ADD CONSTRAINT MONEY_FAMILY_ID_FK2 FOREIGN KEY (MONEY_FAMILY_ID2) REFERENCES MONEY_FAMILY "
   m_Conn.Execute (SQL)

 SQL = "CREATE TABLE PARAMETER_PROCESS "
    SQL = SQL & "( "
    SQL = SQL & "PARAMETER_PROCESS_ID         ID_TYPE NOT NULL, "
    SQL = SQL & "PARAMETER_PROCESS_NO       CODE_TYPE NOT NULL, "
    SQL = SQL & "PARAMETER_PROCESS_NAME       CODE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
    SQL = SQL & ") "
    m_Conn.Execute (SQL)

SQL = "CREATE GENERATOR PARAMETER_PROCESS_SEQ "
   m_Conn.Execute (SQL)

    SQL = "ALTER TABLE PARAMETER_PROCESS ADD CONSTRAINT PARAMETER_PROCESS_ID_PK PRIMARY KEY (PARAMETER_PROCESS_ID); "
   m_Conn.Execute (SQL)

    SQL = "CREATE TABLE PARAMETER_ITEM "
    SQL = SQL & "( "
    SQL = SQL & "PARAMETER_ITEM_ID        ID_TYPE NOT NULL, "
    SQL = SQL & "PARAMETER_ID        ID_TYPE NOT NULL, "
    SQL = SQL & "PROCESS_ID           ID_TYPE NOT NULL, "
    SQL = SQL & "SELECT_FLAG           FLAG_TYPE, "
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
    SQL = SQL & ") "
    m_Conn.Execute (SQL)

  SQL = "CREATE GENERATOR PARAMETER_ITEM_SEQ "
   m_Conn.Execute (SQL)

 SQL = "ALTER TABLE PARAMETER_ITEM ADD CONSTRAINT PARAMETER_ITEM_ID_PK PRIMARY KEY (PARAMETER_ITEM_ID) "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE PARAMETER_ITEM ADD CONSTRAINT PARAMETER_ITEM_PARAMETER_FK FOREIGN KEY (PARAMETER_ID) REFERENCES PARAMETER_PROCESS "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE PARAMETER_ITEM ADD CONSTRAINT PARAMETER_ITEM_PROCESS_FK FOREIGN KEY (PROCESS_ID) REFERENCES PROCESS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_17_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_18_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_18_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

    SQL = "CREATE TABLE JOB_PARAMETER "
    SQL = SQL & "( "
    SQL = SQL & "JOB_PARAMETER_ID        ID_TYPE    NOT NULL, "
    SQL = SQL & "JOB_ID              ID_TYPE    NOT NULL, "
    SQL = SQL & "PARAMETER_PROCESS_ID        ID_TYPE    NOT NULL, "
    SQL = SQL & "JOB_PARAMETER_DESC             DESC_TYPE, "
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE  NOT NULL "
    SQL = SQL & ") "

    m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR JOB_PARAMETER_SEQ "
   m_Conn.Execute (SQL)

 SQL = "ALTER TABLE JOB_PARAMETER ADD CONSTRAINT JOB_PARAMETER_ID_PK PRIMARY KEY (JOB_PARAMETER_ID) "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE JOB_PARAMETER ADD CONSTRAINT PARAMETER_PROCESS_JOB_ID_FK FOREIGN KEY (JOB_ID) REFERENCES JOB "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE JOB_PARAMETER ADD CONSTRAINT PARAMETER_PROCESS_ID_FK FOREIGN KEY (PARAMETER_PROCESS_ID) REFERENCES PARAMETER_PROCESS "
   m_Conn.Execute (SQL)

   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_18_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_19_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_19_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE UNIT ADD PERIOD_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_19_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_19_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_19_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)


   SQL = "CREATE GENERATOR EMP_CHILD_SEQ "
   m_Conn.Execute (SQL)




   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_19_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_20_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_20_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
    SQL = "CREATE TABLE FORMULA_VARIABLE "
    SQL = SQL & "( "
    SQL = SQL & "   FORMULA_VARIABLE_ID   ID_TYPE NOT NULL,"
    SQL = SQL & "   FORMULA_ID            ID_TYPE NOT NULL,"
    SQL = SQL & "   VARIABLE_ID           ID_TYPE NOT NULL,"
    SQL = SQL & "   VARIABLE_VALUE        MONEY_TYPE NOT NULL,"
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE  NOT NULL "
    SQL = SQL & ") "
    m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR FORMULA_VARIABLE_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FORMULA_VARIABLE ADD CONSTRAINT FORMULA_VARIABLE_ID_PK PRIMARY KEY (FORMULA_VARIABLE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA_VARIABLE ADD CONSTRAINT FORMULA_VARIABLE_FOM_ID_FK FOREIGN KEY (FORMULA_ID) REFERENCES FORMULA "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FORMULA_ITEM ADD AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FORMULA ADD RMC MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FORMULA ADD PMC MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FORMULA_TYPE ADD INTERMEDIAT_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR ESTIMATE_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_20_seub = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function


Public Function Patch_1_0_12_20_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_20_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
   SQL = "CREATE GENERATOR JOB_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_20_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_21_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_21_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE DO_ITEM ADD DOLLAR MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD COUNTRY_CURRENCY2 ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD COUNTRY_CURRENCY1 ID_TYPE "
   m_Conn.Execute (SQL)

    SQL = "CREATE TABLE MEMO_BANK "
    SQL = SQL & "( "
    SQL = SQL & "MEMO_BANK_ID            ID_TYPE    NOT NULL, "
    SQL = SQL & "BILLING_DOC_ID     ID_TYPE    NOT NULL, "
    SQL = SQL & "MEMO_BANK_NO            CODE_TYPE    NOT NULL, "
    SQL = SQL & "DESCRIPTION        DESC_TYPE, "
     SQL = SQL & "MEMO_BANK_DATE          DATE_TYPE, "
      SQL = SQL & "COEFFICIENT        MONEY_TYPE, "
       SQL = SQL & "AMOUNT_THAI        MONEY_TYPE, "
        SQL = SQL & "AMOUNT_OTHER       MONEY_TYPE, "
         SQL = SQL & "TAX                MONEY_TYPE, "
    SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
    SQL = SQL & "MODIFY_BY ID_TYPE  NOT NULL "
    SQL = SQL & ") "

    m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR MEMO_BANK_SEQ "
   m_Conn.Execute (SQL)

 SQL = "ALTER TABLE MEMO_BANK ADD CONSTRAINT MEMO_BANK_ID_PK PRIMARY KEY (MEMO_BANK_ID) "
   m_Conn.Execute (SQL)
 SQL = "ALTER TABLE MEMO_BANK ADD CONSTRAINT MEMO_BANK_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_21_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_22_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_22_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE MEMO_BANK ADD EXCHANGE_DATE DATE_TYPE "
   m_Conn.Execute (SQL)

   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_22_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function


Public Function Patch_1_0_12_21_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_21_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   SQL = "CREATE TABLE BANK"
   SQL = SQL & "("
   SQL = SQL & "   BANK_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   BANK_NO          CODE_TYPE NOT NULL,"
   SQL = SQL & "   BANK_NAME         NAME_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BANK ADD CONSTRAINT BANK_ID_PK PRIMARY KEY (BANK_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BALANCE_ACCUM ADD CONSTRAINT BALANCE_ACCUM_ID_PK PRIMARY KEY (BALANCE_ACCUM_ID) "
   m_Conn.Execute (SQL)
            
   SQL = "CREATE GENERATOR BANK_SEQ "
   m_Conn.Execute (SQL)
            
   
   SQL = "CREATE TABLE BANK_BRANCH"
   SQL = SQL & "("
   SQL = SQL & "   BBRANCH_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   BBRANCH_NO          CODE_TYPE NOT NULL,"
   SQL = SQL & "   BBRANCH_NAME         NAME_TYPE NOT NULL,"
   SQL = SQL & "   BANK_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BANK_BRANCH ADD CONSTRAINT BANK_BRANCH_ID_PK PRIMARY KEY (BBRANCH_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BANK_BRANCH ADD CONSTRAINT BANK_BRANCH_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR BBRANCH_SEQ "
   m_Conn.Execute (SQL)
   '====
   
   SQL = "CREATE TABLE PAYMENT"
   SQL = SQL & "("
   SQL = SQL & "      PAYMENT_ID         ID_TYPE NOT NULL,"
   SQL = SQL & "      PAYMENT_NO         CODE_TYPE NOT NULL,"
   SQL = SQL & "      PAYMENT_DATE       NAME_TYPE NOT NULL,"
   SQL = SQL & "      TO_BANK_ID         ID_TYPE,"
   SQL = SQL & "      TO_BANK_BRANCH     ID_TYPE,"
   SQL = SQL & "      ACCOUNT_NO         CODE_TYPE,"
   SQL = SQL & "      COMMIT_FLAG        FLAG_TYPE,"
   SQL = SQL & "      INTERNAL_FLAG        FLAG_TYPE,"
   SQL = SQL & "      TX_TYPE            CODE_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_ID_PK PRIMARY KEY (PAYMENT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_BANK_ID_FK FOREIGN KEY (TO_BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_BBRANCH_ID_FK FOREIGN KEY (TO_BANK_BRANCH) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR PAYMENT_SEQ "
   m_Conn.Execute (SQL)
               
   SQL = "CREATE TABLE PAYMENT_ITEM"
   SQL = SQL & "("
   SQL = SQL & "   PAYMENT_ITEM_ID    ID_TYPE NOT NULL,"
   SQL = SQL & "   PAYMENT_ID         ID_TYPE NOT NULL,"
   SQL = SQL & "   PAYMENT_TYPE       ID_TYPE,"
   SQL = SQL & "   CHECK_NO           CODE_TYPE,"
   SQL = SQL & "   BANK_ID            ID_TYPE,"
   SQL = SQL & "   BANK_BRANCH        ID_TYPE,"
   SQL = SQL & "   BANK_NOTE          DESC_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_ID_PK PRIMARY KEY (PAYMENT_ITEM_ID) "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_PMT_ID_FK FOREIGN KEY (PAYMENT_ID) REFERENCES PAYMENT "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_BBANK_ID_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PAYMENT_ITEM_SEQ "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_21_seub = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_22_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_22_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   SQL = "ALTER TABLE PAYMENT ADD ACCEPT_BY ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_ACCEPT_BY_FK FOREIGN KEY (ACCEPT_BY) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD TOTAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD PAY_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_22_seub = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_23_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_23_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   SQL = "CREATE TABLE JOB_VERIFY "
   SQL = SQL & "( "
   SQL = SQL & "JOB_VERIFY_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "JOB_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID          ID_TYPE NOT NULL, "
   SQL = SQL & "NOTE          DESC_TYPE, "
   SQL = SQL & "VERIFY_FLAG          FLAG_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
               
   SQL = "CREATE GENERATOR JOB_VERIFY_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_VERIFY ADD CONSTRAINT JOB_VERIFY_ID_PK PRIMARY KEY (JOB_VERIFY_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_VERIFY ADD CONSTRAINT JOB_VERIFY_JOB_ID_FK FOREIGN KEY (JOB_ID) REFERENCES JOB "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_VERIFY ADD CONSTRAINT JOB_VERIFY_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_23_seub = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_24_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_24_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                              
   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD BANK_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD BBRANCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CHECK_NO CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CHECK_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_BBRANCH_ID_FK FOREIGN KEY (BBRANCH_ID) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)
   
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_24_seub = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_25_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_25_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   
   SQL = "ALTER TABLE BILLING_DOC ADD VADILITY DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BILLING_DOC ADD DELIVERY DATE_TYPE"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_25_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_26_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_26_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_DESC DESC_TYPE "
   m_Conn.Execute (SQL)
     
  SQL = "ALTER TABLE BILLING_DOC ADD SHIPMENT DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BILLING_DOC ADD REFER_INV DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PACKING_OF DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD SHIP_NAME DATE_TYPE "
   m_Conn.Execute (SQL)
   
 SQL = "ALTER TABLE BILLING_DOC ADD SHIPPING_MARKS DATE_TYPE "
   m_Conn.Execute (SQL)


   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_26_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_27_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_27_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   
   SQL = "ALTER TABLE BILLING_DOC DROP SHIP_NAME "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC DROP SHIPPING_MARKS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD SHIP_NAME DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD SHIPPING_MARKS NAME_TYPE "
   m_Conn.Execute (SQL)
  
   SQL = "ALTER TABLE BILLING_DOC ADD SHIPON_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PKGLST_ITEM ADD SET_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_27_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_25_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_25_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                                 
   SQL = "UPDATE BILLING_DOC BD SET TOTAL_PRICE ="
   SQL = SQL & "("
   SQL = SQL & "   SELECT SUM(DI.TOTAL_PRICE) FROM DO_ITEM DI"
   SQL = SQL & "   WHERE Di.DO_ID = Bd.BILLING_DOC_ID"
   SQL = SQL & ")"
   SQL = SQL & "WHERE ((SELECT COUNT(*) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID)) > 0)"
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_25_seub = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_28_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_28_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE CURRENCY_EX "
   SQL = SQL & "( "
   SQL = SQL & "CURRENCY_EX_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "EXCHANGE_DATE          DATE_TYPE NOT NULL, "
   SQL = SQL & "US      MONEY_TYPE , "
   SQL = SQL & "EURO      MONEY_TYPE , "
   SQL = SQL & "YEN      MONEY_TYPE , "
   SQL = SQL & "SS      MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CURRENCY_EX_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CURRENCY_EX ADD CONSTRAINT CURRENCY_EX_ID_PK PRIMARY KEY (CURRENCY_EX_ID) "
   m_Conn.Execute (SQL)
   
    m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_28_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_29_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_29_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE MEMO_BANK ADD UNIT ID_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_29_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_30_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_30_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE DO_ITEM ADD QTYP MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PACK MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_30_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_30_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_30_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR QUOATATION_NUMBER_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_30_seub = True
End Function

Public Function Patch_1_0_12_31_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_31_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_PERCENT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD REF STRING_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CREDIT MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_31_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_32_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_32_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE MEMO_BANK ADD TICKET_TYPE CODE_TYPE "
   m_Conn.Execute (SQL)
         
    SQL = "ALTER TABLE MEMO_BANK ADD INVOICE_NO CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_BANK ADD CURRENCY_OTHER CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_BANK ADD RATIO MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_BANK ADD COST_BAHT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_BANK ADD TAX_PERCENT MONEY_TYPE "
   m_Conn.Execute (SQL)
         
    SQL = "ALTER TABLE MEMO_BANK ADD TAX_IN MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_BANK ADD VAT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_BANK ADD PAID_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_32_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_33_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_33_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR PKGLST_NUMBER_SEQ "
   m_Conn.Execute (SQL)
    
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_33_jill = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_41_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_41_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE EXTRACT_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "EXTRACT_ITEM_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "QUANTITY_EXTRACT_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "TOTAL_AMT      MONEY_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR EXTRACT_ITEM_SEQ "
   m_Conn.Execute (SQL)
  
   SQL = "ALTER TABLE EXTRACT_ITEM ADD CONSTRAINT EXTRACT_ITEM_ID_PK PRIMARY KEY (EXTRACT_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRACT_ITEM ADD CONSTRAINT EXTRACT_ITEM_QTE_ID_FK FOREIGN KEY (QUANTITY_EXTRACT_ID) REFERENCES QUANTITY_EXTRACT "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE EXTRACT_ITEM ADD CONSTRAINT EXTRACT_ITEM_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)


   SQL = "ALTER TABLE FORMULA_ITEM ADD REAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "UPDATE FORMULA_ITEM SET REAL_AMOUNT = ITEM_PERCENT "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_41_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_42_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_42_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR EXPORT_NUMBER_SEQ "
   m_Conn.Execute (SQL)
  
   SQL = "CREATE GENERATOR TRANSFER_NUMBER_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_42_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_77_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_77_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD TEMP_DO_NO CODE_TYPE "
   m_Conn.Execute (SQL)
        
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_77_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_78_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_78_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE HGROUP_ITEM DROP CONSTRAINT HGROUP_ITEM_HGROUP_ID_FK "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE HGROUP_ITEM DROP HOUSE_GROUP_ID "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE HGROUP_ITEM ADD PART_GROUP_ID ID_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE HGROUP_ITEM ADD CONSTRAINT HGROUP_ITEM_PGROUP_ID_FK FOREIGN KEY (PART_GROUP_ID) REFERENCES PART_GROUP "
   m_Conn.Execute (SQL)
        
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_78_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_79_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_79_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_79_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_80_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_80_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "CREATE TABLE COST_EXPENSE "
   SQL = SQL & "( "
   SQL = SQL & "COST_EXPENSE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "COST_PRODUCTION_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "EXPENSE_TYPE      ID_TYPE NOT NULL, "
   SQL = SQL & "PACKAGE_TYPE      ID_TYPE, "
   SQL = SQL & "RATIO_TYPE      ID_TYPE, "
   SQL = SQL & "EXPENSE_AMOUNT      MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE COST_EXPENSE ADD CONSTRAINT COST_EXPENSE_ID_ID_PK PRIMARY KEY (COST_EXPENSE_ID) "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE COST_EXPENSE ADD CONSTRAINT COST_EXPENSE_CSTPRD_ID_FK FOREIGN KEY (COST_PRODUCTION_ID) REFERENCES COST_PRODUCTION "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE COST_EXPENSE ADD CONSTRAINT COST_EXPENSE_EXP_ID_FK FOREIGN KEY (EXPENSE_TYPE) REFERENCES PARAMETER_PROCESS "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COST_EXPENSE_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD PARCEL_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD RAW_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD EXPENSE_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COSTPRD_ITEM ADD RAW_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_80_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_82_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_82_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_82_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_83_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_83_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "ALTER TABLE JOB ADD FORMULA_ID ID_TYPE "
   m_Conn.Execute (SQL)
        
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_83_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_85_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_85_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "ALTER TABLE RECEIPT_ITEM ADD PAID_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE RECEIPT_ITEM ADD DEBIT_CREDIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD PAID_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD DEBIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD CREDIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD DO_TOTAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD REVENUE_TOTAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD DOCUMENT_SUBTYPE ID_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_ID ID_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_PM_ID_FK FOREIGN KEY (PAYMENT_ID) REFERENCES PAYMENT "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD BANK_BRANCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
        
   SQL = "ALTER TABLE BILLING_DOC ADD BANK_NOTE DESC_TYPE "
   m_Conn.Execute (SQL)
        
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_85_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_86_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_86_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "ALTER TABLE LOT_ITEM ADD TOTAL_NEW_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
                
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_86_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_88_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_88_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "ALTER TABLE LOT_ITEM ADD BAG_RETURN FLAG_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE INVENTORY_DOC ADD ADJUST_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE LOT_ITEM ADD ACTUAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE LOT_ITEM ADD ACTUAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE LOT_ITEM ADD CURRENT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE LOT_ITEM ADD CURRENT_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE RECEIPT_ITEM ADD DRCR_REASON ID_TYPE "
   m_Conn.Execute (SQL)
                
   SQL = "ALTER TABLE RECEIPT_ITEM ADD CASH_DISCOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_88_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_89_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_89_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "ALTER TABLE CUSTOMER ADD EXP_CODE CODE_TYPE "
   m_Conn.Execute (SQL)
                
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_89_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_90_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_90_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "ALTER TABLE DO_ITEM ADD EXTRA_DISCOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
                
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "UPDATE PART_ITEM SET PARCEL_TYPE = 2 WHERE PART_TYPE = 21 "
   m_Conn.Execute (SQL)

   SQL = "UPDATE BILLING_DOC SET COMMIT_FLAG = 'Y' WHERE DOCUMENT_DATE <= '2006-02-28 23:5959' "
   m_Conn.Execute (SQL)

   Set BD = New CBillingDoc
   Call glbDaily.PatchExtraDiscount(BD, IsOK, False, glbErrorLog)
   Set BD = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_90_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_92_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_92_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
        
   SQL = "CREATE INDEX BD_CHECK_NO_IDX ON BILLING_DOC (CHECK_NO) "
   m_Conn.Execute (SQL)
                
   SQL = "CREATE INDEX BD_PMT_TYPE_IDX ON PAYMENT_ITEM (PAYMENT_TYPE) "
   m_Conn.Execute (SQL)
                
   SQL = "CREATE INDEX PMI_CHECK_NO_IDX ON PAYMENT_ITEM (CHECK_NO) "
   m_Conn.Execute (SQL)
                
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_92_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_93_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_93_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD FEE_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
                                
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_93_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_94_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_94_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM ADD NEED_TOTAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
                                
   SQL = "ALTER TABLE LOT_ITEM ADD NEED_TOTAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD NEED_AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_94_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_95_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_95_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD POST_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
                                   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE PAYMENT_ITEM SET POST_FLAG = 'Y' "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_95_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_96_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_96_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM ADD MANUAL_PRICE FLAG_TYPE "
   m_Conn.Execute (SQL)
                                   
   SQL = "ALTER TABLE LOT_ITEM ADD EXPENSE_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD CONSTRAINT LOT_ITEM_EXPENSE_TYPE_FK FOREIGN KEY (EXPENSE_TYPE) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE LOT_ITEM SET MANUAL_PRICE = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO DOCUMENT_PRIORITY (DOCUMENT_PRIORITY_ID, DOCUMENT_TYPE, AREA, PRIORITY1, PRIORITY2, CREATE_DATE, CREATE_BY, MODIFY_DATE, MODIFY_BY) VALUES (9, 22, 1, 3, 3, '2000', 1, '2000', 1) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO DOCUMENT_PRIORITY (DOCUMENT_PRIORITY_ID, DOCUMENT_TYPE, AREA, PRIORITY1, PRIORITY2, CREATE_DATE, CREATE_BY, MODIFY_DATE, MODIFY_BY) VALUES (10, 23, 1, 1, 1, '2000', 1, '2000', 1) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_96_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_97_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_97_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO DOCUMENT_PRIORITY (DOCUMENT_PRIORITY_ID, DOCUMENT_TYPE, AREA, PRIORITY1, PRIORITY2, CREATE_DATE, CREATE_BY, MODIFY_DATE, MODIFY_BY) VALUES (11, 4, 1, 999, 999, '2000', 1, '2000', 1) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_97_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_98_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_98_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MASTER_REF ADD TEMP_ID1 ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_REF ADD TEMP_ID2 ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD BANK_ACCOUNT ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CUSTOMER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_CUST_ID FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   Call glbDaily.PatchBankAccount(IsOK, False, glbErrorLog)
   
   SQL = "UPDATE PAYMENT PM SET BANK_ACCOUNT = (SELECT KEY_ID FROM MASTER_REF MR WHERE (MR.KEY_NAME = PM.ACCOUNT_NO) AND (MR.TEMP_ID1 = PM.TO_BANK_ID) AND (MR.TEMP_ID2 = PM.TO_BANK_BRANCH) AND (MR.MASTER_AREA = 3)) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_98_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_1_0_12_99_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim BD As CBillingDoc
Dim IsOK As Boolean

   VersionStr = "1_0_12_99_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PAYMENT ADD RECEIPT_NO CODE_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_99_seub = True
   Exit Function
   
'ErrorHandler:
'   If HasBegin Then
'      m_Conn.RollbackTrans
'   End If
End Function

Public Function Patch_2006_09_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  CASH_DOC "
   SQL = SQL & "( "
   SQL = SQL & "CASH_DOC_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_DATE           DATE_TYPE , "
   SQL = SQL & "BANK_BRANCH           ID_TYPE , "
   SQL = SQL & "BANK_ID           ID_TYPE , "
   SQL = SQL & "BANK_ACCOUNT           ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID           ID_TYPE , "
   SQL = SQL & "DOCUMENT_NO           CODE_TYPE , "
   SQL = SQL & "DOCUMENT_TYPE           ID_TYPE , "
   SQL = SQL & "EMP_ID           ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_ID_PK PRIMARY KEY (CASH_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_BBRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_BNK_ACCOUNT_FK FOREIGN KEY (BANK_ACCOUNT) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_EMP_ID_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE  CHEQUE "
   SQL = SQL & "( "
   SQL = SQL & "CHEQUE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "CHEQUE_NO           CODE_TYPE , "
   SQL = SQL & "CHEQUE_DATE           DATE_TYPE , "
   SQL = SQL & "EFFECTIVE_DATE           DATE_TYPE , "
   SQL = SQL & "CHEQUE_TYPE           ID_TYPE , "
   SQL = SQL & "CHEQUE_STATUS           ID_TYPE , "
   SQL = SQL & "BANK_ID           ID_TYPE , "
   SQL = SQL & "BANK_BRANCH           ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID           ID_TYPE , "
   SQL = SQL & "DIRECTION           ID_TYPE , "
   SQL = SQL & "CHEQUE_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "APAR_ID           ID_TYPE , "
   SQL = SQL & "BANK_FLAG           FLAG_TYPE , "
   SQL = SQL & "TX_DATE           DATE_TYPE , "
   SQL = SQL & "TX_NO CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_ID_PK PRIMARY KEY (CHEQUE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_BRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_CHEQUE_TYPE_FK FOREIGN KEY (CHEQUE_TYPE) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_ENTERPRISE_ID_FK FOREIGN KEY (ENTERPRISE_ID) REFERENCES ENTERPRISE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_APAR_ID_FK FOREIGN KEY (APAR_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE CASH_TRAN "
   SQL = SQL & "( "
   SQL = SQL & "CASH_TRAN_ID      ID_TYPE NOT NULL, "
    SQL = SQL & "BILLING_DOC_ID           ID_TYPE , "
   SQL = SQL & "CHECK_ID           ID_TYPE , "
   SQL = SQL & "AMOUNT           MONEY_TYPE , "
   SQL = SQL & "BANK_ID           ID_TYPE , "
    SQL = SQL & "BANK_BRANCH           ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID           ID_TYPE , "
   SQL = SQL & "TX_TYPE           CODE_TYPE , "
   SQL = SQL & "PAYMENT_TYPE           ID_TYPE , "
   SQL = SQL & "BANK_ACCOUNT           ID_TYPE , "
   SQL = SQL & "CASH_DOC_ID           ID_TYPE , "
   SQL = SQL & "FEE_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "NET_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "TX_DATE           DATE_TYPE , "
   SQL = SQL & "TX_NO           CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_ID_PK PRIMARY KEY (CASH_TRAN_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BANK_AMOUNT_FK FOREIGN KEY (BANK_ACCOUNT) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CHEQUE_FK FOREIGN KEY (CHECK_ID) REFERENCES CHEQUE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CASH_DOC_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CHEQUE_SEQ "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CASH_DOC_SEQ "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CASH_TRAN_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_21_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_09_22_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_22_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE CASH_TRAN ADD CUSTOMER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_22_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_09_23_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_23_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_23_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_09_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE CASH_DOC ADD  CUSTOMER_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_09_25_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_25_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE CHEQUE ADD  POST_FLAG FLAG_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_25_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_09_25_1_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_25_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE DO_ITEM ADD TRANSFER_WAGE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE DO_ITEM ADD STD_TRANSFER_CHARGE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE DOCUMENT_PRIORITY ADD TX_TYPE CODE_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE CUSTOMER_GRADE ADD MANUAL_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_25_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_09_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " UPDATE CHEQUE SET POST_FLAG = 'N' "
   m_Conn.Execute (SQL)
    
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_26_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_09_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  CASH_DOC_POST "
   SQL = SQL & "( "
   SQL = SQL & "CASH_DOC_POST_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "CASH_DOC_ID           ID_TYPE , "
   SQL = SQL & "BANK_BRANCH           ID_TYPE , "
   SQL = SQL & "BANK_ID           ID_TYPE , "
   SQL = SQL & "CHEQUE_ID           ID_TYPE , "
   SQL = SQL & "CHEQUE_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_ID_PK PRIMARY KEY (CASH_DOC_POST_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CASH_DOC_ID_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BBRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CHEQUE_FK FOREIGN KEY (CHEQUE_ID) REFERENCES CHEQUE  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CASH_DOC_POST_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_26_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_09_26_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_26_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE CHEQUE DROP CONSTRAINT CHEQUE_BANK_BRANCH_FK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE DROP CONSTRAINT CHEQUE_BANK_ID_FK "
   m_Conn.Execute (SQL)
      
    SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_BRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_26_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_09_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_09_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE BILLING_DOC ADD TOTAL_RCP MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_09_27_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_02_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_02_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  COST_RAW "
   SQL = SQL & "( "
   SQL = SQL & "COST_RAW_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "COST_EXPENSE_ID           ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID           ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COST_RAW ADD CONSTRAINT COST_RAW_PK PRIMARY KEY (COST_RAW_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE COST_RAW ADD CONSTRAINT COST_RAW_EXP_FK FOREIGN KEY (COST_EXPENSE_ID) REFERENCES COST_EXPENSE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE COST_RAW ADD CONSTRAINT COST_RAW_PART_ITEM_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COST_RAW_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_02_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_03_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_03_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
     
   SQL = "ALTER TABLE QUANTITY_EXTRACT ADD PROCESS_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE QUANTITY_EXTRACT ADD AREA ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXTRACT_ITEM ADD STD_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "DELETE FROM EXTRACT_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM QUANTITY_EXTRACT "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_03_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  PRTITEM_SET  "
   SQL = SQL & "( "
   SQL = SQL & "PRTITEM_SET_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MASTER_REF_ID                   ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PRTITEM_SET  ADD CONSTRAINT PRTITEM_SET_PK  PRIMARY KEY (PRTITEM_SET_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PRTITEM_SET  ADD CONSTRAINT PRTITEM_SET_MR_FK  FOREIGN KEY (MASTER_REF_ID) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PRTITEM_SET  ADD CONSTRAINT PRTITEM_SET_PART_ITEM_FK  FOREIGN KEY (PART_ITEM_ID ) REFERENCES PART_ITEM  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PRTITEM_SET_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_04_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_06_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_06_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  
   SQL = " ALTER TABLE RECEIPT_ITEM ADD PART_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD RETURN_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD LOCATION_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD PART_TYPE ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD RETURN_TOTAL_PRICE MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD RETURN_AVG_PRICE MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD AVG_WEIGHT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD FEATURE_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD RETURN_DISCOUNT_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD CONFIG_CODE CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD DISPLAY_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD COUNTRY_CURRENCY1 ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD COUNTRY_CURRENCY2 ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD WEIGHT_PER_PACK MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD PACK_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD PRICE_PER_PACK MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD DISCOUNT_PER_PACK MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD MANUAL_FLAG FLAG_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD MANUAL_CODE CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD MANUAL_NAME CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD RATE_TYPE ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD TRANSFER_WAGE MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD STD_TRANSFER_CHARGE MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD ITEM_DESC DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_06_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_09_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_09_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE RECEIPT_ITEM ADD LINK_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_09_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_09_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_09_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE RECEIPT_ITEM ADD AVG_PRICE MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_09_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_12_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_12_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE TABLE  SALE_ORDER  "
   SQL = SQL & "( "
   SQL = SQL & "SALE_ORDER_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "DO_ID                   ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "LOCATION_ID                    ID_TYPE , "
   SQL = SQL & "ITEM_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "TOTAL_WEIGHT                    MONEY_TYPE , "
   SQL = SQL & "AVG_WEIGHT                    MONEY_TYPE , "
   SQL = SQL & "TOTAL_PRICE                    MONEY_TYPE , "
   SQL = SQL & "AVG_PRICE                    MONEY_TYPE , "
   SQL = SQL & "GUI_ID                    ID_TYPE , "
   SQL = SQL & "PIG_STATUS                    ID_TYPE , "
   SQL = SQL & "LINK_ID                    ID_TYPE , "
   SQL = SQL & "CONFIG_CODE                    CODE_TYPE , "
   SQL = SQL & "SOC_ID                    ID_TYPE , "
   SQL = SQL & "FEATURE_ID                    ID_TYPE , "
   SQL = SQL & "ITEM_DESC                    DESC_TYPE , "
   SQL = SQL & "DEPOSIT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "DISCOUNT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "DISPLAY_ID                    ID_TYPE , "
   SQL = SQL & "PO_ID                    ID_TYPE , "
   SQL = SQL & "PO_NO                    CODE_TYPE , "
   SQL = SQL & "FROM_PERIOD                    DATE_TYPE , "
   SQL = SQL & "TO_PERIOD                    DATE_TYPE , "
   SQL = SQL & "DOLLAR                    MONEY_TYPE , "
   SQL = SQL & "COUNTRY_CURRENCY2                    ID_TYPE , "
   SQL = SQL & "COUNTRY_CURRENCY1                    ID_TYPE , "
   SQL = SQL & "QTYP                    MONEY_TYPE , "
   SQL = SQL & "PACK                    MONEY_TYPE , "
   SQL = SQL & "DISCOUNT_PERCENT                    MONEY_TYPE , "
   SQL = SQL & "WEIGHT_PER_PACK                    MONEY_TYPE , "
   SQL = SQL & "PACK_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "PRICE_PER_PACK                    MONEY_TYPE , "
   SQL = SQL & "DISCOUNT_PER_PACK                    MONEY_TYPE , "
   SQL = SQL & "MANUAL_FLAG                    FLAG_TYPE , "
   SQL = SQL & "MANUAL_CODE                    CODE_TYPE , "
   SQL = SQL & "MANUAL_NAME                    CODE_TYPE , "
   SQL = SQL & "RATE_TYPE                    ID_TYPE , "
   SQL = SQL & "EXTRA_DISCOUNT                    MONEY_TYPE , "
   SQL = SQL & "TRANSFER_WAGE                    MONEY_TYPE , "
   SQL = SQL & "STD_TRANSFER_CHARGE                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER  ADD CONSTRAINT SALE_ORDER_PK  PRIMARY KEY (SALE_ORDER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SALE_ORDER  ADD CONSTRAINT SALE_ORDER_DO_ID_FK  FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SALE_ORDER  ADD CONSTRAINT SALE_ORDER_PIG_STATUS_ID_FK  FOREIGN KEY (PIG_STATUS ) REFERENCES PRODUCT_STATUS  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SALE_ORDER_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_12_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE TABLE  CONFIG_DOC  "
   SQL = SQL & "( "
   SQL = SQL & "CONFIG_DOC_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CONFIG_DOC_TYPE                    ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID             ID_TYPE , "
   SQL = SQL & "PREFIX                    CODE_TYPE , "
   SQL = SQL & "CODE1                    CODE_TYPE , "
   SQL = SQL & "YEAR_TYPE                    ID_TYPE , "
   SQL = SQL & "CODE2                    CODE_TYPE , "
   SQL = SQL & "MONTH_TYPE                    ID_TYPE , "
   SQL = SQL & "CODE3                    CODE_TYPE , "
   SQL = SQL & "DIGIT_AMOUNT                    ID_TYPE , "
   SQL = SQL & "RUNNING_NO                    ID_TYPE , "
   SQL = SQL & "LAST_NO                    CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC  ADD CONSTRAINT CONFIG_DOC_PK  PRIMARY KEY (CONFIG_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CONFIG_DOC_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_13_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_16_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_16_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE BILLING_DOC ADD RUNNING_NO ID_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_16_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_18_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE CUSTOMER ADD CREDIT_LIMIT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_18_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
     
   m_Conn.BeginTrans
   HasBegin = True
   
   Call glbDaily.PatchDocumentPriority
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_19_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_19_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_19_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE TABLE  CUSTOMER_PICTURE "
   SQL = SQL & "( "
   SQL = SQL & "CUSTOMER_PICTURE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID                    ID_TYPE , "
   SQL = SQL & "CUSTOMER_PICTURE_TYPE             ID_TYPE , "
   SQL = SQL & "CUSTOMER_PICTURE_PATH                   DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_PICTURE  ADD CONSTRAINT CUSTOMER_PICTURE_PK  PRIMARY KEY (CUSTOMER_PICTURE_ID) "
   m_Conn.Execute (SQL)
      
   SQL = " ALTER TABLE CUSTOMER_PICTURE  ADD CONSTRAINT CUSTOMER_PICTURE_CUSTOMER_ID_FK  FOREIGN KEY (CUSTOMER_ID ) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CUSTOMER_PICTURE_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_19_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE BILLING_DOC ADD AGREEMENT_DATA DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILLING_DOC ADD AGREEMENT_FINANCE DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_20_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE BILLING_DOC ADD OLD_CREDIT_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_10_25_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_10_25_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE GENERATOR PICTURE_FILE_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_10_25_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_11_02_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_02_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE UC_RATE ADD LOG_RATE_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_02_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_11_15_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_15_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE TABLE  MEMO_NOTE  "
   SQL = SQL & "( "
   SQL = SQL & "MEMO_NOTE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MEMO_NOTE_DATE_CREATE                    DATE_TYPE , "
   SQL = SQL & "MEMO_NOTE_DATE_FINISH                    DATE_TYPE , "
   SQL = SQL & "MEMO_NOTE_DATE_FINISH_REAL                    DATE_TYPE , "
   SQL = SQL & "MEMO_NOTE_CREATE_BY                    ID_TYPE , "
   SQL = SQL & "MEMO_NOTE_CREATE_TO                    ID_TYPE , "
   SQL = SQL & "MEMO_NOTE_WARN                    FLAG_TYPE , "
   SQL = SQL & "MEMO_NOTE_TYPE                    ID_TYPE , "
   SQL = SQL & "MEMO_NOTE_STATUS                    ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID             ID_TYPE , "
   SQL = SQL & "MEMO_NOTE_SUBJECT                    DESC_TYPE , "
   SQL = SQL & "MEMO_NOTE_DESCRIPTION                    DESC_TYPE , "
   SQL = SQL & "MEMO_NOTE_RESOLUTION                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_ID_PK  PRIMARY KEY (MEMO_NOTE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_CREATE_BY_FK  FOREIGN KEY (MEMO_NOTE_CREATE_BY ) REFERENCES EMPLOYEE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_CREATE_TO_FK  FOREIGN KEY (MEMO_NOTE_CREATE_TO ) REFERENCES EMPLOYEE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_TYPE_FK  FOREIGN KEY (MEMO_NOTE_TYPE) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_STATUS_FK  FOREIGN KEY (MEMO_NOTE_STATUS) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR MEMO_NOTE_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_15_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_11_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " UPDATE MEMO_NOTE SET MEMO_NOTE_CREATE_BY = NULL "
   m_Conn.Execute (SQL)
   
   SQL = " UPDATE MEMO_NOTE SET MEMO_NOTE_CREATE_TO = NULL "
   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE MEMO_NOTE DROP CONSTRAINT MEMO_NOTE_CREATE_BY_FK "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE DROP CONSTRAINT MEMO_NOTE_CREATE_TO_FK "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE USER_ACCOUNT  ADD CONSTRAINT USER_ID_PK  PRIMARY KEY (USER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_CREATE_BY_FK  FOREIGN KEY (MEMO_NOTE_CREATE_BY ) REFERENCES USER_ACCOUNT  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE MEMO_NOTE  ADD CONSTRAINT MEMO_NOTE_CREATE_TO_FK  FOREIGN KEY (MEMO_NOTE_CREATE_TO ) REFERENCES USER_ACCOUNT  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_19_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2006_11_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " UPDATE RECEIPT_ITEM SET DEBIT_CREDIT_AMOUNT = RETURN_TOTAL_PRICE - RETURN_DISCOUNT_AMOUNT "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_21_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_01_22_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_01_22_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE BILLING_DOC ADD ENTRY_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE BILLING_DOC ADD EXIT_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_01_22_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_01_23_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_01_23_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
     
   m_Conn.BeginTrans
   HasBegin = True
   
   Call glbDaily.PatchDocumentPriority
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_01_23_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_01_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_01_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
  SQL = "CREATE TABLE  CSTITEM_RAW  "
   SQL = SQL & "( "
   SQL = SQL & "CSTITEM_RAW_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "COST_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "ITEM_COST                    MONEY_TYPE , "
   SQL = SQL & "ITEM_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CSTITEM_RAW ADD CONSTRAINT CSTITEM_RAW_ID_PK PRIMARY KEY (CSTITEM_RAW_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CSTITEM_RAW ADD CONSTRAINT CSTITEM_RAW_CSTITM_ID_FK FOREIGN KEY (COST_ITEM_ID) REFERENCES COST_ITEM  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CSTITEM_RAW ADD CONSTRAINT CSTITEM_RAW_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CSTITEM_RAW_SEQ  "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_01_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_02_01_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_01_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE PART_ITEM ADD MAXIMUM_ALLOW MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_01_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_02_07_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_07_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE LOT_ITEM ADD AUTO_PRICE FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
  SQL = "UPDATE LOT_ITEM SET AUTO_PRICE = 'N' "
   m_Conn.Execute (SQL)
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_07_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_06_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_06_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE INVENTORY_DOC ADD CREDIT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_06_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_06_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_06_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE LOT_ITEM ADD ITEM_DESC_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD CONSTRAINT LOT_ITEM_DESC_ID_FK FOREIGN KEY (ITEM_DESC_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_06_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_07_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_07_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE INVENTORY_DOC ADD PR_NO CODE_TYPE "   ' ใบขอซื้อ ที่ แผนกภายใน องค์กรออกมาเพื่อให้ แผนกจัดซื้อออก PO ให้ ลูกค้า
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_07_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_07_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_07_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "UPDATE PART_ITEM PI SET PI.PART_TYPE = 32 WHERE PI.PART_TYPE IN (SELECT DISTINCT PI1.PART_TYPE  FROM PART_ITEM  PI1 LEFT OUTER JOIN PART_TYPE PT ON  (PI1.PART_TYPE = PT.PART_TYPE_ID) WHERE  PT.PART_GROUP_ID = 18)  "   '
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_07_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_12_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_12_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
  SQL = "CREATE TABLE  EXPENSE_DETAIL  "
   SQL = SQL & "( "
   SQL = SQL & "EXPENSE_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "EXPENSE_DETAIL_TYPE                    ID_TYPE , "
   SQL = SQL & "EXPENSE_DETAIL_DESC                    DESC_TYPE , "
   SQL = SQL & "EXPENSE_DETAIL_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "EXPENSE_DETAIL_AVG                    MONEY_TYPE , "
   SQL = SQL & "EXPENSE_DETAIL_PRICE                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPENSE_DETAIL ADD CONSTRAINT EXPENSE_DETAIL_ID_PK PRIMARY KEY (EXPENSE_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR EXPENSE_DETAIL_SEQ  "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_12_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE TABLE  EXPENSE  "
   SQL = SQL & "( "
   SQL = SQL & "EXPENSE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "EXPENSE_NO                    CODE_TYPE , "
   SQL = SQL & "EXPENSE_DESC                    DESC_TYPE , "
   SQL = SQL & "EXPENSE_DATE                    DATE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPENSE ADD CONSTRAINT EXPENSE_ID_PK PRIMARY KEY (EXPENSE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR EXPENSE_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_13_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_13_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_13_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "ALTER TABLE EXPENSE_DETAIL ADD EXPENSE_ID ID_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPENSE_DETAIL ADD CONSTRAINT EXPENSE_DETAIL_EXPENSE_ID_FK FOREIGN KEY (EXPENSE_ID) REFERENCES EXPENSE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_13_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_13_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_02_13_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "ALTER TABLE EXPENSE_DETAIL ADD CONSTRAINT EXPENSE_DETAIL_EXPENSE_TYPE_FK FOREIGN KEY (EXPENSE_DETAIL_TYPE) REFERENCES PARAMETER_PROCESS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_13_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_02_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_02_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
  SQL = "CREATE TABLE  CUSTOMER_ACCOUNT_LIST "
   SQL = SQL & "( "
   SQL = SQL & "CUSTOMER_ACCOUNT_LIST_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID                    ID_TYPE , "
   SQL = SQL & "ACCOUNT_LIST_TYPE                    ID_TYPE , "
   SQL = SQL & "PART_GROUP_ID                    ID_TYPE , "
   SQL = SQL & "FEATURE_TYPE                    ID_TYPE , "
   SQL = SQL & "BANK_ACCOUNT_ID                    ID_TYPE , "
   SQL = SQL & "DEBIT_ID                    ID_TYPE , "
   SQL = SQL & "CREDIT_ID                    ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUSTOMER_ACCOUNT_LIST_ID_PK PRIMARY KEY (CUSTOMER_ACCOUNT_LIST_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUS_ACC_LIST_PART_GROUP_FK FOREIGN KEY (PART_GROUP_ID) REFERENCES PART_GROUP "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUS_ACC_LIST_FEATURE_FK FOREIGN KEY (FEATURE_TYPE) REFERENCES FEATURE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUS_ACC_LIST_BANK_ACCOUNT_FK FOREIGN KEY (BANK_ACCOUNT_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUS_ACC_LIST_DEBIT_FK FOREIGN KEY (DEBIT_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUS_ACC_LIST_CREDIT_FK FOREIGN KEY (CREDIT_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_ACCOUNT_LIST ADD CONSTRAINT CUS_ACC_LIST_CUSTOMER_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CUSTOMER_ACCOUNT_LIST_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_19_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
   
Public Function Patch_2007_02_23_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_02_23_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB_INOUT ADD PARAM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_02_23_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_02_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_03_02_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE MASTER_REF ADD MASTER_FLAG FLAG_TYPE "   '
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_01_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
   
Public Function Patch_2007_03_02_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_03_02_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "UPDATE MASTER_REF SET MASTER_FLAG = 'N' "   '
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_01_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
   
Public Function Patch_2007_03_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  SUP_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "SUP_ITEM_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "DO_ID                    ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "LOCATION_ID                    ID_TYPE , "
   SQL = SQL & "ITEM_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "TOTAL_WEIGHT                    MONEY_TYPE , "
   SQL = SQL & "AVG_WEIGHT                    MONEY_TYPE , "
  SQL = SQL & "TOTAL_PRICE                    MONEY_TYPE , "
  SQL = SQL & "AVG_PRICE                    MONEY_TYPE , "
  SQL = SQL & "GUI_ID                    ID_TYPE , "
  SQL = SQL & "PIG_STATUS                    ID_TYPE , "
  SQL = SQL & "LINK_ID                    ID_TYPE , "
  SQL = SQL & "CONFIG_CODE                    CODE_TYPE , "
   SQL = SQL & "SOC_ID                    ID_TYPE , "
   SQL = SQL & "FEATURE_ID                    ID_TYPE , "
   SQL = SQL & "ITEM_DESC                    DESC_TYPE , "
   SQL = SQL & "DEPOSIT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "DISCOUNT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "DISPLAY_ID                    ID_TYPE , "
   SQL = SQL & "PO_ID                    ID_TYPE , "
   SQL = SQL & "PO_NO                    CODE_TYPE , "
   SQL = SQL & "FROM_PERIOD                    DATE_TYPE , "
   SQL = SQL & "TO_PERIOD                    DATE_TYPE , "
   SQL = SQL & "DOLLAR                    MONEY_TYPE , "
    SQL = SQL & "COUNTRY_CURRENCY2                    ID_TYPE , "
   SQL = SQL & "COUNTRY_CURRENCY1                    ID_TYPE , "
   SQL = SQL & "QTYP                    MONEY_TYPE,"
    SQL = SQL & "PACK                    MONEY_TYPE , "
   SQL = SQL & "DISCOUNT_PERCENT                    MONEY_TYPE , "
   SQL = SQL & "WEIGHT_PER_PACK                    MONEY_TYPE , "
   SQL = SQL & "PACK_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "PRICE_PER_PACK                    MONEY_TYPE , "
   SQL = SQL & "DISCOUNT_PER_PACK                    MONEY_TYPE , "
   SQL = SQL & "MANUAL_FLAG                    FLAG_TYPE , "
   SQL = SQL & "MANUAL_CODE                    CODE_TYPE , "
   SQL = SQL & "MANUAL_NAME                    CODE_TYPE , "
   SQL = SQL & "RATE_TYPE                    ID_TYPE , "
   SQL = SQL & "EXTRA_DISCOUNT                    MONEY_TYPE , "
   SQL = SQL & "TRANSFER_WAGE                    MONEY_TYPE , "
   SQL = SQL & "STD_TRANSFER_CHARGE                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUP_ITEM ADD CONSTRAINT SUP_ITEM_ID_PK PRIMARY KEY (SUP_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUP_ITEM ADD CONSTRAINT SUP_ITEM_DO_ITEM_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SUP_ITEM_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_13_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_13_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_03_13_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE BILLING_DOC ADD DO_NO CODE_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD TRUCK_NO CODE_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD DELIVERY_FEE MONEY_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD SENDER_NAME NAME_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD RECEIVE_NAME NAME_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD DEPARTMENT_ID ID_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD QUE_NO CODE_TYPE "   '
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PR_NO CODE_TYPE "   '
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_13_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_13_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_03_13_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE CHEQUE ADD SUPPLIER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_SUP_ID_FK FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_13_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1 " & Err.DESCRIPTION
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_14_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_03_14_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE CHEQUE ADD TRANSFER_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
  SQL = "UPDATE CHEQUE SET TRANSFER_FLAG = 'Y' "
   m_Conn.Execute (SQL)
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_14_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1 " & Err.DESCRIPTION
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_15_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_15_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE CASH_DOC ADD SUPPLIER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_SUP_ID_FK FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CASH_TRAN ADD SUPPLIER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_SUP_ID_FK FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)
      
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_15_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_15_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_15_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "UPDATE CHEQUE SET TRANSFER_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_15_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_15_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_15_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CASH_DOC_POST ADD POST_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE CASH_DOC_POST SET POST_TYPE =  " & POST_CLEAR
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_15_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   Dim m_Sp As CSystemParam
   Set m_Sp = New CSystemParam
   m_Sp.AddEditMode = SHOW_ADD
   m_Sp.PARAM_NAME = "SHOW_IMPORT_INVENTORY"
   m_Sp.PARAM_VALUE = "N"
   m_Sp.PARAM_TYPE = "I"
   m_Sp.PARAM_DESC = "เป็นตัวแปรที่จะแสดงปุ่มนำเข้าจากคลังหรือเปล่าเพราะจะย้ายการนำเข้าไปที่บัญชีซื้อ"
   Call m_Sp.AddEditData
      
   Set m_Sp = Nothing
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_20_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_20_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_20_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE CHEQUE ADD CHEQUE_RCP_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
       
   SQL = "ALTER TABLE CHEQUE ADD CHEQUE_UMN_DATE DATE_TYPE "              'UP MONEY DATE วันที่ขึ้นเงิน
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_20_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_20_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_20_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE JOB_INOUT SET PARAM_ID = NULL WHERE PARAM_ID NOT IN ( SELECT PARAMETER_PROCESS_ID FROM PARAMETER_PROCESS) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_20_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_20_2_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_20_2_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                
   m_Conn.BeginTrans
   HasBegin = True
   
   Call glbDaily.PatchDocumentPriority2
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_20_2_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_21_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_21_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                
   m_Conn.BeginTrans
   HasBegin = True
   
   Call glbDaily.PatchDocumentPriority3
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_21_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_22_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_22_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                
   m_Conn.BeginTrans
   HasBegin = True
   
   Call glbDaily.PatchDocumentPriority4
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_22_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM ADD PO_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUP_ITEM ADD PO_NO CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_26_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "UPDATE SUP_ITEM SET PO_ID =  NULL "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUP_ITEM ADD CONSTRAINT SUP_ITEM_PO_ID_FK FOREIGN KEY (PO_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_26_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  GL_DETAIL "
   SQL = SQL & "( "
   SQL = SQL & "GL_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "BILLING_DOC_ID                    ID_TYPE , "
   SQL = SQL & "GL_DESC                    ID_TYPE , "
   SQL = SQL & "GL_TYPE                    ID_TYPE , "
   SQL = SQL & "GL_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR GL_DETAIL_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_28_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_28_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_28_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_DETAIL ADD CONSTRAINT GL_DETAIL_ID_PK PRIMARY KEY (GL_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_28_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_28_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_28_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_DETAIL ADD CONSTRAINT GL_DETAIL_BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_28_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_03_28_4_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_28_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_DETAIL ADD GL_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE GL_DETAIL ADD CONSTRAINT GL_DETAIL_GL_ID_FK FOREIGN KEY (GL_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_28_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_28_5_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_28_5_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_DETAIL DROP GL_DESC  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE GL_DETAIL ADD GL_DESC DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_28_5_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_03_28_6_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_03_28_6_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PV_NO CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_28_6_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_02_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_02_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  EVALUATE_PAY "
   SQL = SQL & "( "
   SQL = SQL & "EVALUATE_PAY_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "SUPPLIER_ID                    ID_TYPE , "
   SQL = SQL & "EVALUATE_DATE                    DATE_TYPE , "
   SQL = SQL & "EVALUATE_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR EVALUATE_PAY_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_02_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_02_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_02_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EVALUATE_PAY ADD CONSTRAINT EVALUATE_PAY_ID_PK PRIMARY KEY (EVALUATE_PAY_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_02_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_02_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_02_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EVALUATE_PAY ADD CONSTRAINT EVALUATE_PAY_SUPPLIER_ID_FK FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_02_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_02_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_02_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EVALUATE_PAY ADD EVALUATE_PAY_DESC DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_02_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_09_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_09_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM ADD TEMP_LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE SUP_ITEM SET TEMP_LOCATION_ID = LOCATION_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_09_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_09_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_09_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM DROP  LOCATION_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUP_ITEM ADD  LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_09_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_09_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_09_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM ADD CONSTRAINT SUP_ITEM_LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE SUP_ITEM SET LOCATION_ID = TEMP_LOCATION_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_09_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_09_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_09_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM ADD TEMP_PART_ITEM_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE SUP_ITEM SET TEMP_PART_ITEM_ID = PART_ITEM_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_09_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_09_5_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_09_5_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM DROP  PART_ITEM_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUP_ITEM ADD  PART_ITEM_ID ID_TYPE "
   m_Conn.Execute (SQL)
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_09_5_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_09_6_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_09_6_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE SUP_ITEM ADD CONSTRAINT SUP_ITEM_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE SUP_ITEM SET PART_ITEM_ID = TEMP_PART_ITEM_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_09_6_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_04_27_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_27_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE MASTER_REF ADD SUM_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  MASTER_REF SET  SUM_FLAG  = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_27_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_04_27_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_27_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD EFFECTIVE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE BILLING_DOC SET EFFECTIVE_FLAG = 'N' WHERE BILLING_DOC_ID NOT IN (SELECT BILLING_DOC_ID FROM CASH_TRAN) "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILLING_DOC SET EFFECTIVE_FLAG = 'Y' WHERE BILLING_DOC_ID IN (SELECT BILLING_DOC_ID FROM CASH_TRAN) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_27_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_04_27_2_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_04_27_2_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE REPORT_CONFIG ADD COMPUTER_NAME NAME_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_04_27_2_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_05_22_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_05_22_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   '--------------------------------------1
   
   SQL = "UPDATE  MASTER_REF SET  MASTER_FLAG  =  'Y' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1100'    "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - ภายนอก' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1100'    "
   m_Conn.Execute (SQL)
   
   '--------------------------------------2
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บมจ.มิตรภาพโภคภัณฑ์ MP' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1203'    "
   m_Conn.Execute (SQL)
   '--------------------------------------3
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บมจ.มิตรภาพโภคภัณฑ์ MA1' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1204'    "
   m_Conn.Execute (SQL)
   '--------------------------------------4
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บมจ.มิตรภาพโภคภัณฑ์ MA2' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1205'    "
   m_Conn.Execute (SQL)
   '--------------------------------------5
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บจก.มิตรภาพไฮบริด MH' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1206'    "
   m_Conn.Execute (SQL)
   
   '--------------------------------------6
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บจก.พันธุ์สุกรไทยพัฒนา DTS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1207'    "
   m_Conn.Execute (SQL)
   '--------------------------------------7
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บจก.เอ็มจีฟาร์ม่า MGP' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1212'    "
   m_Conn.Execute (SQL)
   '--------------------------------------8
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บจก.ไฮควอลิตี้ซีเมน HQS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1221'    "
   m_Conn.Execute (SQL)
   '--------------------------------------9
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บจก.ควอลิตี้มีท QMC' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1222'    "
   m_Conn.Execute (SQL)
      '--------------------------------------10
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'ลูกหนี้การค้า - บจก.แก่นตะวันไบโอเทค KTV' WHERE MASTER_AREA =  9 AND KEY_CODE =  '113-1226'    "
   m_Conn.Execute (SQL)
   
      '--------------------------------------11
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-MP' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-1002'    "
   m_Conn.Execute (SQL)
      '--------------------------------------12
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-MA1' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-1003'    "
   m_Conn.Execute (SQL)
      '--------------------------------------13
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-MA2' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-1004'    "
   m_Conn.Execute (SQL)
      '--------------------------------------14
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-MH' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-1201'    "
   m_Conn.Execute (SQL)
      '--------------------------------------15
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-DTS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-1202'    "
   m_Conn.Execute (SQL)
   
      '--------------------------------------16
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-HQS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-1203'    "
   m_Conn.Execute (SQL)
      '--------------------------------------17
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-QMC' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-2001'    "
   m_Conn.Execute (SQL)
      '--------------------------------------18
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายอาหาร-ภายนอก' WHERE MASTER_AREA =  9 AND KEY_CODE =  '411-2006'    "
   m_Conn.Execute (SQL)
      '--------------------------------------19
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-MP' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-1002'    "
   m_Conn.Execute (SQL)
      '--------------------------------------20
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-MA1' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-1003'    "
   m_Conn.Execute (SQL)
   
      '--------------------------------------21
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-MA2' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-1004'    "
   m_Conn.Execute (SQL)
      '--------------------------------------22
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-DTS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-1005'    "
   m_Conn.Execute (SQL)
      '--------------------------------------23
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-MH' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-1201'    "
   m_Conn.Execute (SQL)
      '--------------------------------------24
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-HQS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-1203'    "
   m_Conn.Execute (SQL)
      '--------------------------------------25
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-QMC' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-2001'    "
   m_Conn.Execute (SQL)
   
      '--------------------------------------26
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้จากการขายวัตถุดิบ-ภายนอก' WHERE MASTER_AREA =  9 AND KEY_CODE =  '412-2006'    "
   m_Conn.Execute (SQL)
      '--------------------------------------27
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-MP' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-1002'    "
   m_Conn.Execute (SQL)
      '--------------------------------------28
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-MA1' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-1003'    "
   m_Conn.Execute (SQL)
      '--------------------------------------29
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-MA2' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-1004'    "
   m_Conn.Execute (SQL)
      '--------------------------------------30
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-MH' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-1201'    "
   m_Conn.Execute (SQL)

      '--------------------------------------31
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-DTS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-1202'    "
   m_Conn.Execute (SQL)
      '--------------------------------------32
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-HQS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-1203'    "
   m_Conn.Execute (SQL)
      '--------------------------------------33
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-QMC' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-2001'    "
   m_Conn.Execute (SQL)
      '--------------------------------------34
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้ค่าขนส่ง-ภายนอก' WHERE MASTER_AREA =  9 AND KEY_CODE =  '413-2006'    "
   m_Conn.Execute (SQL)
      '--------------------------------------35
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-MP' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9004'    "
   m_Conn.Execute (SQL)
         
         '--------------------------------------36
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-MA1' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9005'    "
   m_Conn.Execute (SQL)
      '--------------------------------------37
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-MA2' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9006'    "
   m_Conn.Execute (SQL)
      '--------------------------------------38
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-MH' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9007'    "
   m_Conn.Execute (SQL)
      '--------------------------------------39
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-DTS' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9009'    "
   m_Conn.Execute (SQL)
      '--------------------------------------40
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-QMC' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9101'    "
   m_Conn.Execute (SQL)
   
         '--------------------------------------41
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-KTV' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9104'    "
   m_Conn.Execute (SQL)
      '--------------------------------------42
   SQL = "UPDATE MASTER_REF SET KEY_NAME = 'รายได้อื่นๆ-ภายนอก' WHERE MASTER_AREA =  9 AND KEY_CODE =  '420-9999'    "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_05_22_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_22_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_05_22_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 1, PRIORITY2 = 1 WHERE DOCUMENT_TYPE = 23 "
   m_Conn.Execute (SQL)
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_05_22_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_28_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_05_28_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD PAY_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
            
   SQL = "UPDATE BILLING_DOC SET PAY_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_05_28_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_29_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_05_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   Dim m_Sp As CMasterRef
   Set m_Sp = New CMasterRef
   m_Sp.AddEditMode = SHOW_ADD
   m_Sp.KEY_CODE = "412-2008"
   m_Sp.KEY_NAME = "รายได้จากการขายวัตถุดิบ-KTV"
   m_Sp.MASTER_AREA = ACCOUNT_LIST
   Call m_Sp.AddEditData
      
   Set m_Sp = New CMasterRef
   m_Sp.AddEditMode = SHOW_ADD
   m_Sp.KEY_CODE = "413-2008"
   m_Sp.KEY_NAME = "รายได้ค่าขนส่ง-KTV"
   m_Sp.MASTER_AREA = ACCOUNT_LIST
   Call m_Sp.AddEditData
   
   Set m_Sp = Nothing
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_29_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_29_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim ItemCount  As Long
   
   VersionStr = "2007_05_29_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   Dim TempColl    As Collection
   Set TempColl = New Collection
   Call LoadMaster(Nothing, TempColl, ACCOUNT_LIST, , , 2)
   
   Dim m_Sp As CCustomerAccountList
   Dim m_Rs As ADODB.Recordset
   Dim Mr  As CMasterRef
   Set m_Sp = New CCustomerAccountList
   Set m_Rs = New ADODB.Recordset
   
   Call m_Sp.SetFieldValue("CUSTOMER_ACCOUNT_LIST_ID", -1)
   Call m_Sp.SetFieldValue("CUSTOMER_CODE", "M-5002")
   Call m_Sp.SetFieldValue("FEATURE_TYPE_NO", "06")
   Call m_Sp.QueryData(2, m_Rs, ItemCount)
      
   Call m_Sp.PopulateFromRS(2, m_Rs)
   
   Set Mr = GetObject("CMasterRef", TempColl, "413-2008")
   
   Call m_Sp.SetFieldValue("CREDIT_ID", Mr.KEY_ID)
   m_Sp.ShowMode = SHOW_EDIT
   Call m_Sp.AddEditData
   
   Set m_Rs = Nothing
   Set m_Sp = Nothing
   
   Set m_Sp = New CCustomerAccountList
   Set m_Rs = New ADODB.Recordset
   Call m_Sp.SetFieldValue("CUSTOMER_ACCOUNT_LIST_ID", -1)
   Call m_Sp.SetFieldValue("CUSTOMER_CODE", "M-5002")
   Call m_Sp.SetFieldValue("PART_GROUP_NO", "00")
   Call m_Sp.QueryData(2, m_Rs, ItemCount)
      
   Call m_Sp.PopulateFromRS(2, m_Rs)
   
   Set Mr = GetObject("CMasterRef", TempColl, "412-2008")
   
   Call m_Sp.SetFieldValue("CREDIT_ID", Mr.KEY_ID)
   m_Sp.ShowMode = SHOW_EDIT
   Call m_Sp.AddEditData
   
   
   Set m_Sp = Nothing
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_03_29_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_06_04_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_04_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE CASH_TRAN ADD UNDER_PAY MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CASH_TRAN ADD OVER_PAY MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CASH_TRAN ADD ADVANCE_PAY MONEY_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
            
   SQL = "UPDATE BILLING_DOC SET PAY_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_04_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_06_06_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_06_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE PART_ITEM ADD OLD_PART_ID ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
            
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_06_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_06_22_1_jill() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_22_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CASH_DOC_POST ADD BILLING_DOC_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_22_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_06_22_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_22_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BILL_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_22_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_06_28_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC BD SET BD.PAID_AMOUNT = (SELECT SUM(RI.PAID_AMOUNT) FROM RECEIPT_ITEM RI WHERE RI.DO_ID =  BD.BILLING_DOC_ID)  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_28_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_06_21_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_21_1_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE CASH_TRAN ADD WH_PAY MONEY_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
            
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_21_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_06_29_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_GROUP ADD SET_PRODUCT_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_29_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_06_29_2_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_06_29_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE FEATURE_TYPE ADD FEATURE_GROUP_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FEATURE_TYPE ADD CONSTRAINT FEATURE_TYPE_GROUP_ID_FK FOREIGN KEY (FEATURE_GROUP_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_GROUP ADD CONSTRAINT PART_GROUP_SET_ID_FK FOREIGN KEY (SET_PRODUCT_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_06_29_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_11_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_11_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE SUPPLIER ADD SUPPLIER_CHEQUE_NAME NAME_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_11_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_18_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CASH_DOC_POST ADD WH_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_18_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_18_2_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_18_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CASH_TRAN ADD INTERREST_PAY MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_18_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_18_3_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_18_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CASH_DOC_POST ADD INTERREST_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_18_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_25_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD DUE_AMOUNT ID_TYPE "     'จำนวนวันที่ต้องการให้ส่งของ เช่น duedate เป็นวันที่ 15 จำนวน due_amount เป็น 7 วัน ก็ 15-21/07/50
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_26_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Dim CP  As CCashDocPost
   Dim m_Rs  As ADODB.Recordset
   Dim ItemCount As Long
   Set CP = New CCashDocPost
   Set m_Rs = New ADODB.Recordset
   
   Call CP.QueryData(1, m_Rs, ItemCount)
   While Not m_Rs.EOF
      Call CP.PopulateFromRS(1, m_Rs)
      
      SQL = "UPDATE BILLING_DOC SET DOCUMENT_NO = '" & CP.GetFieldValue("BILLING_DOC_NO") & "' WHERE BILLING_DOC_ID = " & CP.GetFieldValue("BILLING_DOC_ID")
      m_Conn.Execute (SQL)
      
      m_Rs.MoveNext
   Wend
   
   Set m_Rs = Nothing
   Set CP = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_26_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_08_14_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_08_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
   m_Conn.BeginTrans
   HasBegin = True
   
   Dim CP  As CCashDocPost
   Dim m_Rs  As ADODB.Recordset
   Dim ItemCount As Long
   Set CP = New CCashDocPost
   Set m_Rs = New ADODB.Recordset
   
   Call CP.QueryData(1, m_Rs, ItemCount)
   While Not m_Rs.EOF
      Call CP.PopulateFromRS(1, m_Rs)
      
      SQL = "UPDATE BILLING_DOC SET DOCUMENT_NO = '" & CP.GetFieldValue("BILLING_DOC_NO") & "' WHERE BILLING_DOC_ID = " & CP.GetFieldValue("BILLING_DOC_ID")
      m_Conn.Execute (SQL)
      
      m_Rs.MoveNext
   Wend
   
   Set m_Rs = Nothing
   Set CP = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_08_14_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_24_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_24_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE JOB ADD YYYYMM CODE_TYPE "     '
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE JOB SET YYYYMM = SUBSTR(JOB_DATE, 1, 7) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_24_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_02_13_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2008_02_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD YYYYMM CODE_TYPE "     '
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE INVENTORY_DOC SET YYYYMM = SUBSTR(DOCUMENT_DATE, 1, 7) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_02_13_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_05_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_05_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  MONTHLY_ACCUM "
   SQL = SQL & "( "
   SQL = SQL & "MONTHLY_ACCUM_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "YYYYMM                    CODE_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "LOCATION_ID                    ID_TYPE , "
   SQL = SQL & "BALANCE_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "TOTAL_INCLUDE_PRICE                    MONEY_TYPE , "
   SQL = SQL & "AVG_PRICE                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR MONTHLY_ACCUM_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_05_04_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_05_04_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_05_04_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD CONSTRAINT MONTHLY_ACCUM_ID_PK PRIMARY KEY (MONTHLY_ACCUM_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_05_04_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_05_04_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_05_04_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD CONSTRAINT MONTHLY_ACCUM_PI_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD CONSTRAINT MONTHLY_ACCUM_LI_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_05_04_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_05_02_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2013_05_02_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  PLAN_PART "
   SQL = SQL & "( "
   SQL = SQL & "PLAN_PART_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "PLAN_AREA                   ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "PLAN_DATE                    DATE_TYPE , "
   SQL = SQL & "PLAN_IN                    MONEY_TYPE , "
   SQL = SQL & "PLAN_OUT                    MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PLAN_PART_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_05_02_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_05_02_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2013_05_02_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PLAN_PART ADD CONSTRAINT PLAN_PART_ID_PK PRIMARY KEY (PLAN_PART_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_05_02_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_05_02_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_05_02_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PLAN_PART ADD CONSTRAINT PLAN_PART_PI_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_05_02_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_05_02_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_05_02_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PLAN_PART ADD NOTE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_05_02_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_05_15_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_05_15_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_ITEM ADD CANCEL_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PART_ITEM SET CANCEL_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_05_15_1_yong = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_06_03_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_06_03_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PLAN_PART ADD CANCEL_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PLAN_PART SET CANCEL_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_06_03_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_06_17_1_yong() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_06_17_1_yong"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE PART_ITEM ADD NUMBER_PLC_ID CODE_TYPE UNIQUE"
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_06_17_1_yong = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_06_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_06_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_ITEM ADD DEFAULT_LOCATION ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_06_26_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_06_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2013_06_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_ITEM ADD CONSTRAINT PART_ITEM_LOCATION_FK FOREIGN KEY (DEFAULT_LOCATION) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_06_26_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_01_13_1_pui() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "2015_01_13_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD GENERATION  DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD  REFERENCE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD FARM_NAME DESC_TYPE "
   m_Conn.Execute (SQL)
                              
'   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_TYPE ID_TYPE "
'   m_Conn.Execute (SQL)

   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_01_13_1_pui = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_2015_02_02_1_pui() As Boolean '
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "2015_02_02_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE SUP_ITEM  ADD PROJECT_NAME_ID  ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM ADD  PROJECT_NAME_ID  ID_TYPE "
   m_Conn.Execute (SQL)
   

   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_02_02_1_pui = True
   Exit Function
   
ErrorHandler:
   ''Debug.Print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_2015_02_02_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_02_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE SUP_ITEM ADD CONSTRAINT SUP_ITEM_PROJECT_NAME_ID_FK FOREIGN KEY (PROJECT_NAME_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)


   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_02_2_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_02_02_3_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_02_3_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)


    SQL = "ALTER TABLE LOT_ITEM ADD CONSTRAINT LOT_ITEM_PROJECT_NAME_ID_FK FOREIGN KEY (PROJECT_NAME_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_02_3_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function





Public Function Patch_2015_02_20_1_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_02_20_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  ACC_FOL "
   SQL = SQL & "( "
   SQL = SQL & "ACC_FOL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID                 ID_TYPE , "
   SQL = SQL & "FOL_NOTE                    DESC_TYPE , "
   SQL = SQL & "FOL_DATE                    DATE_TYPE , "
   SQL = SQL & "CANCEL_FLAG                   FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL , "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR ACC_FOL_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_02_20_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_02_20_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_20_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE ACC_FOL  ADD CONSTRAINT ACC_FOL_ID_PK PRIMARY KEY (ACC_FOL_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_20_2_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_02_20_3_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_20_3_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE ACC_FOL ADD CONSTRAINT ACC_FOL_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER"
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_20_3_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_02_20_4_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_20_4_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  MKT_FOL "
   SQL = SQL & "( "
   SQL = SQL & "MKT_FOL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID                 ID_TYPE , "
   SQL = SQL & "FOL_NOTE                    DESC_TYPE , "
   SQL = SQL & "FOL_DATE                    DATE_TYPE , "
   SQL = SQL & "CANCEL_FLAG                   FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR  MKT_FOL_SEQ  "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_20_4_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_02_20_5_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_20_5_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE  MKT_FOL  ADD CONSTRAINT MKT_FOL_ID_PK PRIMARY KEY (MKT_FOL_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_20_5_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'Public Function Patch_2015_02_20_6_pui() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2015_02_20_6_pui"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE MKT_FOL ADD CONSTRAINT CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER"
'
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2015_02_20_6_pui = True
'   Exit Function
'
'ErrorHandler:
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function

Public Function Patch_2015_02_20_6_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_02_20_6_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE MKT_FOL ADD CONSTRAINT MKT_FOL_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER"

   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_02_20_6_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function


Public Function Patch_2015_03_18_1_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_03_18_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE CUSTOMER ADD MAX_CREDIT LEN_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_18_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_03_21_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_03_21_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD SUSPEND_SALES FLAG_TYPE DEFAULT 'N'  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_21_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_03_21_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2015_03_21_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE CUSTOMER SET SUSPEND_SALES = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_21_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_03_27_1_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_03_27_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  CHEQUE_DOC"
   SQL = SQL & "( "
   SQL = SQL & "CHEQUE_DOC_ID                       ID_TYPE NOT NULL, "
  SQL = SQL & "CHEQUE_DOC_NO                      CODE_TYPE , "
  SQL = SQL & "CHEQUE_DOC_DATE                DATE_TYPE , "
   SQL = SQL & "BADCHEQUE_FLAG                FLAG_TYPE , "
   SQL = SQL & "PASSCHEQUE_FLAG              FLAG_TYPE , "
   SQL = SQL & "BADCHEQUE_DATE                DATE_TYPE , "
   SQL = SQL & "PASSCHEQUE_DATE               DATE_TYPE , "
   SQL = SQL & "CUSTOMER_ID                          ID_TYPE , "
   SQL = SQL & "CREATE_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                              ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                       DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                              ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CHEQUE_DOC_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_27_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_03_27_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_03_27_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CHEQUE_DOC ADD CONSTRAINT CHEQUE_DOC_ID_PK PRIMARY KEY (CHEQUE_DOC_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_27_2_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_03_27_3_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2015_03_27_3_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CHEQUE_DOC ADD CONSTRAINT CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_27_3_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_03_27_4_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_03_27_4_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  RECEIPT_CHEQUE_DOC"
   SQL = SQL & "( "
   SQL = SQL & "RECEIPT_CHEQUE_DOC_ID                       ID_TYPE NOT NULL, "
  SQL = SQL & "RECEIPT_CHEQUE_DOC_NO                      CODE_TYPE , "
  SQL = SQL & "RECEIPT_CHEQUE_DOC_DATE                DATE_TYPE , "
   SQL = SQL & "CHEQUE_DOC_ID                                            ID_TYPE , "
   SQL = SQL & "CREATE_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                              ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                       DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                              ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR RECEIPT_CHEQUE_DOC_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_27_4_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_03_27_5_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_03_27_5_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD CONSTRAINT RECEIPT_CHEQUE_DOC_ID_PK PRIMARY KEY (RECEIPT_CHEQUE_DOC_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_27_5_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_03_27_6_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2015_03_27_6_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD CONSTRAINT CHEQUE_DOC_ID_FK FOREIGN KEY (CHEQUE_DOC_ID) REFERENCES  CHEQUE_DOC"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_03_27_6_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_05_21_1_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_05_21_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD BILLING_DOC_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_05_21_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_05_21_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_05_21_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD CONSTRAINT  BILLING_DOC_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_05_21_2_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_05_26_1_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_05_26_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD RECEIPT_ITEM_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_05_26_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_05_26_2_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_05_26_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD CASH_DISCOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_05_26_2_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_05_26_3_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_05_26_3_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_CHEQUE_DOC ADD PAID_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_05_26_3_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_06_10_1_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_06_10_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE  CHEQUE_DOC ADD AMOUNT_CHEQUE  MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_06_10_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_08_04_1_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_08_04_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD CHEQUE_DOC_ID  ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_08_04_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_08_04_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_08_04_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT  CHEQUE_DOC_FK FOREIGN KEY (CHEQUE_DOC_ID) REFERENCES CHEQUE_DOC "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_08_04_2_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_09_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  PLANNING "
   SQL = SQL & "( "
   SQL = SQL & "PLANNING_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "PLANNING_AREA                   ID_TYPE , "
   SQL = SQL & "PLANNING_FROM                    DATE_TYPE , "
   SQL = SQL & "PLANNING_TO                    DATE_TYPE , "
   SQL = SQL & "PLANNING_DATE                    DATE_TYPE , "
   SQL = SQL & "PLANNING_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PLANNING ADD CONSTRAINT PLANNING_ID_PK PRIMARY KEY (PLANNING_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PLANNING_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_10_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_10_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_09_10_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER PLANNING_BA FOR PLANNING "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.PLANNING_ID  = GEN_ID (PLANNING_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_10_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_11_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_09_11_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  PLANNING_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "PLANNING_ITEM_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "PLANNING_ID                   ID_TYPE , "
   SQL = SQL & "PLANNING_SUB_TYPE                   ID_TYPE , "                  '1.ประมาณวัตถดิบที่ใช้ 2.ประมาณผลิตภัณฑ์ที่ได้ 3.ประมาณวัตถุดิบที่จะรับเข้า
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , "
   SQL = SQL & "PLAN_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "NOTE                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PLANNING_ITEM ADD CONSTRAINT PLANNING_ITEM_ID_PK PRIMARY KEY (PLANNING_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PLANNING_ITEM_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_11_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_11_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_09_11_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER PLANNING_ITEM_BA FOR PLANNING_ITEM "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.PLANNING_ITEM_ID  = GEN_ID (PLANNING_ITEM_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_11_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_11_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_09_11_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE PLANNING_ITEM ADD CONSTRAINT  PLANNING_ITEM_FK FOREIGN KEY (PLANNING_ID) REFERENCES PLANNING "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_09_11_3_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_11_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_09_11_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE PLANNING_ITEM ADD CONSTRAINT  PLANNING_ITEM_PART_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_11_4_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_17_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_09_17_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE PART_ITEM ADD NUMBER_LAB_ID CODE_TYPE UNIQUE"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_17_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function


Public Function Patch_2015_09_15_1_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_09_15_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONDITION   ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PAID_TYPE  ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_15_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_15_2_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_09_15_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT CONDITION_FK FOREIGN KEY (CONDITION) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)

 SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT PAID_TYPE_FK FOREIGN KEY (PAID_TYPE) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_09_15_2_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_25_1_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_09_25_1_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE BILLING_DOC ADD  CLOSE_FLAG FLAG_TYPE  " 'ติดตามใบค้างส่ง หรือยกเลิก การติดตามใบ PO
   'DEFAULT 'N'
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_25_1_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2015_09_25_2_pui() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_09_25_2_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CHEQUE_DOC ADD BANK_ID   ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CHEQUE_DOC ADD BANK_BRANCH_ID  ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_09_25_2_pui = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_09_25_3_pui() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_09_25_3_pui"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE CHEQUE_DOC ADD CONSTRAINT BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)

 SQL = "ALTER TABLE CHEQUE_DOC ADD CONSTRAINT BANK_BRANCH_ID_FK FOREIGN KEY (BANK_BRANCH_ID) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_09_25_3_pui = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_10_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD AUTO_GEN_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_13_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_13_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_10_13_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD GEN_COMMIT_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_13_2_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_13_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_10_13_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "UPDATE BILLING_DOC SET AUTO_GEN_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_13_3_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_13_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_10_13_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET GEN_COMMIT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_13_4_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_10_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  ALERT_BOX "
   SQL = SQL & "( "
   SQL = SQL & "ALERT_BOX_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "ALERT_BOX_FROM                    DATE_TYPE , "
   SQL = SQL & "ALERT_BOX_TO                    DATE_TYPE , "
   SQL = SQL & "ALERT_BOX_TYPE                   ID_TYPE , "
   SQL = SQL & "ALERT_BOX_DESC                    DESC_TYPE , "
   SQL = SQL & "ALERT_ALL_FLAG                    FLAG_TYPE , "
   SQL = SQL & "ALERT_CANCEL_FLAG                    FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ALERT_BOX ADD CONSTRAINT ALERT_BOX_ID_PK PRIMARY KEY (ALERT_BOX_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR ALERT_BOX_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_19_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_19_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_10_19_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER ALERT_BOX_BA FOR ALERT_BOX "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.ALERT_BOX_ID  = GEN_ID (ALERT_BOX_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_19_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_20_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_10_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  ALERT_DETAIL "
   SQL = SQL & "( "
   SQL = SQL & "ALERT_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "ALERT_BOX_ID                   ID_TYPE , "
   SQL = SQL & "USER_ID                  ID_TYPE , "
   SQL = SQL & "READ_FLAG                    FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ALERT_DETAIL ADD CONSTRAINT ALERT_DETAIL_ID_PK PRIMARY KEY (ALERT_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR ALERT_DETAIL_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_20_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_20_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_10_20_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER ALERT_DETAIL_BA FOR ALERT_DETAIL "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.ALERT_DETAIL_ID  = GEN_ID (ALERT_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_20_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_20_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_10_20_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE ALERT_DETAIL ADD CONSTRAINT  ALERT_DETAIL_FK FOREIGN KEY (ALERT_BOX_ID) REFERENCES ALERT_BOX "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2015_10_20_3_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_10_20_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_10_20_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE ALERT_DETAIL ADD CONSTRAINT  ALERT_DETAIL_USER_FK FOREIGN KEY (USER_ID) REFERENCES USER_ACCOUNT "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_10_20_4_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_11_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_11_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE USER_ACCOUNT ADD LOGON_STATUS ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_11_10_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_12_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_12_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD PO_APPROVED_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_12_10_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_12_10_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2015_12_10_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "UPDATE BILLING_DOC SET PO_APPROVED_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_12_10_2_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_12_10_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_12_10_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD CREATE_NAME CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD APPROVE_NAME CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_12_10_3_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_12_13_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_12_13_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD MODIFY_NAME CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_12_13_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_12_14_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_12_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC DROP CREATE_BY "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_12_14_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2015_12_14_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2015_12_14_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC DROP MODIFY_BY "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2015_12_14_2_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_03_07_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_03_07_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE USER_ACCOUNT ADD IP_ADDRESS USERNAME_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_03_07_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_03_09_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_03_09_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE SUP_ITEM ADD PART_SUPPLIER_NAME NAME_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_03_09_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_03_15_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_03_15_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE USER_ACCOUNT ADD REAL_NAME DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_03_15_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_03_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_03_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE SUP_ITEM ADD PART_NOTE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_03_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_05_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_05_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  AUTHEN_PO "
   SQL = SQL & "( "
   SQL = SQL & "AUTHEN_PO_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "AUTHEN_PO_GROUP           ID_TYPE , "
   SQL = SQL & "AUTHEN_PO_FROM           MONEY_TYPE , "
   SQL = SQL & "AUTHEN_PO_TO           MONEY_TYPE , "
   SQL = SQL & "AUTHEN_PO_DESC           DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE AUTHEN_PO ADD CONSTRAINT AUTHEN_PO_ID_PK PRIMARY KEY (AUTHEN_PO_ID) "
   m_Conn.Execute (SQL)
      
   SQL = "CREATE GENERATOR AUTHEN_PO_SEQ "
   m_Conn.Execute (SQL)
      
    m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_05_12_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2016_05_12_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_05_12_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  AUTHEN_DETAIL "
   SQL = SQL & "( "
   SQL = SQL & "AUTHEN_DETAIL_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "AUTHEN_PO_ID           ID_TYPE , "
   SQL = SQL & "AUTHEN_AREA          ID_TYPE , "
   SQL = SQL & "AUTHEN_USER         CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE AUTHEN_DETAIL ADD CONSTRAINT AUTHEN_DETAIL_ID_PK PRIMARY KEY (AUTHEN_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE AUTHEN_DETAIL ADD CONSTRAINT AUTHEN_PO_ID_FK FOREIGN KEY (AUTHEN_PO_ID) REFERENCES AUTHEN_PO  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR AUTHEN_DETAIL_SEQ "
   m_Conn.Execute (SQL)
      
    m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_05_12_2_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2016_05_12_3_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_05_12_3_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE BILLING_DOC ADD VERIFY_BY_NAME CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_05_12_3_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_06_01_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2016_06_01_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER AUTHEN_PO_BA FOR AUTHEN_PO "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.AUTHEN_PO_ID  = GEN_ID (AUTHEN_PO_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_06_01_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_06_01_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2016_06_01_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER AUTHEN_DETAIL_BA FOR AUTHEN_DETAIL "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.AUTHEN_DETAIL_ID  = GEN_ID (AUTHEN_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_06_01_2_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2016_09_22_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2016_09_22_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE BILLING_DOC ADD TAX_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_09_22_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_03_16_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_03_16_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  INVENTORY_ACT "
   SQL = SQL & "( "
   SQL = SQL & "INVENTORY_ACT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "INVENTORY_ACT_AREA                   ID_TYPE , "
   SQL = SQL & "INVENTORY_ACT_DATE                    DATE_TYPE , "
   SQL = SQL & "INVENTORY_ACT_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_ACT ADD CONSTRAINT INVENTORY_ACT_ID PRIMARY KEY (INVENTORY_ACT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR INVENTORY_ACT_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_03_16_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_03_16_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_03_16_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  INVENTORY_ACT_ITEM  "
   SQL = SQL & "( "
   SQL = SQL & "INVENTORY_ACT_ITEM_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "INVENTORY_ACT_ID                   ID_TYPE , "
   SQL = SQL & "INVENTORY_ACT_ITEM_TYPE                   ID_TYPE , "                  '1.ประมาณวัตถดิบที่ใช้ 2.ประมาณผลิตภัณฑ์ที่ได้ 3.ประมาณวัตถุดิบที่จะรับเข้า
   SQL = SQL & "PART_ITEM_ID                   ID_TYPE , "
   SQL = SQL & "INVENTORY_ACT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "NOTE                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_ACT_ITEM ADD CONSTRAINT INVENTORY_ACT_ITEM_ID_PK PRIMARY KEY (INVENTORY_ACT_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR INVENTORY_ACT_ITEM_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_03_16_2_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_15_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_15_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  TARGET "
   SQL = SQL & "( "
   SQL = SQL & "TARGET_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "YEAR_NO                   ID_TYPE , "
   SQL = SQL & "TARGET_DESC                  DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME             CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TARGET ADD CONSTRAINT TARGET_ID_PK PRIMARY KEY (TARGET_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR TARGET_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_15_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_15_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_15_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER TARGET_BA FOR TARGET "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.TARGET_ID  = GEN_ID (TARGET_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_15_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  TARGET_DETAIL "
   SQL = SQL & "( "
   SQL = SQL & "TARGET_DETAIL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "TARGET_ID                   ID_TYPE , "
   SQL = SQL & "EMP_ID                   ID_TYPE , "
   SQL = SQL & "TARGET_PRICE1                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE2                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE3                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE4                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE5                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE6                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE7                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE8                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE9                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE10                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE11                   MONEY_TYPE , "
   SQL = SQL & "TARGET_PRICE12                  MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME             CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TARGET_DETAIL ADD CONSTRAINT TARGET_DETAIL_ID_PK PRIMARY KEY (TARGET_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR TARGET_DETAIL_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER TARGET_DETAIL_BA FOR TARGET_DETAIL "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.TARGET_DETAIL_ID  = GEN_ID (TARGET_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE TARGET_DETAIL ADD CONSTRAINT  TARGET_DETAIL_FK FOREIGN KEY (TARGET_ID) REFERENCES TARGET "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_3_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE TARGET_DETAIL ADD CONSTRAINT  TARGET_DETAIL_EMP_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_4_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_04_16_5_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_5_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  MASTER_VALID "
   SQL = SQL & "( "
   SQL = SQL & "MASTER_VALID_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MASTER_VALID_NO                   CODE_TYPE , "
   SQL = SQL & "MASTER_VALID_DESC                  DESC_TYPE , "
   SQL = SQL & "MASTER_VALID_TYPE                   ID_TYPE , "
   SQL = SQL & "VALID_FROM                  DATE_TYPE , "
   SQL = SQL & "VALID_TO                   DATE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME             CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MASTER_VALID ADD CONSTRAINT MASTER_VALID_ID_PK PRIMARY KEY (MASTER_VALID_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR MASTER_VALID_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_5_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_6_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_6_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER MASTER_VALID_BA FOR MASTER_VALID "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.MASTER_VALID_ID  = GEN_ID (MASTER_VALID_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_6_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_04_16_7_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_7_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  COMMISSION_SALE "
   SQL = SQL & "( "
   SQL = SQL & "COMMISSION_SALE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MASTER_VALID_ID                   ID_TYPE , "
   SQL = SQL & "COMMISSION_SALE_AREA                   ID_TYPE , "   ' 1.Com ได้ตอนเก็บเงิน 2.Com ได้เงินหักเงินทันที 3.Com ได้เงินหักเงินทันที-ผู้จัดการ
   SQL = SQL & "COMMISSION_SALE_TYPE                  ID_TYPE , "      'จ่ายเป็น % , รับหรือหักเป็น บาท
   SQL = SQL & "SELL_FROM                   MONEY_TYPE , "
   SQL = SQL & "SELL_TO                   MONEY_TYPE , "
   SQL = SQL & "COMMISSION_SALE_AMOUNT                   MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME             CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COMMISSION_SALE ADD CONSTRAINT COMMISSION_SALE_ID_PK PRIMARY KEY (COMMISSION_SALE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COMMISSION_SALE_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_7_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_8_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_8_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER COMMISSION_SALE_BA FOR COMMISSION_SALE "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.COMMISSION_SALE_ID  = GEN_ID (COMMISSION_SALE_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_8_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_9_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_9_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE COMMISSION_SALE ADD CONSTRAINT  COMMISSION_SALE_FK FOREIGN KEY (MASTER_VALID_ID) REFERENCES MASTER_VALID "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_9_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_10_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_10_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  COMMISSION_BUDGET_CHART "
   SQL = SQL & "( "
   SQL = SQL & "COMMISSION_BUDGET_CHART_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MASTER_VALID_ID         ID_TYPE , "
   SQL = SQL & "PARENT_ID                   ID_TYPE , "
   SQL = SQL & "EMP_ID                   ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME          CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COMMISSION_BUDGET_CHART ADD CONSTRAINT COMMISSION_BUDGET_CHART_ID_PK PRIMARY KEY (COMMISSION_BUDGET_CHART_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COMMISSION_BUDGET_CHART_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_10_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_11_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_11_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER COMMISSION_BUDGET_CHART_BA FOR COMMISSION_BUDGET_CHART "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.COMMISSION_BUDGET_CHART_ID  = GEN_ID (COMMISSION_BUDGET_CHART_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_11_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_12_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_12_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE COMMISSION_BUDGET_CHART ADD CONSTRAINT  COMMISSION_BUDGET_CHART_FK FOREIGN KEY (MASTER_VALID_ID) REFERENCES MASTER_VALID "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_12_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_13_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_13_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE COMMISSION_BUDGET_CHART ADD CONSTRAINT  COMMISSION_BUDGET_CHART_EMP_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_13_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_14_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_14_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE COMMISSION_BUDGET_CHART ADD CONSTRAINT  COMMISSION_BUDGET_CHART_PR_FK FOREIGN KEY (PARENT_ID) REFERENCES COMMISSION_BUDGET_CHART "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_14_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_15_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_15_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  COMMISSION_SUBTRACT "
   SQL = SQL & "( "
   SQL = SQL & "COMMISSION_SUBTRACT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MONTH_ID         ID_TYPE , "
   SQL = SQL & "YEAR_NO                   ID_TYPE , "
   SQL = SQL & "YYYYMM                   CODE_TYPE , "
   SQL = SQL & "EMP_ID                   ID_TYPE , "
   SQL = SQL & "CUSTOMER_ID                   ID_TYPE , "
   SQL = SQL & "COMMISSION_SUBTRACT_AMOUNT                   MONEY_TYPE , "
   SQL = SQL & "COMMISSION_SUBTRACT_DESC                   DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME          CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COMMISSION_SUBTRACT ADD CONSTRAINT COMMISSION_SUBTRACT_ID_PK PRIMARY KEY (COMMISSION_SUBTRACT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COMMISSION_SUBTRACT_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_15_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_16_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_16_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER COMMISSION_SUBTRACT_BA FOR COMMISSION_SUBTRACT "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.COMMISSION_SUBTRACT_ID  = GEN_ID (COMMISSION_SUBTRACT_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_16_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_17_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_17_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE COMMISSION_SUBTRACT ADD CONSTRAINT  COMMISSION_SUBTRACT_EMP_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_17_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_18_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_16_18_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE COMMISSION_SUBTRACT ADD CONSTRAINT  COMMISSION_SUBTRACT_CUS_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_18_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function


Public Function Patch_2017_04_16_19_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_19_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  COMMISSION_COST "
   SQL = SQL & "( "
   SQL = SQL & "COMMISSION_COST_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "MASTER_VALID_ID         ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                   ID_TYPE , "
   SQL = SQL & "COST_AMOUNT                   MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME          CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COMMISSION_COST ADD CONSTRAINT COMMISSION_COST_ID_PK PRIMARY KEY (COMMISSION_COST_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COMMISSION_COST_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_19_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_20_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_20_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER COMMISSION_COST_BA FOR COMMISSION_COST "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.COMMISSION_COST_ID  = GEN_ID (COMMISSION_COST_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_16_20_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_21_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_21_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE COMMISSION_COST ADD CONSTRAINT  COMMISSION_COST_PART_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_21_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_16_22_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_04_16_22_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE COMMISSION_COST ADD CONSTRAINT  COMMISSION_COST_FK FOREIGN KEY (MASTER_VALID_ID) REFERENCES MASTER_VALID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_04_16_22_jill = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function


Public Function Patch_2017_04_17_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_17_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_TYPE ADD COMMISSION_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_17_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_17_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_17_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "UPDATE PART_TYPE SET COMMISSION_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_17_2_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_04_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_04_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE TARGET_DETAIL ADD CONSTRAINT UQ_TARGET_DETAIL UNIQUE (TARGET_ID,EMP_ID)"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_04_25_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_05_01_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_05_01_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD YYYYMM CODE_TYPE "     '
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_05_01_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_05_01_2_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_05_01_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC SET YYYYMM = SUBSTR(DOCUMENT_DATE, 1, 7) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_05_01_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'Public Function Patch_2017_05_01_3_jill() As Boolean
'On Error GoTo ErrorHandler
''On Error Resume Next
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2017_05_01_3_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "                 CREATE TRIGGER BILLING_DOC_BI FOR BILLING_DOC "
'   SQL = SQL & "  ACTIVE BEFORE INSERT "
'   SQL = SQL & "  POSITION 0 "
'   SQL = SQL & "  AS "
'   SQL = SQL & "  BEGIN "
'   SQL = SQL & "        NEW.YYYYMM = LEFT(NEW.DOCUMENT_DATE,7); "
'   SQL = SQL & "  END; "
'   m_Conn.Execute (SQL)
'
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2017_05_01_3_jill = True
'   Exit Function
'
'ErrorHandler:
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2017_05_01_4_jill() As Boolean
'On Error GoTo ErrorHandler
''On Error Resume Next
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2017_05_01_4_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "                 CREATE TRIGGER BILLING_DOC_BU FOR BILLING_DOC "
'   SQL = SQL & "  ACTIVE BEFORE UPDATE "
'   SQL = SQL & "  POSITION 0 "
'   SQL = SQL & "  AS "
'   SQL = SQL & "  BEGIN "
'   SQL = SQL & "        NEW.YYYYMM = LEFT(NEW.DOCUMENT_DATE,7); "
'   SQL = SQL & "  END; "
'   m_Conn.Execute (SQL)
'
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2017_05_01_4_jill = True
'   Exit Function
'
'ErrorHandler:
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function

Public Function Patch_2017_05_08_1_jill() As Boolean
On Error GoTo ErrorHandler
'On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_05_08_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE COMMISSION_SUBTRACT ADD CONSTRAINT UQ_COMMISSION_SUBTRACT UNIQUE (YYYYMM,EMP_ID,CUSTOMER_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_05_08_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_07_05_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_07_05_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  COMMISSION_CREDIT "
   SQL = SQL & "( "
   SQL = SQL & "COMMISSION_CREDIT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID                   ID_TYPE , "
   SQL = SQL & "COMMISSION_CREDIT_AMOUNT                   ID_TYPE , "
   SQL = SQL & "COMMISSION_CREDIT_DESC                   DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_NAME          CODE_TYPE, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE, "
   SQL = SQL & "MODIFY_NAME          CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COMMISSION_CREDIT ADD CONSTRAINT COMMISSION_CREDIT_ID_PK PRIMARY KEY (COMMISSION_CREDIT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COMMISSION_CREDIT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_07_05_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_07_05_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_07_05_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER COMMISSION_CREDIT_BA FOR COMMISSION_CREDIT "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.COMMISSION_CREDIT_ID  = GEN_ID (COMMISSION_CREDIT_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_07_05_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_07_05_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_07_05_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE COMMISSION_CREDIT ADD CONSTRAINT  COMMISSION_CREDIT_CUSTOMER_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_07_05_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_05_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_05_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  WEIGHT_ACTUAL  "
   SQL = SQL & "( "
   SQL = SQL & "WEIGHT_ACTUAL_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_NO                   CODE_TYPE , "
   SQL = SQL & "WEIGHT_IN                  ID_TYPE , "
   SQL = SQL & "WEIGHT_OUT                  ID_TYPE , "
   SQL = SQL & "WEIGHT_ID                    CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE WEIGHT_ACTUAL ADD CONSTRAINT WEIGHT_ACTUAL_ID_PK PRIMARY KEY (WEIGHT_ACTUAL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR WEIGHT_ACTUAL_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_05_05_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_05_31_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_05_31_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE EMPLOYEE ADD EMP_RESIGN_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EMPLOYEE SET EMP_RESIGN_FLAG = 'N' "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_05_31_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_06_29_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_06_29_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE PLANNING ADD PLAN_VERSION ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PLANNING SET PLAN_VERSION = 1 "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_06_29_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_07_17_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_07_17_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE CUSTOMER ADD CHECK_CREDIT_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE CUSTOMER SET CHECK_CREDIT_FLAG = 'N' "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_07_17_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_07_17_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_07_17_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE USER_ACCOUNT ALTER COLUMN USER_STATUS TO USER_STATUS2 "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_07_17_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_07_21_01_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_07_21_01_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE JOB_WAREHOUSE "
   SQL = SQL & "( "
   SQL = SQL & "JOB_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "JOB_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "JOB_DATE          DATE_TYPE NOT NULL, "
   SQL = SQL & "TIME_SHIFT_ID         DATE_TYPE NOT NULL, "
   SQL = SQL & "JOB_DESC          DESC_TYPE , "
   SQL = SQL & "PROCESS_ID      ID_TYPE , "
   SQL = SQL & "APPROVED_BY      ID_TYPE , "
   SQL = SQL & "RESPONSE_BY      ID_TYPE , "
   SQL = SQL & "MIX_DATE   DATE_TYPE , " 'วันที่ผลิต
   SQL = SQL & "START_DATE      DATE_TYPE , "
   SQL = SQL & "FINISH_DATE      DATE_TYPE , "
   SQL = SQL & "SUCCESS_FLAG      FLAG_TYPE , "
   SQL = SQL & "INVENTORY_DOC_ID      ID_TYPE, "
   SQL = SQL & "JOB_DOC_TYPE      ID_TYPE, "
   SQL = SQL & "PART_ITEM_ID      ID_TYPE, "
   SQL = SQL & "STD_AMOUNT      MONEY_TYPE, "
   SQL = SQL & "ACTUAL_AMOUNT      MONEY_TYPE, "
   SQL = SQL & "FORMULA_ID      ID_TYPE, "
   SQL = SQL & "YYYYMM      CODE_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE JOB_INOUT_WAREHOUSE "
   SQL = SQL & "( "
   SQL = SQL & "JOB_INOUT_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "JOB_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID           ID_TYPE NOT NULL, " 'เบอร์สินค้า
   SQL = SQL & "PRODUCT_TYPE_ID      ID_TYPE, " 'ชนิดของอาหาร
   SQL = SQL & "LOT_NO      ID_TYPE, " 'Lot การผลิต
   SQL = SQL & "BATCH_NO          CODE_TYPE , " 'จำนวน แบท การผลิต
   SQL = SQL & "BIN_NO          ID_TYPE , " 'ถังเก็บอาหาร
   SQL = SQL & "PALLET_NO          ID_TYPE , " 'ที่วางอาหาร
   SQL = SQL & "LOCK_NO          ID_TYPE , " 'ล๊อคที่วางพาเลท
   SQL = SQL & "GOOD_AMOUNT      MONEY_TYPE, " 'จำนวนดี
   SQL = SQL & "LOSE_AMOUNT      MONEY_TYPE, " 'จำนวนเสีย
   SQL = SQL & "WEIGHT_PER_PACK  MONEY_TYPE , " 'น้ำหนักต่อถุง
   SQL = SQL & "PACK_AMOUNT   MONEY_TYPE , " 'น้ำหนักที่ผลิตได้ ทั้ง (ดี+เสีย) x น้ำหนักต่อถุง
   SQL = SQL & "REST_AMOUNT    MONEY_TYPE , " 'เศษที่เหลือจากการบรรจุ
   SQL = SQL & "PACK_DATE   DATE_TYPE , " 'วันที่ผลิต
   SQL = SQL & "TIME_PACK_BEGIN      TIME_TYPE, " 'เวลาเริ่มบรรจุ
   SQL = SQL & "TIME_PACK_END      TIME_TYPE, " 'เวลาหลังบรรจุ
   SQL = SQL & "TX_AMOUNT          MONEY_TYPE , " 'จำนวนที่บรรจุได้ทั้งหมด ของดี+ของเสีย
   SQL = SQL & "NOTE      DESC_TYPE, " 'หมายเหตุ
   
   SQL = SQL & "LOCATION_ID           ID_TYPE NOT NULL, " 'ที่จัดเก็บ
   SQL = SQL & "LINK_ID           ID_TYPE , "
   SQL = SQL & "TX_TYPE          CODE_TYPE , " 'ประเภทสินค้า I = สินค้าผลิตเสร็จ ,E=วัตถุดิบที่ใช้

   SQL = SQL & "SERIAL_NUMBER          NAME_TYPE , "
   SQL = SQL & "INOUT_REF          CODE_TYPE , "
   SQL = SQL & "FROM_FORMULA       ID_TYPE , "
   SQL = SQL & "AVG_PRICE       MONEY_TYPE , "
   SQL = SQL & "GROUP_NO   ID_TYPE , "
   SQL = SQL & "STD_AMOUNT    MONEY_TYPE , "
   SQL = SQL & "PARAM_ID   ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   
   SQL = "CREATE GENERATOR JOB_WAREHOUSE_SEQ "
   m_Conn.Execute (SQL)
   
  SQL = "CREATE GENERATOR JOB_INOUT_WAREHOUSE_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_WAREHOUSE ADD CONSTRAINT JOB_WAREHOUSE_PK PRIMARY KEY (JOB_ID) "
   m_Conn.Execute (SQL)
''
   SQL = "ALTER TABLE JOB_INOUT_WAREHOUSE ADD CONSTRAINT JOB_INOUT_WAREHOUSE_PK PRIMARY KEY (JOB_INOUT_ID) "
   m_Conn.Execute (SQL)
''
   SQL = " ALTER TABLE JOB_WAREHOUSE ADD CONSTRAINT PROCESS_ID_FK6 FOREIGN KEY (PROCESS_ID) REFERENCES PROCESS"
   m_Conn.Execute (SQL)
''
   SQL = "ALTER TABLE JOB_WAREHOUSE ADD CONSTRAINT APPROVED_BY_FK1 FOREIGN KEY (APPROVED_BY) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)
''
   SQL = "ALTER TABLE JOB_WAREHOUSE ADD CONSTRAINT RESPONSE_BY_FK1 FOREIGN KEY (RESPONSE_BY) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_07_21_01_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2017_08_09_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_09_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE LOCATION ADD LOCATION_GROUP ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE LOCATION SET LOCATION_GROUP =0  "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_09_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_14_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_14_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE JOB_INOUT ADD PRODUCT_TYPE_ID      ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD LOT_NO      ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD BIN_NO          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD PALLET_NO          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD LOCK_NO          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD GOOD_AMOUNT      MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT ADD LOSE_AMOUNT      MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD REST_AMOUNT    MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD PACK_DATE   DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD TIME_PACK_BEGIN      TIME_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD TIME_PACK_END      TIME_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT ADD NOTE      DESC_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_14_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1 " & Err.DESCRIPTION
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_14_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_14_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = " CREATE TABLE INVENTORY_WH_DOC  "
  SQL = SQL & "( "
  SQL = SQL & "INVENTORY_WH_DOC_ID ID_TYPE NOT NULL,"
  SQL = SQL & "DOCUMENT_NO CODE_TYPE NOT NULL,"
  SQL = SQL & "DOCUMENT_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "DOCUMENT_DESC DESC_TYPE,"
  SQL = SQL & "BILL_NO CODE_TYPE,"
  SQL = SQL & "DO_NO CODE_TYPE,"
  SQL = SQL & "TRUCK_NO CODE_TYPE,"
  SQL = SQL & "SUPPLIER_ID ID_TYPE,"
  SQL = SQL & "DELIVERY_ID ID_TYPE,"
  SQL = SQL & "DELIVERY_FEE MONEY_TYPE,"
  SQL = SQL & "SENDER_NAME NAME_TYPE,"
  SQL = SQL & "RECEIVE_NAME NAME_TYPE,"
  SQL = SQL & "DOCUMENT_TYPE ID_TYPE NOT NULL,"
  SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL,"
  SQL = SQL & "EMP_ID ID_TYPE,"
  SQL = SQL & "COMMIT_FLAG FLAG_TYPE,"
  SQL = SQL & "SALE_FLAG FLAG_TYPE,"
  SQL = SQL & "REASON_ID ID_TYPE,"
  SQL = SQL & "QUE_NO CODE_TYPE,"
  SQL = SQL & "ENTRY_DATE DATE_TYPE,"
  SQL = SQL & "EXIT_DATE DATE_TYPE,"
  SQL = SQL & "EXCEPTION_FLAG FLAG_TYPE,"
  SQL = SQL & "DEPARTMENT_ID ID_TYPE,"
  SQL = SQL & "ADJUST_FLAG FLAG_TYPE,"
  SQL = SQL & "CREDIT MONEY_TYPE,"
  SQL = SQL & "PR_NO CODE_TYPE,"
  SQL = SQL & "YYYYMM CODE_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE LOT_ITEM_WH  "
   SQL = SQL & "( "
  SQL = SQL & "LOT_ITEM_WH_ID ID_TYPE NOT NULL,"
  SQL = SQL & "INVENTORY_WH_DOC_ID ID_TYPE,"
  SQL = SQL & "PART_ITEM_ID ID_TYPE NOT NULL,"
  SQL = SQL & "LOCATION_ID ID_TYPE NOT NULL,"
  SQL = SQL & "ACTUAL_UNIT_PRICE MONEY_TYPE,"
  SQL = SQL & "INCLUDE_UNIT_PRICE MONEY_TYPE,"
  SQL = SQL & "TOTAL_ACTUAL_PRICE MONEY_TYPE,"
  SQL = SQL & "TOTAL_INCLUDE_PRICE NUMERIC(15, 7),"
  SQL = SQL & "PREVIOUS_PRICE MONEY_TYPE,"
  SQL = SQL & "PREVIOUS_AMOUNT MONEY_TYPE,"
  SQL = SQL & "NEW_PRICE MONEY_TYPE,"
  SQL = SQL & "TX_AMOUNT MONEY_TYPE,"
  SQL = SQL & "NEW_AMOUNT MONEY_TYPE,"
  SQL = SQL & "TRANSACTION_SEQ ID_TYPE NOT NULL,"
  SQL = SQL & "GUI_ID ID_TYPE,"
  SQL = SQL & "CALCULATE_FLAG FLAG_TYPE,"
  SQL = SQL & "TX_TYPE CODE_TYPE,"
  SQL = SQL & "LEFT_AMOUNT MONEY_TYPE,"
  SQL = SQL & "LOT_REF_ID ID_TYPE,"
  SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MULTIPLIER ID_TYPE,"
  SQL = SQL & "LAYOUT_ID ID_TYPE,"
  SQL = SQL & "LINK_ID ID_TYPE,"
  SQL = SQL & "PACKAGING_AMT MONEY_TYPE,"
  SQL = SQL & "ENTRY_WEIGHT MONEY_TYPE,"
  SQL = SQL & "EXIT_WEIGHT MONEY_TYPE,"
  SQL = SQL & "WEIGHT_AMOUNT MONEY_TYPE,"
  SQL = SQL & "PACKAGE_WEIGHT MONEY_TYPE,"
  SQL = SQL & "OTHER_WEIGHT MONEY_TYPE,"
  SQL = SQL & "PERCENT_HUMID MONEY_TYPE,"
  SQL = SQL & "HUMID_WEIGHT MONEY_TYPE,"
  SQL = SQL & "PACKAGING_ID ID_TYPE,"
  SQL = SQL & "SUPPLIER_WEIGHT MONEY_TYPE,"
  SQL = SQL & "CALCULATE_TYPE ID_TYPE,"
  SQL = SQL & "PACKAGE_AMOUNT MONEY_TYPE,"
  SQL = SQL & "ACTUAL_PKG_AMOUNT MONEY_TYPE,"
  SQL = SQL & "PUREXP_ID1 ID_TYPE,"
  SQL = SQL & "PUREXP_ID2 ID_TYPE,"
  SQL = SQL & "EXPENSE1 MONEY_TYPE,"
  SQL = SQL & "EXPENSE2 MONEY_TYPE,"
  SQL = SQL & "TOTAL_WEIGHT MONEY_TYPE,"
  SQL = SQL & "RAW_PRICE MONEY_TYPE,"
  SQL = SQL & "DISCOUNT_AMT MONEY_TYPE,"
  SQL = SQL & "RAW_TOT_PRICE MONEY_TYPE,"
  SQL = SQL & "ITEM_DESC DESC_TYPE,"
  SQL = SQL & "TO_DEPARTMENT ID_TYPE,"
  SQL = SQL & "EXTRA_NAME NAME_TYPE,"
  SQL = SQL & "SALE_TOT_PRICE MONEY_TYPE,"
  SQL = SQL & "CALCULATE_WEIGHT MONEY_TYPE,"
  SQL = SQL & "RAW_COST MONEY_TYPE,"
  SQL = SQL & "EXPENSE_COST MONEY_TYPE,"
  SQL = SQL & "TOTAL_NEW_PRICE MONEY_TYPE,"
  SQL = SQL & "BAG_RETURN FLAG_TYPE,"
  SQL = SQL & "ACTUAL_AMOUNT MONEY_TYPE,"
  SQL = SQL & "ACTUAL_PRICE MONEY_TYPE,"
  SQL = SQL & "CURRENT_AMOUNT MONEY_TYPE,"
  SQL = SQL & "CURRENT_PRICE MONEY_TYPE,"
  SQL = SQL & "NEED_TOTAL_AMOUNT MONEY_TYPE,"
  SQL = SQL & "NEED_TOTAL_PRICE MONEY_TYPE,"
  SQL = SQL & "NEED_AVG_PRICE MONEY_TYPE,"
  SQL = SQL & "MANUAL_PRICE FLAG_TYPE,"
  SQL = SQL & "EXPENSE_TYPE ID_TYPE,"
  SQL = SQL & "ITEM_DESC_ID ID_TYPE,"
  SQL = SQL & "AUTO_PRICE FLAG_TYPE,"
  SQL = SQL & "PROJECT_NAME_ID ID_TYPE "
  SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR INVENTORY_WH_DOC_SEQ "
   m_Conn.Execute (SQL)
   
  SQL = "CREATE GENERATOR LOT_ITEM_WH_SEQ "
   m_Conn.Execute (SQL)
   
   '***************************
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSTRAINT INVENTORY_WH_DOC_ID_PK PRIMARY KEY (INVENTORY_WH_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSTRAINT INVENTORY_WH_DOC_EMP_ID_FK2 FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE(EMP_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSTRAINT INVENTORY_WH_DOC_SUPPLIER_ID_FK FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER(SUPPLIER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSTRAINT INVENTORYWH_DOC_SUPPLIER_ID_FK FOREIGN KEY (DELIVERY_ID) REFERENCES SUPPLIER(SUPPLIER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSTRAINT INVENTORY_REASON_ID_FK2 FOREIGN KEY (REASON_ID) REFERENCES REASON(REASON_ID) "
   m_Conn.Execute (SQL)
   
   '***************************

   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_ID_PK PRIMARY KEY (LOT_ITEM_WH_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_DESC_ID_FK FOREIGN KEY (ITEM_DESC_ID) REFERENCES MASTER_REF(KEY_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_DOC_ID_FK FOREIGN KEY (INVENTORY_WH_DOC_ID) REFERENCES INVENTORY_WH_DOC(INVENTORY_WH_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_LAYOUT_ID_FK FOREIGN KEY (LAYOUT_ID) REFERENCES LAY_OUT(LAY_OUT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION(LOCATION_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_PACKAGING_ID_FK FOREIGN KEY (PACKAGING_ID) REFERENCES PACKAGING(PACKAGING_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_PART_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM(PART_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_PROJECT_NAME_ID_FK FOREIGN KEY (PROJECT_NAME_ID) REFERENCES MASTER_REF(KEY_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_REF_ID_FK FOREIGN KEY (LOT_REF_ID) REFERENCES LOT_ITEM_WH(LOT_ITEM_WH_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOT_ITEM_WH_TO_DEPARTMENT_FK FOREIGN KEY (TO_DEPARTMENT) REFERENCES LAY_OUT(LAY_OUT_ID) "
   m_Conn.Execute (SQL)
   

   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_14_2_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2017_08_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE JOB ADD INVENTORY_WH_DOC_ID ID_TYPE NOT NULL DEFAULT 0 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD CONSTRAINT INVENTORY_WH_DOC_ID_FK2 FOREIGN KEY (INVENTORY_WH_DOC_ID) REFERENCES INVENTORY_WH_DOC(INVENTORY_WH_DOC_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_15_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_15_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_15_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR LOT_NUMBER_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_15_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_16_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_16_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE JOB_INOUT ALTER COLUMN LOT_NO TYPE CODE_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_16_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_16_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_16_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOT_NO      CODE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOT_ITEM_WH ADD BIN_NO          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOT_ITEM_WH ADD PALLET_NO          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOCK_NO          ID_TYPE "
   m_Conn.Execute (SQL)


   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_16_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_21_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_21_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD GOOD_AMOUNT      MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOT_ITEM_WH ADD WEIGHT_PER_PACK      MONEY_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_21_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_24_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_24_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR LOAD_GOODS_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_24_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONDITION1      ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONDITION2          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONDITION3          ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD ENTRY_WEIGHT          MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD EXIT_WEIGHT          MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD TOTAL_WEIGHT          MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CUSTOMER_ID          ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD WEIGHT_NOTE          DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD NOTE          DESC_TYPE "
   m_Conn.Execute (SQL)

'  SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSTRAINT IV_WH_DOC_CUS_ID_FK2 FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID) "
'   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_28_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_28_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE LOCATION ADD MAX_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE LOCATION SET MAX_AMOUNT =0  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCATION ADD RX_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE LOCATION SET RX_AMOUNT =0  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOCATION ADD UNIT_ID ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_28_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_29_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_29_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = " CREATE TABLE PALLET_DOC  "
  SQL = SQL & "( "
  SQL = SQL & "PALLET_DOC_ID ID_TYPE NOT NULL,"
  SQL = SQL & "LOT_ITEM_WH_ID ID_TYPE NOT NULL,"
  SQL = SQL & "PALLET_DOC_NO ID_TYPE NOT NULL,"
  SQL = SQL & "CAPACITY_AMOUNT ID_TYPE,"
  SQL = SQL & "TRUCK_NO DESC_TYPE,"
  SQL = SQL & "EMP_ID ID_TYPE,"
  SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "CREATE GENERATOR PALLET_DOC_SEQ "
   m_Conn.Execute (SQL)
   
   '***************************
   SQL = "ALTER TABLE PALLET_DOC ADD CONSTRAINT PALLET_DOC_ID_PK PRIMARY KEY (PALLET_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PALLET_DOC ADD CONSTRAINT PL_LOT_ITEM_WH_ID_FK FOREIGN KEY (LOT_ITEM_WH_ID) REFERENCES LOT_ITEM_WH(LOT_ITEM_WH_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_29_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2017_08_31_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_31_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD FULL_PALLET_FROM       ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD FULL_PALLET_TO       ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD FULL_UNIT_PER_PALLET       ID_TYPE "
   m_Conn.Execute (SQL)
   
      SQL = "ALTER TABLE LOT_ITEM_WH ADD SCRAP_PALLET       ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD SCRAP_UNIT_PER_PALLET       ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_08_31_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_09_01_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_09_01_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = " CREATE TABLE LOT  "
  SQL = SQL & "( "
  SQL = SQL & "LOT_ID ID_TYPE NOT NULL,"
  SQL = SQL & "LOT_NO  DESC_TYPE NOT NULL,"
  SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "CREATE GENERATOR LOT_SEQ "
   m_Conn.Execute (SQL)
   
   '***************************
   SQL = "ALTER TABLE LOT ADD CONSTRAINT LOT_ID_PK PRIMARY KEY (LOT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP LOT_NO "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT DROP LOT_NO "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD LOT_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD BATCH_NO CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_09_01_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2017_09_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_09_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   SQL = "ALTER TABLE PALLET_DOC ADD LOT_ID       ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PALLET_DOC ADD TX_TYPE          CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_09_12_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_09_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_09_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
  SQL = "ALTER TABLE JOB_INOUT DROP PRODUCT_TYPE_ID "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT DROP BIN_NO "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT DROP PALLET_NO  "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT DROP LOCK_NO   "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT DROP GOOD_AMOUNT   "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE JOB_INOUT DROP LOSE_AMOUNT    "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT DROP REST_AMOUNT  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT DROP PACK_DATE    "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT DROP TIME_PACK_BEGIN     "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT DROP TIME_PACK_END     "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB_INOUT DROP NOTE     "
   m_Conn.Execute (SQL)
   
   
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_09_15_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1 " & Err.DESCRIPTION
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_09_15_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_09_15_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE LOT_ITEM_WH ADD PRODUCT_TYPE_ID      ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOSE_AMOUNT      MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD REST_AMOUNT    MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD PACK_DATE   DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD TIME_PACK_BEGIN      TIME_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD TIME_PACK_END      TIME_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD NOTE      DESC_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_09_15_2_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1 " & Err.DESCRIPTION
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_09_22_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_09_22_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC DROP LOT_ID     "
   m_Conn.Execute (SQL)
  
  SQL = "ALTER TABLE LOT_ITEM_WH ADD LOT_ID      ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_09_22_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1 " & Err.DESCRIPTION
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_09_23_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_09_23_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = " CREATE TABLE LOT_DOC  "
  SQL = SQL & "( "
  SQL = SQL & "LOT_DOC_ID ID_TYPE NOT NULL,"
  SQL = SQL & "LOT_ID  ID_TYPE NOT NULL,"
  SQL = SQL & "LOT_ITEM_WH_ID  ID_TYPE NOT NULL,"
  SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ") "
   m_Conn.Execute (SQL)


   SQL = "CREATE GENERATOR LOT_DOC_SEQ "
   m_Conn.Execute (SQL)
   
   '***************************
   SQL = "ALTER TABLE LOT_DOC ADD CONSTRAINT LOT_DOC_ID_PK PRIMARY KEY (LOT_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE LOT_DOC ADD CONSTRAINT LOT_ITEM_WH_ID_FK FOREIGN KEY (LOT_ITEM_WH_ID) REFERENCES LOT_ITEM_WH "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PALLET_DOC ALTER COLUMN LOT_ID  TO LOT_DOC_ID  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PALLET_DOC DROP CONSTRAINT PL_LOT_ITEM_WH_ID_FK    "
   m_Conn.Execute (SQL)

  SQL = "ALTER TABLE PALLET_DOC DROP LOT_ITEM_WH_ID    "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE PALLET_DOC ADD CONSTRAINT LOT_DOC_ID_FK FOREIGN KEY (LOT_DOC_ID) REFERENCES LOT_DOC "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP LOT_ID    "
   m_Conn.Execute (SQL)

   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_09_23_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2017_10_11_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_10_11_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD PROCESS_ID          ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD APPROVED_BY          ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD RESPONSE_BY          ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD START_DATE          DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD FINISH_DATE          DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD SUCCESS_FLAG          FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_10_11_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_10_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_10_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_DOC ADD LOT_DOC_ID_REF          ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_10_12_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_10_14_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_10_14_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD LOAD_FLAG          FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_10_14_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_10_18_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_10_18_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD PACK_AMOUNT          MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_10_18_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_10_31_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_10_31_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD BILLING_DOC_ID     ID_TYPE    NOT NULL "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_10_31_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_01_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_11_01_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD WEEK_CREDIT_LIMIT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_11_01_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_02_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_11_02_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ALTER COLUMN TRUCK_NO TYPE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_11_02_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_06_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_11_06_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  PACK_PRODUCTION "
   SQL = SQL & "( "
   SQL = SQL & "PACK_PRODUCTION_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "PACK_PRODUCTION_NO                   CODE_TYPE , "
   SQL = SQL & "PACK_PRODUCTION_AREA                   ID_TYPE , "
   SQL = SQL & "PACK_PRODUCTION_DATE                    DATE_TYPE , "
   SQL = SQL & "PACK_PRODUCTION_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE PACK_PRODUCTION ADD CONSTRAINT PACK_PRODUCTION_ID_PK PRIMARY KEY (PACK_PRODUCTION_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PACK_PRODUCTION_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE  PACK_PRODUCTION_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "PACK_PRODUCTION_ITEM_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "PACK_PRODUCTION_ID                   ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE , " 'เบอร์อาหาร
   SQL = SQL & "TX_AMOUNT                    MONEY_TYPE , " 'จำนวน
   SQL = SQL & "WEIGHT_PER_PACK                    ID_TYPE , " 'ขนาด
   SQL = SQL & "PACK_AMOUNT                    MONEY_TYPE , " 'จำนวนถุง
   SQL = SQL & "PALLET_LABEL_YELLOW                   CODE_TYPE , " 'ป้ายชี้บ่งสีเหลือง
   SQL = SQL & "PALLET_LABEL_GREEN                  CODE_TYPE , " 'ป้ายชี้บ่งสีเขียว
   SQL = SQL & "SEWING_THREAD                  ID_TYPE , " 'ด้ายที่เย็บ 1=ขาวหรือ 2=ขาว-แดง
   SQL = SQL & "NOTE                    DESC_TYPE , " 'หมายเหตุ
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACK_PRODUCTION_ITEM ADD CONSTRAINT PACK_PRODUCTION_ITEM_ID_PK PRIMARY KEY (PACK_PRODUCTION_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PACK_PRODUCTION_ITEM_SEQ  "
   m_Conn.Execute (SQL)
   
      
    m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_11_06_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTransF
   End If
End Function
Public Function Patch_2017_11_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_11_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD HEAD_PACK_NO ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_11_12_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_11_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_DOC ADD BIN_NO ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_11_15_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_11_27_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_11_27_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD EMP_CHECK_CAR_ID      ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD EMP_CHECK_PRODUCT_ID          ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_11_27_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_12_11_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_12_11_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PALLET_DOC ADD BALANCE_FLAG      FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  PALLET_DOC SET BALANCE_FLAG ='Y' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_12_11_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_12_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_12_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_ITEM ADD CUSTOMER_SALE_TYPE      ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_12_12_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_12_22_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_12_22_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PALLET_DOC ALTER COLUMN PALLET_DOC_NO TYPE DESC_TYPE"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_12_22_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_12_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_12_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR BALANCE_GOODS_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2017_12_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_01_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_01_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE LOT_DOC DROP LOT_DOC_ID_REF "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_DOC ADD LOT_DOC_ID_REF          ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_DOC ADD CONSTRAINT FK_LOT_DOC FOREIGN KEY (LOT_DOC_ID_REF) REFERENCES LOT_DOC(LOT_DOC_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_01_05_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_01_30_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_01_30_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE LOT_ITEM_WH DROP BILLING_DOC_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD BILLING_DOC_ID          ID_TYPE "
   m_Conn.Execute (SQL)
   

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_01_30_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_02_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_02_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CASH_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD CASH_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_02_15_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_02_20_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_02_20_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT ADD LOT_DATE  DATE_TYPE NOT NULL"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_DOC ADD CONSTRAINT LOT_ID_FK FOREIGN KEY (LOT_ID) REFERENCES LOT(LOT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PALLET_DOC ALTER COLUMN CAPACITY_AMOUNT TYPE MONEY_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_02_20_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_02_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_02_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE PART_ITEM SET CANCEL_FLAG = 'Y' WHERE PART_TYPE IN (21,42,20,19,15,10)  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_02_28_1_jill = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_03_02_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_03_02_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_DOC ADD BALANCE_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_03_02_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_03_06_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_03_06_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  INVENTORY_WH_ACT "
   SQL = SQL & "( "
   SQL = SQL & "INVENTORY_WH_ACT_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "INVENTORY_WH_ACT_AREA                   ID_TYPE , "
   SQL = SQL & "INVENTORY_WH_ACT_DATE                    DATE_TYPE , "
   SQL = SQL & "INVENTORY_WH_ACT_DESC                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_ACT ADD CONSTRAINT INVENTORY_WH_ACT_ID PRIMARY KEY (INVENTORY_WH_ACT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR INVENTORY_WH_ACT_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_03_06_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_03_06_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_03_06_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  INVENTORY_WH_ACT_ITEM  "
   SQL = SQL & "( "
   SQL = SQL & "INVENTORY_WH_ACT_ITEM_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "INVENTORY_WH_ACT_ID                   ID_TYPE , "
   SQL = SQL & "INVENTORY_WH_ACT_ITEM_TYPE                   ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                   ID_TYPE , "
   SQL = SQL & "INVENTORY_WH_SYS_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "INVENTORY_WH_ACT_AMOUNT                    MONEY_TYPE , "
   SQL = SQL & "NOTE                    DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_ACT_ITEM ADD CONSTRAINT INVENTORY_WH_ACT_ITEM_ID_PK PRIMARY KEY (INVENTORY_WH_ACT_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR INVENTORY_WH_ACT_ITEM_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_03_06_2_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_03_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_03_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_ITEM ADD ANIMAL_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_03_15_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_03_22_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_03_22_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD BL_START_DATE  DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_03_22_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_03_31_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_03_31_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE GENERATOR LOAD_GOODS_BK_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_03_31_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_04_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE JOB ADD INVENTORY_WH_DOC_ID_INPUT ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD CONSTRAINT INVENTORY_WH_DOC_ID_INPUT_FK2 FOREIGN KEY (INVENTORY_WH_DOC_ID_INPUT) REFERENCES INVENTORY_WH_DOC(INVENTORY_WH_DOC_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_05_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_04_06_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_06_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE JOB ADD FROM_BATCH_NO ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD TO_BATCH_NO ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD BATCH_TOTAL ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_06_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_04_07_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_07_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE GENERATOR RE_BAG_JOB_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR RE_BULK_JOB_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_04_07_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_2018_04_10_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_10_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE LOT_ITEM_WH DROP LOT_REF_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP MULTIPLIER "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP CALCULATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ACTUAL_UNIT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP INCLUDE_UNIT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP TOTAL_ACTUAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP TOTAL_INCLUDE_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PREVIOUS_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PREVIOUS_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP NEW_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP NEW_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP GUI_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP LEFT_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP LAYOUT_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP LINK_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PACKAGING_AMT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ENTRY_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP EXIT_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP WEIGHT_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PACKAGE_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP OTHER_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PERCENT_HUMID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP HUMID_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PACKAGING_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP SUPPLIER_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PACKAGE_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ACTUAL_PKG_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PUREXP_ID1 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PUREXP_ID2 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP EXPENSE1 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP EXPENSE2 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP TOTAL_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP RAW_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP DISCOUNT_AMT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP RAW_TOT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ITEM_DESC "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP TO_DEPARTMENT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP EXTRA_NAME "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP SALE_TOT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP CALCULATE_WEIGHT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP RAW_COST "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP EXPENSE_COST "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP TOTAL_NEW_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP BAG_RETURN "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ACTUAL_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ACTUAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP CURRENT_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP CURRENT_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP NEED_TOTAL_AMOUNT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP NEED_TOTAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP NEED_AVG_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP MANUAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP EXPENSE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP ITEM_DESC_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP AUTO_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PROJECT_NAME_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH DROP PALLET_NO "
   m_Conn.Execute (SQL)


   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_04_10_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2018_04_19_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_19_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOAD_AMOUNT_FLAG          FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOAD_TRUE          MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ALTER COLUMN LOAD_FLAG TYPE NAME_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_19_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_04_23_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_23_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT ADD CONSTRAINT UQ_LOT UNIQUE (LOT_NO)  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_23_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_04_23_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_23_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD INVENTORY_WH_DOC_ID ID_TYPE  "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT INVENTORY_WH_DOC_ID_FK3 FOREIGN KEY (INVENTORY_WH_DOC_ID) REFERENCES INVENTORY_WH_DOC(INVENTORY_WH_DOC_ID) "
   m_Conn.Execute (SQL)
   
'   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT BILLING_DOC_ID_FK2 FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) "
'   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_23_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'INVENTORY_WH_DOC_ID

Public Function Patch_2018_04_26_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_26_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  FREELANCE  "
   SQL = SQL & "( "
   SQL = SQL & "FREELANCE_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "FREELANCE_CODE                  CODE_TYPE , "
   SQL = SQL & "FREELANCE_NAME                  CODE_TYPE , "
   SQL = SQL & "FREELANCE_LASTNAME          CODE_TYPE , "
   SQL = SQL & "FREELANCE_RESIGN_FLAG     FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FREELANCE ADD CONSTRAINT FREELANCE_ID_PK PRIMARY KEY (FREELANCE_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR FREELANCE_SEQ  "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_26_1_dear = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_04_27_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_04_27_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE CUSTOMER ADD FREELANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD CONSTRAINT FREELANCE_ID_FK FOREIGN KEY (FREELANCE_ID) REFERENCES FREELANCE(FREELANCE_ID) "
   m_Conn.Execute (SQL)
       
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_04_27_1_dear = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_04_30_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_04_30_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  INCENTIVE  "
   SQL = SQL & "( "
   SQL = SQL & "INCENTIVE_ID                          ID_TYPE NOT NULL, "
   SQL = SQL & "FREELANCE_ID                        ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID                         ID_TYPE , "
   SQL = SQL & "INCENTIVE_PER_PACK             MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT INCENTIVE_ID_PK PRIMARY KEY (INCENTIVE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT INCENTIVE_FREELANCE_ID_FK FOREIGN KEY (FREELANCE_ID) REFERENCES FREELANCE(FREELANCE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT INCENTIVE_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM(PART_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR INCENTIVE_SEQ  "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_04_30_1_dear = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function


Public Function Patch_2018_05_03_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_05_03_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT UQ_INCENTIVE UNIQUE (FREELANCE_ID,PART_ITEM_ID) "
   m_Conn.Execute (SQL)
   
'   ALTER TABLE INCENTIVE ADD CONSTRAINT UQ_INCENTIVE UNIQUE (FREELANCE_ID,PART_ITEM_ID);
       
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_05_03_1_dear = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_05_14_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_05_14_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "UPDATE PART_ITEM SET CANCEL_FLAG='N' WHERE  CANCEL_FLAG IS NULL "
   m_Conn.Execute (SQL)
          
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_05_14_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_05_16_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_05_16_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "update rdb$relation_fields set rdb$null_flag = NULL where (rdb$field_name = 'BILLING_DOC_ID') and (rdb$relation_name = 'LOT_ITEM_WH') "
   m_Conn.Execute (SQL)

   SQL = "update lot_item_wh set billing_doc_id = null where billing_doc_id = 0 "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT FK_BILLING_DOC_ID_LIW FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC(BILLING_DOC_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_05_16_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_05_30_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_05_30_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE JOB ADD BATCH_DETAIL DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_05_30_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_06_04_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_06_04_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE GENERATOR BULK_JOBPLAN_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_06_04_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2018_06_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_06_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "ALTER TABLE JOB ADD LOCK_DOC_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_06_12_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2018_06_13_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_06_13_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE JOB ADD JOB_ID_REF ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD CONSTRAINT JOB_ID_REF_FK2 FOREIGN KEY (JOB_ID_REF) REFERENCES JOB(JOB_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_06_13_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_06_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_06_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  LOCK_DATE  "
   SQL = SQL & "( "
   SQL = SQL & "LOCK_DATE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "FROM_DATE                    DATE_TYPE , "
   SQL = SQL & "TO_DATE                    DATE_TYPE , "
   SQL = SQL & "FROM_INVENTORY_DATE                     DATE_TYPE , "
   SQL = SQL & "TO_INVENTORY_DATE                     DATE_TYPE , "
   SQL = SQL & "FROM_INVOICE_DATE                     DATE_TYPE , "
   SQL = SQL & "TO_INVOICE_DATE                     DATE_TYPE , "
   SQL = SQL & "FROM_RECEIPT_DATE                     DATE_TYPE , "
   SQL = SQL & "TO_RECEIPT_DATE                     DATE_TYPE , "
   SQL = SQL & "LOCK_TYPE                    ID_TYPE , "          ' 1 IS GLOBLE
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOCK_DATE  ADD CONSTRAINT LOCK_DATE_PK  PRIMARY KEY (LOCK_DATE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR LOCK_DATE_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_06_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_06_25_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_06_25_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "INSERT INTO LOCK_DATE (LOCK_DATE_ID,FROM_DATE,TO_DATE,FROM_INVENTORY_DATE,TO_INVENTORY_DATE,FROM_INVOICE_DATE,TO_INVOICE_DATE,FROM_RECEIPT_DATE,TO_RECEIPT_DATE,LOCK_TYPE,CREATE_DATE,CREATE_BY,MODIFY_DATE,MODIFY_BY) "
   SQL = SQL & "  VALUES "
   SQL = SQL & "  (1,'2012-08-01 00:00:00','2012-08-31 00:00:00','2012-08-01 00:00:00','2012-08-31 00:00:00','2012-08-01 00:00:00','2012-08-31 00:00:00','2012-08-01 00:00:00','2012-08-31 00:00:00',1,'2012-08-20 15:50:00',1,'2012-08-20 15:50:00',1) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_06_25_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_06_28_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_06_28_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  FREELANCE_ITEM  "
   SQL = SQL & "( "
   SQL = SQL & "FREELANCE_ITEM_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "FREELANCE_ID                 ID_TYPE , "
   SQL = SQL & "CUSTOMER_ID                  ID_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE FREELANCE_ITEM ADD CONSTRAINT FREELANCE_ITEM_ID_PK PRIMARY KEY (FREELANCE_ITEM_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_06_28_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_06_28_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_06_28_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE INCENTIVE DROP CONSTRAINT CUSTOMER_ID_ICT_FK "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE DROP CUSTOMER_ID  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE DROP DOCUMENT_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD CUSTOMER_ID ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INCENTIVE ADD DOCUMENT_TYPE ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE INCENTIVE SET DOCUMENT_TYPE = 1 "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE DROP CONSTRAINT INCENTIVE_ID_PK; "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE DROP CONSTRAINT INCENTIVE_FREELANCE_ID_FK;"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE DROP CONSTRAINT INCENTIVE_PART_ITEM_ID_FK;"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE DROP CONSTRAINT UQ_INCENTIVE;"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT INCENTIVE_ID_PK PRIMARY KEY (INCENTIVE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT INCENTIVE_FREELANCE_ID_FK FOREIGN KEY (FREELANCE_ID) REFERENCES FREELANCE(FREELANCE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT INCENTIVE_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM(PART_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT CUSTOMER_ID_ICT_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INCENTIVE ADD CONSTRAINT UQ_INCENTIVE UNIQUE (FREELANCE_ID,PART_ITEM_ID,CUSTOMER_ID) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_06_28_2_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_03_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_03_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD OUT_STOCK_FLAG          FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE LOT_ITEM_WH SET OUT_STOCK_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_07_03_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_09_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim Rs As ADODB.Recordset
Dim ItemCount As Long
Dim cCus As CCustomer
Dim FLI As CFreelanceItem

   VersionStr = "2018_07_09_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   Set Rs = New ADODB.Recordset
   
   Set cCus = New CCustomer
   cCus.CUSTOMER_ID = -1
   cCus.FREELANCE_ID = 0
   Call cCus.QueryData(Rs, ItemCount)
    
   Set cCus = Nothing
   Set cCus = New CCustomer
   While Not Rs.EOF
        Call cCus.PopulateFromRS(5, Rs)
        Set FLI = New CFreelanceItem
        FLI.AddEditMode = SHOW_ADD
        FLI.FREELANCE_ID = cCus.FREELANCE_ID
        FLI.CUSTOMER_ID = cCus.CUSTOMER_ID
        Call FLI.AddEditData
        Set FLI = Nothing
      Rs.MoveNext
   Wend
  Set cCus = Nothing

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_07_09_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_09_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_09_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER DROP FREELANCE_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_07_09_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_16_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_16_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD SUCCESS_FLAG          FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILLING_DOC SET SUCCESS_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_07_16_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_18_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_18_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE DOMAIN STATUS_TYPE AS CHAR(1) CHARACTER SET UNICODE_FSS DEFAULT 'N' CHECK (VALUE IN ('Y','N','A','I','C','L')) COLLATE UNICODE_FSS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_07_18_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_18_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_18_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM  ADD BILLING_DOC_SO_ID          ID_TYPE "
   m_Conn.Execute (SQL)

   
   SQL = "ALTER TABLE BILLING_DOC ALTER COLUMN SUCCESS_FLAG TYPE STATUS_TYPE"
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_07_18_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_07_24_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_24_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE GENERATOR RE_BAG_RM_JOB_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_07_24_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2018_07_26_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_07_26_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "UPDATE BILLING_DOC SET SUCCESS_FLAG='C' WHERE SUCCESS_FLAG='N' AND DOCUMENT_TYPE<>19 AND INVENTORY_WH_DOC_ID IS NOT NULL "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_07_26_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2018_08_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_08_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE VERSION_PROGRAM "
   SQL = SQL & "( "
   SQL = SQL & "VERSION_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "VERSION_NAME          CODE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY  ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE VERSION_PROGRAM ADD CONSTRAINT VERSION_ID_PK PRIMARY KEY (VERSION_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO VERSION_PROGRAM (VERSION_ID,VERSION_NAME,CREATE_DATE,CREATE_BY,MODIFY_DATE,MODIFY_BY) "
   SQL = SQL & "  VALUES "
   SQL = SQL & "  (1,'0.0.0.0','2012-08-20 15:50:00',1,'2012-08-20 15:50:00',1) "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_08_15_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2018_08_27_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_08_27_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  SUPPLIER_TRANSPORT  "
   SQL = SQL & "( "
   SQL = SQL & "SUPPLIER_TRANSPORT_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "SUPPLIER_TRANSPORT_CODE                CODE_TYPE , "
   SQL = SQL & "SUPPLIER_TRANSPORT_DETAIL                CODE_TYPE , "
   SQL = SQL & "SUPPLIER_ID                  ID_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SUPPLIER_TRANSPORT ADD CONSTRAINT SUPPLIER_TRANSPORT_ID_PK PRIMARY KEY (SUPPLIER_TRANSPORT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SUPPLIER_TRANSPORT ADD CONSTRAINT SUPPLIER_ID_FK5 FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_08_27_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_08_28_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_08_28_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD SUM_WITH_DO_ITEM_ID  ID_TYPE  "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT SUM_WITH_DO_ITEM_ID_FK FOREIGN KEY (SUM_WITH_DO_ITEM_ID) REFERENCES DO_ITEM(DO_ITEM_ID) "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_08_28_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_08_30_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_08_30_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " DROP TABLE SUPPLIER_TRANSPORT  "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE  SUPPLIER_TRANSPORT  "
   SQL = SQL & "( "
   SQL = SQL & "SUPPLIER_TRANSPORT_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "SUPPLIER_TRANSPORT_CODE                CODE_TYPE , "
   SQL = SQL & "SUPPLIER_TRANSPORT_DETAIL                CODE_TYPE , "
   SQL = SQL & "SUPPLIER_ID                  ID_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SUPPLIER_TRANSPORT ADD CONSTRAINT SUPPLIER_TRANSPORT_ID_PK PRIMARY KEY (SUPPLIER_TRANSPORT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SUPPLIER_TRANSPORT ADD CONSTRAINT SUPPLIER_ID_FK5 FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SUPPLIER_TRANSPORT_SEQ "
   m_Conn.Execute (SQL)
   

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_08_30_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_08_31_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_08_31_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD SUPPLIER_TRANSPORT_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DO_ITEM ADD CONSTRAINT SUPPLIER_TRANSPORT_ID_FK FOREIGN KEY (SUPPLIER_TRANSPORT_ID) REFERENCES SUPPLIER_TRANSPORT "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_08_31_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_08_31_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_08_31_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD RATE_DRIVER_TRANSPORT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD RATE_CUSTOMER_TRANSPORT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD RATE_SALE_TRANSPORT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD RATE_FAC_TRANSPORT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD NOTE_TRANSPORT DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_08_31_2_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_09_20_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_09_20_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  BILL_TRANSPORT  "
   SQL = SQL & "( "
   SQL = SQL & "BILL_TRANSPORT_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "BILLING_DOC_ID                  ID_TYPE , " 'REF
   SQL = SQL & "SUPPLIER_TRANSPORT_ID ID_TYPE  , " 'REF
   SQL = SQL & "AROUND                  ID_TYPE , "
   SQL = SQL & "SHIPPING                  DESC_TYPE , "
   SQL = SQL & "SHIPPING_DATE           DATE_TYPE , "
   SQL = SQL & "FIX_RATE_ID           ID_TYPE, "
   SQL = SQL & "NOTE                DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BILL_TRANSPORT ADD CONSTRAINT BILL_TRANSPORT_ID_PK PRIMARY KEY (BILL_TRANSPORT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILL_TRANSPORT ADD CONSTRAINT BILLING_DOC_ID_FK3 FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILL_TRANSPORT ADD CONSTRAINT SUPPLIER_TRANSPORT_ID_FK2 FOREIGN KEY (SUPPLIER_TRANSPORT_ID) REFERENCES SUPPLIER_TRANSPORT "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BILL_TRANSPORT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_09_20_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_09_20_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_09_20_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  BILL_TRANSPORT_ITEM  "
   SQL = SQL & "( "
   SQL = SQL & "BILL_TRANSPORT_ITEM_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "BILL_TRANSPORT_ID                        ID_TYPE NOT NULL, " 'REF
   SQL = SQL & "BILL_TYPE_ID                ID_TYPE , " 'รายการ
   SQL = SQL & "WEIGHT_PER_UNIT                ID_TYPE , " 'น้ำหนัก/หน่วย
   SQL = SQL & "PACK_AMOUNT                 MONEY_TYPE , " 'จำนวนหน่วย
   SQL = SQL & "RATE_PER_UNIT    MONEY_TYPE , "
   SQL = SQL & "TOTAL_PRICE     MONEY_TYPE , "
   SQL = SQL & "CAL_DRIVER    FLAG_TYPE , "
   SQL = SQL & "CAL_CUSTOMER     FLAG_TYPE , "
   SQL = SQL & "NOTE                DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BILL_TRANSPORT_ITEM  ADD CONSTRAINT BILL_TRANSPORT_ITEM_ID_PK PRIMARY KEY (BILL_TRANSPORT_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILL_TRANSPORT_ITEM ADD CONSTRAINT BILL_TRANSPORT_ID_FK3 FOREIGN KEY (BILL_TRANSPORT_ID) REFERENCES BILL_TRANSPORT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILL_TRANSPORT_ITEM ADD CONSTRAINT BILL_TYPE_ID_FK FOREIGN KEY (BILL_TYPE_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BILL_TRANSPORT_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_09_20_2_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_09_28_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_09_28_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILL_TRANSPORT  ADD FIX_RATE_DETAIL          DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILL_TRANSPORT_ITEM ADD CAL_PRICE_IN_PRODUCT          FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILL_TRANSPORT_ITEM SET CAL_PRICE_IN_PRODUCT = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_09_28_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_10_04_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_10_04_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILL_TRANSPORT  ADD NOT_CAL_VAT    FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILL_TRANSPORT SET NOT_CAL_VAT = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_10_04_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_10_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_10_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INCENTIVE ADD UNIT_TYPE  ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INCENTIVE ADD FROM_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INCENTIVE ADD TO_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_10_05_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_10_08_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_10_08_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INCENTIVE DROP CONSTRAINT UQ_INCENTIVE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_10_08_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'
Public Function Patch_2018_10_10_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_10_10_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INCENTIVE ADD AMOUNT_OVER_FLAG  FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE INCENTIVE SET AMOUNT_OVER_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INCENTIVE ADD RATE_OVER  MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_10_10_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_10_17_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_10_17_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE FREELANCE ADD EMP_ID          ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE FREELANCE ADD CONSTRAINT EMP_ID_FK1  FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_10_17_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_11_19_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_11_19_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CONFIG_DOC ADD AUTO_BEGIN_FLAG   FLAG_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_11_19_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_11_26_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_11_26_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  SUPPLIER_ACCOUNT  "
   SQL = SQL & "( "
   SQL = SQL & "SUPPLIER_ACCOUNT_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "SUPPLIER_ACCOUNT_NO               CODE_TYPE , "
   SQL = SQL & "SUPPLIER_ACCOUNT_NAME               CODE_TYPE , "
   SQL = SQL & "SUPPLIER_ACCOUNT_BANK                CODE_TYPE , "
   SQL = SQL & "SUPPLIER_ACCOUNT_BRANCH                CODE_TYPE , "
   SQL = SQL & "SUPPLIER_ID                  ID_TYPE , "
   SQL = SQL & "USE_TRANSPORT_FLAG            FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SUPPLIER_ACCOUNT ADD CONSTRAINT SUPPLIER_ACCOUNT_ID_PK PRIMARY KEY (SUPPLIER_ACCOUNT_ID) "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE SUPPLIER_ACCOUNT ADD CONSTRAINT SUPPLIER_ID_FK6 FOREIGN KEY (SUPPLIER_ID) REFERENCES SUPPLIER "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SUPPLIER_ACCOUNT_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_11_26_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function


Public Function Patch_2018_11_20_28_lek_1() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_11_20_28_lek_1"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE  BILLING_PAYMENT "
   SQL = SQL & "( "
   SQL = SQL & "BILLING_PAYMENT_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_NO                  CODE_TYPE , " 'REF  DOCUMENT_TYPE
   SQL = SQL & "DOCUMENT_DATE           DATE_TYPE, "
   SQL = SQL & "DOCUMENT_TYPE           ID_TYPE, "
   SQL = SQL & "PAYMENT_TO                  DESC_TYPE , "
   SQL = SQL & "DOC_ASSEMBLE               DESC_TYPE , "
   SQL = SQL & "PAYMENT_COST               DESC_TYPE , "
   SQL = SQL & "PAYMENT_DUE           DATE_TYPE , "
   SQL = SQL & "PAYMENT_PART             DESC_TYPE , "
   SQL = SQL & "PAYMENT_DEPT            DESC_TYPE , "
   SQL = SQL & "PAYMENT_DEPT2           DESC_TYPE , "
   SQL = SQL & "PAYMENT_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "PAYMENT_BY          ID_TYPE , "
   SQL = SQL & "NOTE          DESC_TYPE , "
   SQL = SQL & "BILLING_PAYMENT_ID_REF          ID_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BILLING_PAYMENT ADD CONSTRAINT BILLING_PAYMENT_ID_PK PRIMARY KEY (BILLING_PAYMENT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILLING_PAYMENT ADD CONSTRAINT BILLING_PAYMENT_ID_REF_FK  FOREIGN KEY (BILLING_PAYMENT_ID_REF) REFERENCES BILLING_PAYMENT  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BILLING_PAYMENT_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_11_20_28_lek_1 = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_12_03_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_12_03_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_DETAIL ADD BILLING_PAYMENT_ID     ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE GL_DETAIL ADD CONSTRAINT BILLING_PAYMENT_ID_FK2  FOREIGN KEY (BILLING_PAYMENT_ID) REFERENCES BILLING_PAYMENT  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_12_03_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_12_11_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_12_11_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE MASTER_REF ADD SHOW_FLAG FLAG_TYPE "   '
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_12_11_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_12_18_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_12_18_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PRINT_COUNT         ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILLING_DOC SET PRINT_COUNT = NULL "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_12_18_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_12_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2018_12_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE TABLE EX_WORKS_PRICE "
   SQL = SQL & "( "
   SQL = SQL & "EX_WORKS_PRICE_ID  ID_TYPE NOT NULL, "
   SQL = SQL & "EX_WORKS_PRICE_CODE CODE_TYPE NOT NULL, "
   SQL = SQL & "EX_WORKS_PRICE_DATE DATE_TYPE, "
   SQL = SQL & "EX_WORKS_PRICE_DESC DESC_TYPE,  "
   SQL = SQL & "EX_WORKS_PRICE_STATUS ID_TYPE,  "
   SQL = SQL & "EX_WORKS_PRICE_LEVEL FLAG_TYPE,  "
   SQL = SQL & "EX_WORKS_PRICE_TYPE  ID_TYPE,  "
   SQL = SQL & "FROM_ACTIVE_DATE DATE_TYPE, "
   SQL = SQL & "TO_VALID_DATE DATE_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_WORKS_PRICE ADD CONSTRAINT EX_WORKS_PRICE_ID_PK PRIMARY KEY (EX_WORKS_PRICE_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR EX_WORKS_PRICE_SEQ "
   m_Conn.Execute (SQL)


SQL = "CREATE TABLE EX_WORKS_PRICE_ITEM ("
SQL = SQL & "EX_WORKS_PRICE_ITEM_ID ID_TYPE NOT NULL,"
SQL = SQL & "EX_WORKS_PRICE_ID ID_TYPE NOT NULL,"
SQL = SQL & "RATE_TYPE ID_TYPE NOT NULL,"
SQL = SQL & "PACKAGE_RATE MONEY_TYPE NOT NULL,"
SQL = SQL & "PART_ITEM_ID ID_TYPE,"
SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL);"
m_Conn.Execute (SQL)

SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD CONSTRAINT EX_WORKS_PRICE_ITEM_ID_PK PRIMARY KEY (EX_WORKS_PRICE_ITEM_ID);"
m_Conn.Execute (SQL)
''
SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD CONSTRAINT EX_WORKS_PRICE_ID_FK FOREIGN KEY (EX_WORKS_PRICE_ID) REFERENCES EX_WORKS_PRICE(EX_WORKS_PRICE_ID);"
m_Conn.Execute (SQL)

SQL = " ALTER TABLE EX_WORKS_PRICE_ITEM ADD CONSTRAINT EX_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM(PART_ITEM_ID);"
m_Conn.Execute (SQL)

SQL = "CREATE GENERATOR EX_WORKS_PRICE_ITEM_SEQ "
m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2018_12_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_01_08_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_08_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

SQL = "CREATE TABLE EX_DELIVERY_COST_ITEM ("
SQL = SQL & "EX_DELIVERY_COST_ITEM_ID ID_TYPE NOT NULL,"
SQL = SQL & "EX_WORKS_PRICE_ID ID_TYPE NOT NULL,"
SQL = SQL & "RATE_TYPE ID_TYPE NOT NULL,"
SQL = SQL & "RATE_DELIVERY MONEY_TYPE NOT NULL,"
SQL = SQL & "RATE_CUSTOMER MONEY_TYPE NOT NULL,"
SQL = SQL & "CUSTOMER_ID ID_TYPE,"
SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL);"
m_Conn.Execute (SQL)

SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD CONSTRAINT EX_DELIVERY_COST_ITEM_ID_PK PRIMARY KEY (EX_DELIVERY_COST_ITEM_ID);"
m_Conn.Execute (SQL)

SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD CONSTRAINT EX_WORKS_PRICE_ID_FK2 FOREIGN KEY (EX_WORKS_PRICE_ID) REFERENCES EX_WORKS_PRICE(EX_WORKS_PRICE_ID);"
m_Conn.Execute (SQL)

SQL = " ALTER TABLE EX_DELIVERY_COST_ITEM ADD CONSTRAINT EX_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID);"
m_Conn.Execute (SQL)

SQL = "CREATE GENERATOR EX_DELIVERY_COST_ITEM_SEQ "
m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_01_08_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2019_01_08_3_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_08_3_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  DELIVERY_CUS_ITEM  "
   SQL = SQL & "( "
   SQL = SQL & "DELIVERY_CUS_ITEM_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "DELIVERY_CUS_ITEM_CODE                 CODE_TYPE , "
   SQL = SQL & "DELIVERY_CUS_ITEM_NAME              DESC_TYPE , "
   SQL = SQL & "CUSTOMER_ID                  ID_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DELIVERY_CUS_ITEM ADD CONSTRAINT DELIVERY_CUS_ITEM_ID_PK PRIMARY KEY (DELIVERY_CUS_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DELIVERY_CUS_ITEM ADD CONSTRAINT CUSTOMER_ID_FK2 FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID);"
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR DELIVERY_CUS_ITEM_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_01_08_3_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_01_08_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_08_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD DELIVERY_CUS_ITEM_ID         ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE EX_DELIVERY_COST_ITEM ADD CONSTRAINT DELIVERY_CUS_ITEM_ID_FK FOREIGN KEY (DELIVERY_CUS_ITEM_ID) REFERENCES DELIVERY_CUS_ITEM(DELIVERY_CUS_ITEM_ID);"
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_01_08_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_01_23_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_23_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  PROMOTIONAL_ITEM  "
   SQL = SQL & "( "
   SQL = SQL & "PROMOTIONAL_ITEM_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "PROMOTIONAL_DETAIL_ID                  ID_TYPE , "
   SQL = SQL & "CUSTOMER_ID                  ID_TYPE , "
   SQL = SQL & "UNIT_TYPE                 ID_TYPE , "
   SQL = SQL & "PROMOTIONAL_RATE                 MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   


   SQL = "ALTER TABLE PROMOTIONAL_ITEM ADD CONSTRAINT PROMOTIONAL_ITEM_ID_ID_PK PRIMARY KEY (PROMOTIONAL_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PROMOTIONAL_ITEM ADD CONSTRAINT PROMOTIONAL_DETAIL_ID_FK  FOREIGN KEY (PROMOTIONAL_DETAIL_ID) REFERENCES MASTER_REF(KEY_ID);"
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PROMOTIONAL_ITEM ADD CONSTRAINT CUSTOMER_ID_FK3 FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID);"
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR PROMOTIONAL_ITEM_SEQ "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_01_23_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_01_24_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_24_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_COMMISSION_BAG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_CHEER_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_DST_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER1_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER2_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER3_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_COMMISSION_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_CHEER_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_DST_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER1_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER2_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER3_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRICE_THINK_TYPE  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   '*************************************
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_COMMISSION_BAG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_CHEER_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_DST_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_OTHER1_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_OTHER2_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_OTHER3_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_COMMISSION_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_CHEER_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_DST_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_OTHER1_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_OTHER2_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PRO_OTHER3_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_01_24_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_01_24_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_24_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD EX_WORKS_PRICE_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD EX_DELIVERY_COST_ITEM_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PACKAGE_RATE  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD RATE_CUSTOMER  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE BILLING_DOC ADD DELIVERY_CUS_ITEM_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PRICE_THINK_TYPE  ID_TYPE  "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE DO_ITEM ADD CONSTRAINT EX_WORKS_PRICE_ITEM_ID_FK  FOREIGN KEY (EX_WORKS_PRICE_ITEM_ID) REFERENCES EX_WORKS_PRICE_ITEM(EX_WORKS_PRICE_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DO_ITEM ADD CONSTRAINT EX_DELIVERY_COST_ITEM_ID_FK  FOREIGN KEY (EX_DELIVERY_COST_ITEM_ID) REFERENCES EX_DELIVERY_COST_ITEM(EX_DELIVERY_COST_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BILLING_DOC ADD CONSTRAINT DELIVERY_CUS_ITEM_ID_FK2  FOREIGN KEY (DELIVERY_CUS_ITEM_ID) REFERENCES DELIVERY_CUS_ITEM(DELIVERY_CUS_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_01_24_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_01_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_01_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD USER_APPLOVE_PRICE CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD USER_APPLOVE_PRICE_THINK CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_01_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_02_01_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_01_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILL_TRANSPORT_ITEM ADD CAL_DIRECT         FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILL_TRANSPORT_ITEM SET CAL_DIRECT = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_02_01_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_02_04_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_04_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD WEIGHT_PER_PACK  ID_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_02_04_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_02_05_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_05_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_COMMISSION_BAG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_CHEER_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_DST_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_OTHER1_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_OTHER2_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_OTHER3_BAG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_COMMISSION_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_CHEER_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_DST_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_OTHER1_KG MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_OTHER2_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PRO_OTHER3_KG  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
    SQL = "ALTER TABLE SALE_ORDER ADD EX_WORKS_PRICE_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD EX_DELIVERY_COST_ITEM_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD PACKAGE_RATE  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD RATE_CUSTOMER  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD USER_APPLOVE_PRICE CODE_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SALE_ORDER ADD CONSTRAINT EX_WORKS_PRICE_ITEM_ID_FK2  FOREIGN KEY (EX_WORKS_PRICE_ITEM_ID) REFERENCES EX_WORKS_PRICE_ITEM(EX_WORKS_PRICE_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE SALE_ORDER ADD CONSTRAINT EX_DELIVERY_COST_ITEM_ID_FK2  FOREIGN KEY (EX_DELIVERY_COST_ITEM_ID) REFERENCES EX_DELIVERY_COST_ITEM(EX_DELIVERY_COST_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_02_05_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_02_07_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_07_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD WEIGHT_ID      CODE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD ENTRY_WEIGHT_DATE          DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD ENTRY_WEIGHT_TIME         TIME_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD EXIT_WEIGHT_DATE         DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD EXIT_WEIGHT_TIME      TIME_TYPE "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_02_07_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_02_07_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_07_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD FEATURE_ID      ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "update rdb$relation_fields set  rdb$null_flag = NULL where (rdb$field_name = 'PART_ITEM_ID') and (rdb$relation_name = 'LOT_ITEM_WH');"
   m_Conn.Execute (SQL)
   
   SQL = "update rdb$relation_fields set rdb$null_flag = NULL where (rdb$field_name = 'LOCATION_ID') and (rdb$relation_name = 'LOT_ITEM_WH');"
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_02_07_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_02_11_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_11_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = " CREATE TABLE PART_ITEM_SELECT  "
  SQL = SQL & "( "
  SQL = SQL & "PART_ITEM_SELECT_ID ID_TYPE NOT NULL,"
  SQL = SQL & "PART_ITEM_SELECT_NO CODE_TYPE ,"
  SQL = SQL & "PART_ITEM_ID ID_TYPE ,"
  SQL = SQL & "CREATE_DATE DATE_TYPE ,"
  SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
  SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "CREATE GENERATOR PART_ITEM_SELECT_SEQ "
   m_Conn.Execute (SQL)
   
   '***************************
   SQL = "ALTER TABLE PART_ITEM_SELECT ADD CONSTRAINT PART_ITEM_SELECT_ID_PK PRIMARY KEY (PART_ITEM_SELECT_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_02_11_1_lek = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2019_02_13_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_02_13_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD CAL_RATE_DELIVERY_TYPE ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  CUSTOMER SET CAL_RATE_DELIVERY_TYPE =1  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_02_13_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2019_03_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
      
   VersionStr = "2019_03_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE FORMULA  ADD CANCEL_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE FORMULA SET CANCEL_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_03_12_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_03_22_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_03_22_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD RATE_TYPE_CUS ID_TYPE NOT NULL "
   m_Conn.Execute (SQL)
   
      SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD WEIGHT_PER_PACK_CUS ID_TYPE NOT NULL "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_03_22_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_03_26_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_03_26_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILL_TRANSPORT ADD CAL_RATE_DELIVERY_TYPE ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  BILL_TRANSPORT SET CAL_RATE_DELIVERY_TYPE =1  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILL_TRANSPORT ADD EX_DELIVERY_COST_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_03_26_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_03_29_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_03_29_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER1_NAME  DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER2_NAME  DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD PRO_OTHER3_NAME  DESC_TYPE  "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_03_29_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'Public Function Patch_2019_04_03_1_lek() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2019_04_03_1_lek"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE SALE_ORDER ADD CONSIGNMENT_FLAG  FLAG_TYPE  "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE SALE_ORDER SET CONSIGNMENT_FLAG = 'N'  "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE DO_ITEM ADD CONSIGNMENT_FLAG  FLAG_TYPE  "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE DO_ITEM SET CONSIGNMENT_FLAG = 'N'  "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2019_04_03_1_lek = True
'   Exit Function
'ErrorHandler:
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
Public Function Patch_2019_04_03_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_03_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD CONSIGNMENT_FLAG  FLAG_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE INVENTORY_WH_DOC SET CONSIGNMENT_FLAG = 'N'  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD INVENTORY_WH_DOC_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CONSTRAINT INVENTORY_WH_DOC_ID_FK4 FOREIGN KEY (INVENTORY_WH_DOC_ID) REFERENCES INVENTORY_WH_DOC(INVENTORY_WH_DOC_ID) "
   m_Conn.Execute (SQL)
   
 
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_03_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_07_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_07_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CUSTOMER_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE INVENTORY_DOC ADD CONSTRAINT CUSTOMER_ID_FK4 FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_07_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_07_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_07_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD LOCATION_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD CONSTRAINT LOCATION_ID_FK2 FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION(LOCATION_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_07_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_10_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_10_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "CREATE GENERATOR TRANSFER_RQ_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_10_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_10_2_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_10_2_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD APPROVE_MANAGE_NAME CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_10_2_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_24_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_24_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD INVENTORY_DOC_TRN_ID  ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT INVENTORY_DOC_TRN_ID_FK FOREIGN KEY (INVENTORY_DOC_TRN_ID) REFERENCES INVENTORY_DOC(INVENTORY_DOC_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_24_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_25_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_25_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD POST_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  BILLING_DOC SET POST_FLAG ='Y' WHERE DOCUMENT_TYPE=1 "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_04_25_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_04_26_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_04_26_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE EX_PROMOTION_PART_ITEM ("
   SQL = SQL & "EX_PROMOTION_PART_ITEM_ID ID_TYPE NOT NULL,"
   SQL = SQL & "EX_WORKS_PRICE_ID ID_TYPE NOT NULL,"
   SQL = SQL & "CUSTOMER_ID ID_TYPE,"
   SQL = SQL & "RATE_TYPE ID_TYPE NOT NULL,"
   SQL = SQL & "PART_ITEM_ID ID_TYPE,"
   SQL = SQL & "DISCOUNT_AMOUNT MONEY_TYPE,"
   SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
   SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
   SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL);"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD CONSTRAINT EX_PROMOTION_PART_ITEM_ID_PK PRIMARY KEY (EX_PROMOTION_PART_ITEM_ID);"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD CONSTRAINT EX_WORKS_PRICE_ID_FK3  FOREIGN KEY (EX_WORKS_PRICE_ID) REFERENCES EX_WORKS_PRICE(EX_WORKS_PRICE_ID);"
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE EX_PROMOTION_PART_ITEM  ADD CONSTRAINT CUSTOMER_ID_FK5 FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID);"
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE EX_PROMOTION_PART_ITEM ADD CONSTRAINT EX_PROMOTION_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM(PART_ITEM_ID);"
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR EX_PROMOTION_PART_ITEM_SEQ "
   m_Conn.Execute (SQL)



   SQL = "CREATE TABLE EX_PROMOTION_DLC_ITEM ("
   SQL = SQL & "EX_PROMOTION_DLC_ITEM_ID ID_TYPE NOT NULL,"
   SQL = SQL & "EX_WORKS_PRICE_ID ID_TYPE NOT NULL,"
   SQL = SQL & "CUSTOMER_ID ID_TYPE,"
   SQL = SQL & "RATE_TYPE_CUS ID_TYPE NOT NULL,"
   SQL = SQL & "WEIGHT_PER_PACK_CUS ID_TYPE NOT NULL,"
   SQL = SQL & "DELIVERY_CUS_ITEM_ID ID_TYPE,"
   SQL = SQL & "DISCOUNT_AMOUNT MONEY_TYPE,"
   SQL = SQL & "CREATE_DATE DATE_TYPE NOT NULL,"
   SQL = SQL & "CREATE_BY ID_TYPE NOT NULL,"
   SQL = SQL & "MODIFY_DATE DATE_TYPE NOT NULL,"
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL);"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD CONSTRAINT EX_PROMOTION_DLC_ITEM_ID_PK PRIMARY KEY (EX_PROMOTION_DLC_ITEM_ID);"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD CONSTRAINT EX_WORKS_PRICE_ID_FK4  FOREIGN KEY (EX_WORKS_PRICE_ID) REFERENCES EX_WORKS_PRICE(EX_WORKS_PRICE_ID);"
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE EX_PROMOTION_DLC_ITEM  ADD CONSTRAINT CUSTOMER_ID_FK6 FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER(CUSTOMER_ID);"
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE EX_PROMOTION_DLC_ITEM  ADD CONSTRAINT DELIVERY_CUS_ITEM_ID_FK3 FOREIGN KEY (DELIVERY_CUS_ITEM_ID) REFERENCES DELIVERY_CUS_ITEM(DELIVERY_CUS_ITEM_ID);"
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR EX_PROMOTION_DLC_ITEM_SEQ "
   m_Conn.Execute (SQL)
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_04_26_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_05_03_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_05_03_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD EX_PROMOTION_PART_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD EX_PROMOTION_DLC_ITEM_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_AMOUNT_PART  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_AMOUNT_DLC  MONEY_TYPE  "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE DO_ITEM ADD CONSTRAINT EX_PROMOTION_PART_ITEM_ID_FK2  FOREIGN KEY (EX_PROMOTION_PART_ITEM_ID) REFERENCES EX_PROMOTION_PART_ITEM(EX_PROMOTION_PART_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE DO_ITEM ADD CONSTRAINT EX_PROMOTION_DLC_ITEM_ID_FK  FOREIGN KEY (EX_PROMOTION_DLC_ITEM_ID) REFERENCES EX_PROMOTION_DLC_ITEM(EX_PROMOTION_DLC_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_05_03_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_05_10_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_05_10_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE SALE_ORDER ADD EX_PROMOTION_PART_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD EX_PROMOTION_DLC_ITEM_ID  ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD DISCOUNT_AMOUNT_PART  MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALE_ORDER ADD DISCOUNT_AMOUNT_DLC  MONEY_TYPE  "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE SALE_ORDER ADD CONSTRAINT EX_PROMOTION_PART_ITEM_ID_FK3  FOREIGN KEY (EX_PROMOTION_PART_ITEM_ID) REFERENCES EX_PROMOTION_PART_ITEM(EX_PROMOTION_PART_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE SALE_ORDER ADD CONSTRAINT EX_PROMOTION_DLC_ITEM_ID_FK2  FOREIGN KEY (EX_PROMOTION_DLC_ITEM_ID) REFERENCES EX_PROMOTION_DLC_ITEM(EX_PROMOTION_DLC_ITEM_ID)  "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_05_10_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_05_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_05_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOCATION_REST_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOT_ITEM_WH ADD LOCATION_LOSE_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOCATION_REST_ID_FK  FOREIGN KEY (LOCATION_REST_ID) REFERENCES LOCATION(LOCATION_ID)  "
   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE LOT_ITEM_WH ADD CONSTRAINT LOCATION_LOSE_ID_FK  FOREIGN KEY (LOCATION_LOSE_ID) REFERENCES LOCATION(LOCATION_ID)  "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_05_15_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_05_16_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_05_16_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE TABLE  PART_ITEM_PICTURE "
   SQL = SQL & "( "
   SQL = SQL & "PART_ITEM_PICTURE_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID                    ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_PICTURE_TYPE             ID_TYPE , "
   SQL = SQL & "PART_ITEM_PICTURE_PATH                   DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM_PICTURE  ADD CONSTRAINT PART_ITEM_PICTURE_ID_PK  PRIMARY KEY (PART_ITEM_PICTURE_ID) "
   m_Conn.Execute (SQL)
      
   SQL = " ALTER TABLE PART_ITEM_PICTURE  ADD CONSTRAINT PART_ITEM_ID_FK4  FOREIGN KEY (PART_ITEM_ID ) REFERENCES PART_ITEM  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PART_ITEM_PICTURE_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD RECEIVED_BY_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CONSTRAINT  RECEIVED_BY_ID_FK FOREIGN KEY (RECEIVED_BY_ID) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_05_16_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_05_21_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_05_21_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "ALTER TABLE JOB ADD VERIFY_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE JOB ADD VERIFY_NAME CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE JOB SET VERIFY_FLAG = 'Y',VERIFY_NAME='' "
   m_Conn.Execute (SQL)
   
'   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD VERIFY_FLAG FLAG_TYPE "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE INVENTORY_WH_DOC ADD VERIFY_NAME CODE_TYPE "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE INVENTORY_WH_DOC SET VERIFY_FLAG = 'Y',VERIFY_NAME='' WHERE DOCUMENT_TYPE=14 "
'   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_05_21_1_lek = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_05_23_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_05_23_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD CAL_PRICE_PART_CENTER_FLAG FLAG_TYPE DEFAULT 'Y' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD CAL_PRICE_DLC_CENTER_FLAG FLAG_TYPE DEFAULT 'Y' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE CUSTOMER SET CAL_PRICE_PART_CENTER_FLAG = 'Y' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE CUSTOMER SET CAL_PRICE_DLC_CENTER_FLAG = 'Y' "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_05_23_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_06_04_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_06_04_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD EDIT_PRICE_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE BILLING_DOC SET EDIT_PRICE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE FEATURE ADD BILL_DIRECT_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE FEATURE SET BILL_DIRECT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_06_04_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'Public Function Patch_2019_06_06_1_lek() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2019_06_06_1_lek"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE FEATURE ADD BILL_DIRECT_FLAG FLAG_TYPE DEFAULT 'N' "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE FEATURE SET BILL_DIRECT_FLAG = 'N' "
'   m_Conn.Execute (SQL)
'
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2019_06_06_1_lek = True
'   Exit Function
'ErrorHandler:
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
'      glbErrorLog.ShowUserError
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
Public Function Patch_2019_06_18_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_06_18_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE USER_ACCOUNT ALTER COLUMN USER_STATUS1 TO USER_STATUS2 "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_06_18_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_06_20_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_06_20_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  PART_MASTER  "
   SQL = SQL & "( "
   SQL = SQL & "PART_MASTER_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "PART_MASTER_NO               CODE_TYPE , "
   SQL = SQL & "PART_MASTER_NAME               CODE_TYPE , "
   SQL = SQL & "PART_MASTER_TYPE                  ID_TYPE , "
   SQL = SQL & "CANCEL_FLAG            FLAG_TYPE , "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PART_MASTER ADD CONSTRAINT PART_MASTER_ID_PK PRIMARY KEY (PART_MASTER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PART_MASTER_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD PART_MASTER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PART_ITEM ADD CONSTRAINT PART_MASTER_ID_FK FOREIGN KEY (PART_MASTER_ID) REFERENCES PART_MASTER "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_06_20_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_06_26_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_06_26_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "CREATE TABLE  PART_CUS_GROUPS  "
   SQL = SQL & "( "
   SQL = SQL & "PART_CUS_GROUPS_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "CUS_GROUPS_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "PART_MASTER_ID                        ID_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE                        DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY                             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE                         DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY                             ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PART_CUS_GROUPS ADD CONSTRAINT PART_CUS_GROUPS_ID_PK PRIMARY KEY (PART_CUS_GROUPS_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PART_CUS_GROUPS_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PART_CUS_GROUPS ADD CONSTRAINT CUS_GROUPS_ID_FK  FOREIGN KEY (CUS_GROUPS_ID) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE PART_CUS_GROUPS ADD CONSTRAINT PART_MASTER_ID_FK2 FOREIGN KEY (PART_MASTER_ID) REFERENCES PART_MASTER "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM DROP CUSTOMER_SALE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PART_MASTER ADD ANIMAL_TYPE ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_06_26_1_lek = True
   Exit Function

ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_03_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_07_03_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE PART_ITEM  ADD PRODUCT_TYPE_ID ID_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_07_03_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_07_26_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_07_26_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ALTER COLUMN TEMP_DO_NO TYPE DESC_TYPE "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_07_26_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_08_15_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_08_15_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE CUSTOMER  ADD FREE_PRICE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE  CUSTOMER  SET FREE_PRICE_FLAG ='N' "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2019_08_15_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_11_04_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_11_04_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE BILLING_DOC ADD SUPPLIER_ID_TRUE   ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD SUPPLIER_ID_TRUE   ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CONSTRAINT SUPPLIER_ID_TRUE_FK FOREIGN KEY (SUPPLIER_ID_TRUE) REFERENCES SUPPLIER(SUPPLIER_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_11_04_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2019_11_12_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2019_11_12_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE JOB ADD NOTE   DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2019_11_12_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2020_03_27_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2020_03_27_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE DELIVERY_CUS_ITEM ADD HIDE_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE DELIVERY_CUS_ITEM SET HIDE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2020_03_27_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2020_05_28_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2020_05_28_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE EX_WORKS_PRICE ADD EX_WORKS_PRICE_NOTE   STRING_TYPE "
   m_Conn.Execute (SQL)

'EX_WORKS_PRICE_ITEM
   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD APPROVED_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD APPROVED_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_WORKS_PRICE_ITEM SET APPROVED_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD VERIFY_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD VERIFY_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_WORKS_PRICE_ITEM SET VERIFY_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   'EX_DELIVERY_COST_ITEM
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD APPROVED_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD APPROVED_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_DELIVERY_COST_ITEM SET APPROVED_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD VERIFY_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD VERIFY_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_DELIVERY_COST_ITEM SET VERIFY_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
  'EX_PROMOTION_PART_ITEM
   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD APPROVED_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD APPROVED_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_PROMOTION_PART_ITEM SET APPROVED_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD VERIFY_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD VERIFY_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_PROMOTION_PART_ITEM SET VERIFY_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
     'EX_PROMOTION_DLC_ITEM
   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD APPROVED_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD APPROVED_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_PROMOTION_DLC_ITEM SET APPROVED_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD VERIFY_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD VERIFY_NAME   CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EX_PROMOTION_DLC_ITEM SET VERIFY_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD LAST_EDIT_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD LAST_EDIT_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD LAST_EDIT_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD LAST_EDIT_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_WORKS_PRICE_ITEM SET LAST_EDIT_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_DELIVERY_COST_ITEM SET LAST_EDIT_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_PROMOTION_PART_ITEM SET LAST_EDIT_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_PROMOTION_DLC_ITEM SET LAST_EDIT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   

   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD DECLARE_NEW_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD DECLARE_NEW_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD DECLARE_NEW_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD DECLARE_NEW_FLAG   FLAG_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_WORKS_PRICE_ITEM SET DECLARE_NEW_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_DELIVERY_COST_ITEM SET DECLARE_NEW_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_PROMOTION_PART_ITEM SET DECLARE_NEW_FLAG = 'N' "
   m_Conn.Execute (SQL)

   SQL = "UPDATE EX_PROMOTION_DLC_ITEM SET DECLARE_NEW_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD GP_VALUE  MONEY_TYPE "
   m_Conn.Execute (SQL)


   SQL = "ALTER TABLE EX_WORKS_PRICE_ITEM ADD DECLARE_COUNT   ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_DELIVERY_COST_ITEM ADD DECLARE_COUNT   ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_PART_ITEM ADD DECLARE_COUNT   ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EX_PROMOTION_DLC_ITEM ADD DECLARE_COUNT   ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2020_05_28_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2020_06_20_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2020_06_20_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD EX_PROMOTION_EXTRA_PART_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_EXTRA_AMOUNT_PART  MONEY_TYPE  "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SALE_ORDER ADD EX_PROMOTION_EXTRA_PART_ITEM_ID ID_TYPE  "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SALE_ORDER ADD DISCOUNT_EXTRA_AMOUNT_PART  MONEY_TYPE  "
   m_Conn.Execute (SQL)


   SQL = " ALTER TABLE DO_ITEM ADD CONSTRAINT EX_PMT_EXTRA_PART_ITEM_ID_FK  FOREIGN KEY (EX_PROMOTION_EXTRA_PART_ITEM_ID) REFERENCES EX_PROMOTION_PART_ITEM(EX_PROMOTION_PART_ITEM_ID)  "
   m_Conn.Execute (SQL)

   SQL = " ALTER TABLE SALE_ORDER ADD CONSTRAINT EX_PMT_EXTRA_PART_ITEM_ID_FK2 FOREIGN KEY (EX_PROMOTION_EXTRA_PART_ITEM_ID) REFERENCES EX_PROMOTION_PART_ITEM(EX_PROMOTION_PART_ITEM_ID)  "
   m_Conn.Execute (SQL)
   

   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2020_06_20_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2020_06_29_1_lek() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2020_06_29_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

  SQL = "ALTER TABLE USER_ACCOUNT ALTER COLUMN USER_STATUS2 TO USER_STATUS3 "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_2020_06_29_1_lek = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
