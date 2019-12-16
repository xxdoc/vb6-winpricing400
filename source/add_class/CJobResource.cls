VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJobResource"
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

Public JOB_RESOURCE_ID As Long
Public JOB_ID As Long
Public MACHINE_ID As Long
Public MACHINE_ID_HOUR As String
Public MACHINE_ID_HOURN As String
Public EMP_ID As Long
Public EMP_ID_HOUR As String
Public EMP_ID_HOURN As String
Public OCCUPY_DATE As Date
Public OCCUPY_INTERVAL As Double

Public MACHINE_NO As String
Public MACHINE_NAME As String
Public LONG_NAME  As String
Public LAST_NAME  As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public POSITION_ID As Long
Public POSITION_NAME As String
Public id As String

Public PEOPLE As Long
Public HOUR As Long
Public START_DATE As Date
Public FINISH_DATE As Date

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT RES.*,ME.MACHINE_NO,ME.MACHINE_NAME,NM.LONG_NAME,NM.LAST_NAME,PS.POSITION_DESC,PS.POSITION_ID "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM JOB_RESOURCE RES "
   FromStr = FromStr & "LEFT OUTER JOIN MACHINE ME ON (RES.MACHINE_ID = ME.MACHINE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (EM.EMP_ID = RES.EMP_ID) "
    FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EMN ON (EMN.EMP_ID = EM.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EMN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMP_POSITION PS ON (EM.CURRENT_POSITION = PS.POSITION_ID) "
   
   If JOB_RESOURCE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_RESOURCE_ID = " & JOB_RESOURCE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_RESOURCE_ID = " & JOB_RESOURCE_ID & ") "
      End If
   End If
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_ID = " & JOB_ID & ") "
      End If
   End If
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MACHINE_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MACHINE_NAME "
   Else
      OrderByStr = "ORDER BY JOB_RESOURCE_ID "
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
Public Function QueryDataEmp(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim GroupByStr As String
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = " SELECT J.START_DATE,J.FINISH_DATE,COUNT(RES.EMP_ID)  PEOPLE,SUM(RES.OCCUPY_INTERVAL)  OCCUPY_INTERVAL "
   SelectStr2 = " SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = " FROM JOB_RESOURCE RES "
   FromStr = FromStr & " LEFT OUTER JOIN JOB J ON ( J.JOB_ID = RES.JOB_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN MACHINE ME ON (RES.MACHINE_ID = ME.MACHINE_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN EMPLOYEE EM ON (EM.EMP_ID = RES.EMP_ID) "
    FromStr = FromStr & " LEFT OUTER JOIN EMPLOYEE_NAME EMN ON (EMN.EMP_ID = EM.EMP_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN NAME NM ON (EMN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & " LEFT OUTER JOIN EMP_POSITION PS ON (EM.CURRENT_POSITION = PS.POSITION_ID) "
   
   If JOB_RESOURCE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JOB_RESOURCE_ID = " & JOB_RESOURCE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JOB_RESOURCE_ID = " & JOB_RESOURCE_ID & ") "
      End If
   End If
   If JOB_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (J.JOB_ID = " & JOB_ID & ") "
      Else
         WhereStr = WhereStr & "AND (J.JOB_ID = " & JOB_ID & ") "
      End If
   End If

   GroupByStr = ""

   If GroupByStr = "" Then
   GroupByStr = " GROUP BY J.START_DATE , J.FINISH_DATE  "
  Else
    GroupByStr = GroupByStr & " , BY J.START_DATE , J.FINISH_DATE "
  End If

   OrderByStr = ""
   OrderTypeStr = ""

   
   OrderByStr = "ORDER BY J.START_DATE "
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   '''Debug.Print (SQL1)
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Temp As Long, Rs As ADODB.Recordset)
   If Temp = 1 Then
    JOB_RESOURCE_ID = NVLI(Rs("JOB_RESOURCE_ID"), -1)
    JOB_ID = NVLI(Rs("JOB_ID"), -1)
   MACHINE_ID = NVLI(Rs("MACHINE_ID"), -1)
    MACHINE_ID_HOUR = NVLS(Rs("MACHINE_ID_HOUR"), "")
    MACHINE_ID_HOURN = NVLS(Rs("MACHINE_ID_HOURN"), "")
    OCCUPY_DATE = NVLS(Rs("OCCUPY_DATE"), "")
    MACHINE_NO = NVLS(Rs("MACHINE_NO"), "")
    MACHINE_NAME = NVLS(Rs("MACHINE_NAME"), "")
    OCCUPY_INTERVAL = NVLD(Rs("OCCUPY_INTERVAL"), -1)
  ElseIf Temp = 3 Then
   PEOPLE = NVLI(Rs("PEOPLE"), -1)
   OCCUPY_INTERVAL = NVLD(Rs("OCCUPY_INTERVAL"), -1)
   START_DATE = NVLS(Rs("START_DATE"), "")
   FINISH_DATE = NVLS(Rs("FINISH_DATE"), "")

  Else
    JOB_RESOURCE_ID = NVLI(Rs("JOB_RESOURCE_ID"), -1)
    JOB_ID = NVLI(Rs("JOB_ID"), -1)
    EMP_ID = NVLI(Rs("EMP_ID"), -1)
    EMP_ID_HOUR = NVLS(Rs("EMP_ID_HOUR"), "")
    EMP_ID_HOURN = NVLS(Rs("EMP_ID_HOURN"), "")
    POSITION_ID = NVLI(Rs("POSITION_ID"), -1)
    POSITION_NAME = NVLS(Rs("POSITION_DESC"), "")
    OCCUPY_DATE = NVLS(Rs("OCCUPY_DATE"), "")
    LONG_NAME = NVLS(Rs("LONG_NAME"), "")
   LAST_NAME = NVLS(Rs("LAST_NAME"), "")
   OCCUPY_INTERVAL = NVLD(Rs("OCCUPY_INTERVAL"), -1)
  End If
   End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim id As Long

Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = MACHINE_ID
   If MACHINE_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   TempID2 = EMP_ID
   If EMP_ID <= 0 Then
      TempID2 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("JOB_RESOURCE_SEQ", id, glbErrorLog)
      SQL1 = "INSERT INTO JOB_RESOURCE " & _
                     "( " & _
                     "   JOB_RESOURCE_ID, " & _
                     "   JOB_ID, " & _
                     "   MACHINE_ID, " & _
                     "   MACHINE_ID_HOUR, " & _
                     "   MACHINE_ID_HOURN, " & _
                     "   EMP_ID, " & _
                     "   EMP_ID_HOUR, " & _
                     "   EMP_ID_HOURN, " & _
                     "   OCCUPY_DATE, " & _
                     "   OCCUPY_INTERVAL, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         id & "," & _
                            "'" & JOB_ID & "'," & _
                         TempID1 & "," & _
                         "'" & ChangeQuote(Trim(MACHINE_ID_HOUR)) & "'," & _
                         "'" & ChangeQuote(Trim(MACHINE_ID_HOURN)) & "'," & _
                        TempID2 & "," & _
                        "'" & ChangeQuote(Trim(EMP_ID_HOUR)) & "'," & _
                        "'" & ChangeQuote(Trim(EMP_ID_HOURN)) & "'," & _
                        "'" & ChangeQuote(Trim(DateToStringInt(OCCUPY_DATE))) & "'," & _
                        OCCUPY_INTERVAL & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         JOB_RESOURCE_ID = id
   Else 'Show edit
      SQL1 = "UPDATE JOB_RESOURCE SET " & _
                       "   MACHINE_ID = " & TempID1 & "," & _
                      "   MACHINE_ID_HOUR = '" & ChangeQuote(Trim(MACHINE_ID_HOUR)) & "', " & _
                      "   MACHINE_ID_HOURN = '" & ChangeQuote(Trim(MACHINE_ID_HOURN)) & "', " & _
                      "   EMP_ID = " & TempID2 & "," & _
                     "   EMP_ID_HOUR = '" & ChangeQuote(Trim(EMP_ID_HOUR)) & "', " & _
                     "   EMP_ID_HOURN = '" & ChangeQuote(Trim(EMP_ID_HOURN)) & "', " & _
                     "   OCCUPY_DATE= '" & ChangeQuote(Trim(DateToStringInt(OCCUPY_DATE))) & "', " & _
                     "   OCCUPY_INTERVAL = " & OCCUPY_INTERVAL & "," & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE JOB_RESOURCE_ID = " & JOB_RESOURCE_ID
   End If
   '''Debug.Print (SQL1)
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
     SQL1 = "DELETE FROM JOB_RESOURCE WHERE JOB_ID = " & JOB_ID
      m_Conn.Execute (SQL1)
      
End Function




