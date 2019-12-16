VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBacthing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Flag As String
Public SplitFlag As String
Public ProductionId As Long

Public PlanCode As String
Public ProductionDate As String
Public ProductionNumber As String
Public TempProductionNumber As String
Public BatchNumber As String
Public FormulaCode As String
Public FormulaName As String
Public FormulaDate As String
Public BatchStartDate As String
Public BatchEndDate As String
Public DestinationBin As String
Public ProductionWeight As Double
Public FromBatch As Double
Public ToBatch As Double
Public tempFromBatch As Double
Public tempToBatch As Double
Public tempBatchDetail  As String
Public TotalBatch As Double
Public BatchDetail  As String
Public TargetDryMix  As Double
Public TargetWetMix  As Double
Public TargetAfterWetMix  As Double
Public ActualDryMix  As Double
Public ActualWetMix  As Double
Public ActualAfterWetMix  As Double
Public RuningIngredient  As Double
Public IngredientCode As String
Public IngredientName  As String
Public IngredientType As String
Public BinCode As String
Public IngredientTargetWeight As String
Public IngredientActualWeight As String
Public IngredientDeviationWeight As String
Public LotNo As String
Public LotId As Long
Public BIN_NO As Long
Public BIN_NAME As String
Public SKIP As Boolean
Public SKIP_PART_ITEM_NO  As Boolean
Public JOB_ID As Long

