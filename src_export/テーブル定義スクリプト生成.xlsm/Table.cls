VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tableName      As String
Public tableComment   As String

Public columns As ValCollection

Public constraintsPk  As ConstraintPrimary
Public constraintsUks As ValCollection
Public constraintsFks As ValCollection
Public indexes As ValCollection

