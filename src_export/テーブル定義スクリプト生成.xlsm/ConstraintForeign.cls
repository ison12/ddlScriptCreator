VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConstraintForeign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public constraintName As String
Public Table          As Table
Public columns        As ValCollection

Public refTable       As String
Public refColumns     As ValCollection
Public refOption      As String

