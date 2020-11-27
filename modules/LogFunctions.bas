Attribute VB_Name = "LogFunctions"
'---------------------------------------------------
'Source: https://github.com/InventorCode/InventorVBA
'Author: nannerdw
'Last Modified Date: 27 Nov, 2020
'---------------------------------------------------
Option Explicit

Public Enum LogLevel
    kTrace = 1
    kDebug = 2
    kInfo = 3
    kWarn = 4
    kError = 5
    kFatal = 6
End Enum

Public Sub Log(ByVal message As String, Optional level As LogLevel = LogLevel.kInfo)
'Prints a message to the iLogic logger
    Dim iLogicAddin As ApplicationAddIn
    Set iLogicAddin = ThisApplication.ApplicationAddIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
    iLogicAddin.Automation.LogControl.Log level, message
End Sub

Public Sub LogWindow_Show()
'Forces the iLogic log window to be shown
    ThisApplication.UserInterfaceManager.DockableWindows("ilogic.logwindow").Visible = True
End Sub
