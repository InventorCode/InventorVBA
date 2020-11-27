Attribute VB_Name = "iLogicFunctions"
'---------------------------------------------------
'Source: https://github.com/InventorCode/InventorVBA
'Author: nannerdw
'Last Modified Date: 27 Nov, 2020
'---------------------------------------------------
Option Explicit

'***************************************************
Public Function GetiLogicAddin() As ApplicationAddIn
'Description: Returns the iLogic addin
'Dependencies: N/A
'***************************************************
    Const FUNCTION_NAME As String = "GetiLogicAddin"
    Const ILOGIC_ADDIN_GUID As String = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}"
    
    Dim addIn As ApplicationAddIn
    
    On Error Resume Next
    Set addIn = ThisApplication.ApplicationAddIns.ItemById(ILOGIC_ADDIN_GUID)
    
    Select Case Err.Number
    Case 0
        Set GetiLogicAddin = addIn
    Case -2147467259
        MsgBox _
            Title:=FUNCTION_NAME, _
            Prompt:="iLogic Addin could not be found.", _
            Buttons:=vbExclamation
    Case Else
        Err.Raise Err.Number
    End Select
End Function

'************************
Public Sub RuniLogicRule( _
    ByVal RuleName As String, _
    Optional ruleArgs As NameValueMap = Nothing, _
    Optional doc As Document = Nothing)
'Description: Runs an iLogic rule that is stored inside doc
'
'Dependencies:
'   Function GetiLogicAddin
'
'Notes:
'   ThisDoc.Document inside an iLogic rule will refer to the "doc" variable that was passed in to this sub.
'   If doc is not passed into this sub, it defaults to ThisApplication.ActiveDocument.
'************************
    Const SUB_NAME As String = "RuniLogicRule"
    
    If doc Is Nothing Then Set doc = ThisApplication.ActiveDocument
    
    If doc Is Nothing Then
        MsgBox _
            Title:=SUB_NAME, _
            Prompt:="A document must be active.", _
            Buttons:=vbExclamation
        Exit Sub
    End If
    
    Dim addIn As ApplicationAddIn: Set addIn = GetiLogicAddin
    If addIn Is Nothing Then Exit Sub
    
    Dim iLogicAuto As Object: Set iLogicAuto = addIn.Automation
    
    If ruleArgs Is Nothing Then
        iLogicAuto.RunRule doc, RuleName
    Else
        iLogicAuto.RunRuleWithArguments doc, RuleName, ruleArgs
    End If
End Sub

'********************************
Public Sub RunExternaliLogicRule( _
    ByVal RuleName As String, _
    Optional ruleArgs As NameValueMap = Nothing, _
    Optional doc As Document = Nothing)
'Description: Runs an external iLogic Rule.
'
'Dependencies:
'   Function GetiLogicAddin
'
'Notes:
'   ThisDoc.Document inside an iLogic rule will refer to the "doc" variable that was passed in to this sub.
'   If doc is not passed into this sub, it defaults to ThisApplication.ActiveDocument.
'
'   RuleName can be one of the following:
'   1. A fully-qualified path to an external rule (with or without a file extension),
'
'   2. A relative file path (with or without a file extension),
'      starting from one of the iLogic external rule directories,
'      but excluding the top-level external rule folder name itself.
'      If you choose this option, make sure you don't have the same
'      relative file path in multiple external rule directories.
'
'   3. A filename only (with or without a file extension).
'      If you choose this option, make sure you don't have the same filename in multiple folders.
'********************************
    Const SUB_NAME As String = "RunExternaliLogicRule"
    
    If doc Is Nothing Then Set doc = ThisApplication.ActiveDocument
    
    If doc Is Nothing Then
        MsgBox _
            Title:=SUB_NAME, _
            Prompt:="A document must be active.", _
            Buttons:=vbExclamation
        Exit Sub
    End If
    
    Dim addIn As ApplicationAddIn: Set addIn = GetiLogicAddin
    If addIn Is Nothing Then Exit Sub
    
    Dim iLogicAuto As Object: Set iLogicAuto = addIn.Automation
    
    If ruleArgs Is Nothing Then
        iLogicAuto.RunExternalRule doc, RuleName
    Else
        iLogicAuto.RunExternalRuleWithArguments doc, RuleName, ruleArgs
    End If
End Sub
