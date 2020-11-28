Attribute VB_Name = "UtilFunctions"
'---------------------------------------------------
'Source: https://github.com/InventorCode/InventorVBA
'Author: nannerdw
'Last Modified Date: 28 Nov, 2020
'---------------------------------------------------
Option Explicit

Private Sub ViewLocals()
    'This sub is for browsing through the local variables related to a selected object.
    'Select an object in Inventor, run this sub, then view the locals window
    Dim app As Application: Set app = ThisApplication
    Dim doc As Document: Set doc = app.ActiveDocument
    Dim selSet As SelectSet: Set selSet = doc.SelectSet
    
    If selSet.Count > 0 Then
        Dim selectedObj As Object: Set selectedObj = selSet(1)
        
        'Uncomment this only if you know the selected object has a Browser Node.  Otherwise Inventor can crash.
        'Dim selectedObj_BrowserNode As BrowserNode: Set selectedObj_BrowserNode = doc.BrowserPanes("Model").GetBrowserNodeFromObject(selectedObj)
    End If
    
    Stop
End Sub
