Attribute VB_Name = "Modulo4"
Option Explicit

Private Sub SaveCodeModules()
    'This code Exports all VBA modules
    Dim i%, sName$
    
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                .VBComponents(i%).Export "X:\Tools\MyExcelMacros\" & sName$ & ".vba"
            End If
        Next i
    End With
End Sub

Sub ImportCodeModules()
    'This code Imports all VBA modules
    Dim i%, ModuleName
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            ModuleName = .VBComponents(i%).CodeModule.Name
            If ModuleName <> "VersionControl" Then
                If Right(ModuleName, 6) = "Macros" Then
                    .VBComponents.Remove .VBComponents(ModuleName)
                    .VBComponents.Import "X:\Data\MySheet\" & ModuleName & ".vba"
                End If
            End If
        Next i
    End With
End Sub

' ThisWorkbook
Private Sub Workbook_Open()
    ImportCodeModules
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    SaveCodeModules
End Sub
