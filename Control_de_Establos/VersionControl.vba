Attribute VB_Name = "VersionControl"
Option Explicit
' Modulo tomado de https://code.i-harness.com/es/q/20215

Private Sub SaveCodeMod()
'Private Sub SaveCodeModules()
'This code Exports all VBA modules
    Dim i%, sName$
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            Application.StatusBar = _
              "Exportando Código... " & _
              Format(i% / .VBComponents.Count, "0%")
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                '.VBComponents(i%).Export _
                  "X:\Tools\MyExcelMacros\" & sName$ & ".vba"
                .VBComponents(i%).Export _
                  "C:\Users\jpher\OneDrive\Desarrollos\Control_de_Establos\" _
                  & sName$ & ".vba"
            End If
        Next i
        Application.StatusBar = False
    End With
End Sub

Private Sub ImportCodeMod()
' This code Imports all VBA modules
    Dim i%, ModuleName
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            ModuleName = .VBComponents(i%).CodeModule.Name
            If ModuleName <> "VersionControl" Then
                'If Right(ModuleName, 6) = "Macros" Then
                    .VBComponents.Remove .VBComponents(ModuleName)
                    .VBComponents.Import _
                      "C:\Users\jpher\OneDrive\Desarrollos\Control_de_Establos\" _
                      & ModuleName & ".vba"
                    '.VBComponents.Import _
                       "X:\Data\MySheet\" & ModuleName & ".vba"
               'End If
            End If
        Next i
    End With
End Sub

' Este código debe ir in ThisWorkbook
Private Sub Workbook_Open()
    ImportCodeModules
End Sub

' Este código debe ir in ThisWorkbook
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    SaveCodeModules
End Sub
