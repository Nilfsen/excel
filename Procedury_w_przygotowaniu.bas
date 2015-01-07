Attribute VB_Name = "Procedury_w_przygotowaniu"
Option Explicit

Sub export_makr()
    Dim bExport As Boolean
    Dim wbSource As Excel.Workbook
    Dim sciezka As String
    Dim nazwa As String
    Dim cmpVBE As VBIDE.VBComponent
    Dim wersja As String
    Dim c As Range

    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Procedura: export_makr"
    On Error GoTo blad
    wersja = Sheets("ustawienia").Range("B1")
' sprawdz czy istnieje folder exportu
' jeœli nie to stwórz
' stwórz folder wersji
' stwórz log zmian
' zapisz log w folderze wersji i modyfikacje historia

    Set wbSource = Application.Workbooks(ActiveWorkbook.Name)
    If wbSource.VBProject.Protection = 1 Then
    MsgBox "Edytor VBA jest chroniony." & vbCrLf & "Export makr niemo¿liwy"
    Exit Sub
    End If
    
    sciezka = ActiveWorkbook.Path & "\!archiwum\" & wersja & "\"
    
    For Each cmpVBE In wbSource.VBProject.VBComponents
        nazwa = cmpVBE.Name
        For Each c In Sheets("ustawienia").Range("A3:A42").Cells
            If nazwa = c.Value Then bExport = True
            If c = "" Then Exit For
        Next
        Select Case cmpVBE.Type
            Case vbext_ct_ClassModule
                nazwa = nazwa & ".cls"
            Case vbext_ct_MSForm
                nazwa = nazwa & ".frm"
            Case vbext_ct_StdModule
                nazwa = nazwa & ".bas"
            Case vbext_ct_Document
                bExport = False
            Case Else
                bExport = False
        End Select
        If bExport = True Then
            cmpVBE.Export sciezka & nazwa
        End If
    Next cmpVBE
    MsgBox "Zakoñczono export wersji: " & wersja
Exit Sub
blad:
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Sub
