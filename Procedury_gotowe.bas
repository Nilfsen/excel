Attribute VB_Name = "Procedury_gotowe"
' Zestaw  makr potrzebnych do obs³ugi pliku
' Autor: Witold Charewicz ( witia1@o2.pl )
'
Sub export_makr()
    Dim bExport As Boolean
    Dim wbSource As Excel.Workbook
    Dim path As String
    Dim path2 As String
    Dim nazwa As String
    Dim cmpVBE As VBIDE.VBComponent
    Dim wersja As String
    Dim c As Range
    Dim log As String
    Dim get_log As Variant
    Dim x As Integer
    
    ilosc_blad = 0
    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Procedura: export_makr"
    On Error GoTo blad
    wersja = "v" & Sheets("ustawienia").Range("B1")
    Set wbSource = Application.Workbooks(ActiveWorkbook.Name)
        path2 = "D:\!Github\excel\"
    path = ActiveWorkbook.path & "\!archiwum\" & wersja
    Call log_to_txt("Rozpoczêto export wersji: " & wersja, "v" & wersja & "_konwerter_log", path)
    Call nowy_folder(wersja, ActiveWorkbook.path & "\!archiwum\")
    Call log_to_txt("Dodano nowy folder: " & ActiveWorkbook.path & "\!archiwum\" & wersja, "v" & wersja & "_konwerter_log", path)
    If wbSource.VBProject.Protection = 1 Then
        Call log_to_txt("Edytor VBA jest chroniony. Export makr niemo¿liwy. Przerwano.", "v" & wersja & "_konwerter_log", path)
        Exit Sub
    End If
    On Error Resume Next
        Kill path2 & "*.bas"
        Kill path2 & "*.xls"
        Call log_to_txt("Usuniête pliki .bas i .xls z folderu: " & path2, "v" & wersja & "_konwerter_log", path)
    On Error GoTo 0
    
    For Each cmpVBE In wbSource.VBProject.VBComponents
        nazwa = cmpVBE.Name
        bExport = False
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
            cmpVBE.Export path & "\" & nazwa
            Call log_to_txt("Wyeksportowano: " & nazwa & " do:  " & path, "v" & wersja & "_konwerter_log", path)
            If istnieje_folder(path2) = True Then
                cmpVBE.Export path2 & nazwa
                Call log_to_txt("Wyeksportowano: " & nazwa & " do:  " & path2, "v" & wersja & "_konwerter_log", path)
            End If
        End If
    Next cmpVBE
    Sheets("ustawienia").Range("B1") = Sheets("ustawienia").Range("B1") + 1
    get_log = Sheets("ustawienia").Range("D3:E52")
    log = wersja
    For x = 1 To 50
        If get_log(x, 2) <> "" Then
            log = log & vbCrLf & vbTab & get_log(x, 1) & vbTab & vbTab & get_log(x, 2)
        End If
    Next x
    Sheets("ustawienia").Range("E3:Q52").ClearContents
    Call txt_new("lista_zmian.txt", path, log)
    Call log_to_txt("Stworzony plik lista_zmian.txt w katalogu:  " & path, "v" & wersja & "_konwerter_log", path)
    If istnieje_folder(path2) = True Then
        Call txt_add("readme.txt", path2, log)
        Call log_to_txt("Zaktualizowana lista zmian readme.txt dla Github", "v" & wersja & "_konwerter_log", path)
    End If
    ActiveWorkbook.Save
    Call log_to_txt("Zakoñczono export wersji: " & wersja, "v" & wersja & "_konwerter_log", path)
    Call kopiuj_plik("v" & wersja & "_konwerter_log", path, path2, , True)
zakoncz:
If ilosc_blad > 0 Then
' show log b³êdów
End If
' show log konwersji
Exit Sub
blad:
    Call log_to_txt("Przy konwersji wyst¹pi³y b³êdy. Przerwano.", "v" & wersja & "_konwerter_log", path)
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Sub



