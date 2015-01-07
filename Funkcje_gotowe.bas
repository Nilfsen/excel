Attribute VB_Name = "Funkcje_gotowe"
' Zestaw  makr potrzebnych do obsługi pliku
' Autor: Witold Charewicz ( witia1@o2.pl )
'

Option Explicit

Function txt_new(nazwa As String, sciezka As String, Optional tresc As String) As Boolean
    Dim plik, objFSO As Object
    
    gdzie_blad = "Funkcja: txt_new"
    On Error GoTo blad
    If istnieje_folder(sciezka) = False Then
        opisz_blad = "Wskazany folder nie istnieje. " & sciezka
        GoTo blad
    End If
    If Right(LCase(nazwa), 4) <> ".txt" Then nazwa = nazwa & ".txt"
    If Right(sciezka, 1) <> "\" Then sciezka = sciezka & "\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set plik = objFSO.CreateTextFile(sciezka & nazwa, True)
    plik.WriteLine tresc
    plik.Close
    txt_new = True
    Exit Function
blad:
    txt_new = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Function

Function txt_add(nazwa As String, sciezka As String, tresc As String) As Boolean
    Dim plik, objFSO As Object
    
    On Error GoTo blad
    If istnieje_folder(sciezka) = False Then
        opisz_blad = "Wskazany folder nie istnieje. " & sciezka
        GoTo blad
    End If
    If Right(sciezka, 1) <> "\" Then sciezka = sciezka & "\"
    If Right(LCase(nazwa), 4) <> ".txt" Then nazwa = nazwa & ".txt"
    If istnieje_plik(nazwa, sciezka) = False Then
        opisz_blad = "Wskazany plik nie istnieje. " & sciezka & nazwa
        GoTo blad
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set plik = objFSO.OpenTextFile(sciezka & nazwa, 8)
    plik.WriteLine tresc
    plik.Close
    txt_add = True
    Exit Function
blad:
    txt_add = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Function

Function istnieje_plik(nazwa As String, sciezka As String) As Boolean
    Dim objFSO As Object
    
    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Funkcja: log_to_txt"
    On Error GoTo blad
    If Right(sciezka, 1) <> "\" Then sciezka = sciezka & "\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    istnieje_plik = objFSO.fileExists(sciezka & nazwa)
    Exit Function
blad:
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Function

Function istnieje_folder(sciezka As String) As Boolean
    Dim objFSO As Object
    
    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Funkcja: istnieje_folder"
    On Error GoTo blad
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    istnieje_folder = objFSO.folderExists(sciezka)
    Exit Function
blad:
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Function

Function nowy_folder(nazwa As String, sciezka As String) As Boolean
    Dim objFSO, folder As Object
    
    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Funkcja: nowy_folder"
    On Error GoTo blad
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Right(sciezka, 1) <> "\" Then sciezka = sciezka & "\"
    If istnieje_folder(sciezka & nazwa) Then
        nowy_folder = True
        Exit Function
    End If
        Set folder = objFSO.CreateFolder(sciezka & nazwa)
        nowy_folder = True
    Exit Function
blad:
    nowy_folder = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Function

Function log_to_txt(tresc As String, nazwa As String, sciezka As String) As Boolean

    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Funkcja: log_to_txt"
    On Error GoTo blad
    If istnieje_folder(sciezka) = False Then
        opisz_blad = "Wskazany folder nie istnieje. " & sciezka
        GoTo blad
    End If
    If Right(sciezka, 1) <> "\" Then sciezka = sciezka & "\"
    If Right(LCase(nazwa), 4) <> ".txt" Then nazwa = nazwa & ".txt"
    tresc = Format(Now(), "yyyy-mm-dd, hh:mm:ss") & vbTab & tresc
    If istnieje_plik(nazwa, sciezka) = False Then
        Call txt_new(nazwa, sciezka, tresc)
    Else
        Call txt_add(nazwa, sciezka, tresc)
    End If
    log_to_txt = True
    Exit Function
blad:
    log_to_txt = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.Path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
End Function

    
