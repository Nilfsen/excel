Attribute VB_Name = "Funkcje_gotowe"
' Zestaw  makr potrzebnych do obs³ugi pliku
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
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    txt_new = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
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
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    txt_add = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Function

Function istnieje_plik(nazwa As String, sciezka As String) As Boolean
    Dim objFSO As Object
    
    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Funkcja: istnieje_plik"
    On Error GoTo blad
    If Right(sciezka, 1) <> "\" Then sciezka = sciezka & "\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    istnieje_plik = objFSO.fileExists(sciezka & nazwa)
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Function

Function istnieje_folder(sciezka As String) As Boolean
    Dim objFSO As Object
    
    old_gdzie_blad = gdzie_blad
    gdzie_blad = "Funkcja: istnieje_folder"
    On Error GoTo blad
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    istnieje_folder = objFSO.folderExists(sciezka)
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
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
        gdzie_blad = old_gdzie_blad
        Exit Function
    End If
        Set folder = objFSO.CreateFolder(sciezka & nazwa)
        nowy_folder = True
        gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    nowy_folder = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
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
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    log_to_txt = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Function

Function zwick_set_param(numer_parametru As Integer, rodzaj As String, wartosc As Variant, Optional jednostka As String, Optional czy_tablica As Boolean, Optional index_tablicy As Integer, Optional komentarz_skryptu As String) As String
    gdzie_blad = "Funkcja: zwick_set_param"
    On Error GoTo blad
    If numer_parametru = "0" Or numer_parametru = "" Then
        opisz_blad = "Wartoœæ numer_parametru musi byæ numeryczna i wiêksza od 0. Podano: " & numer_parametru
        GoTo blad
    End If
    Select Case LCase(rodzaj)
        Case Is = "boolean"
            Select Case LCase(wartosc)
                Case Is = "true"
                    wartosc = "True"
                Case Is = "tak"
                    wartosc = "True"
                Case Is = "false"
                    wartosc = "False"
                Case Is = "nie"
                    wartosc = "False"
                Case Else
                    opisz_blad = "Nieprawid³owa wartoœæ dla parametru typu Boolean. Parametr: " & numer_parametru & vbTab & wartosc
                    GoTo blad
            End Select
            zwick_set_param = "SetParam " & numer_parametru & " , " & wartosc
        Case Is = "num"
            If wartosc = "" Then wartosc = 0
            If IsNumeric(wartosc) = False Then
                opisz_blad = "Nieprawid³owa wartoœæ dla parametru typu numerycznego.Parametr: " & numer_parametru & vbTab & wartosc
                GoTo blad
            End If
            If czy_tablica = False Or czy_tablica = "" Then
                zwick_set_param = "SetParam " & numer_parametru & " , " & wartosc
                If jednostka <> "" Then zwick_set_param = zwick_set_param & " , " & """" & jednostka & """"
            ElseIf czy_tablica = True Then
                zwick_set_param = "SetArray " & numer_parametru & " , " & index_tablicy & " , " & wartosc
                If index_tablicy = 1 And jednostka <> "" Then zwick_set_param = zwick_set_param & " ," & """" & jednostka & """"
            End If
        Case Is = "text"
            zwick_set_param = "t[ " & numer_parametru & "] = " & """" & wartosc & """"
        Case Else
            opisz_blad = "Nieznany typ parametru: " & rodzaj
            GoTo blad
    End Select
    If komentarz_skryptu <> "" Then zwick_set_param = zwick_set_param & " ; " & komentarz_skryptu
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    zwick_set_param = "b³¹d"
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Function

Function kopiuj_plik(nazwa_pliku As String, kopiuj_z As String, kopiuj_do As String, Optional nowa_nazwa As String, Optional czy_nadpisac_istniejacy As Boolean) As Boolean
    Dim objFSO As Object
    
    gdzie_blad = "Funkcja: kopiuj_plik"
    On Error GoTo blad
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If nazwa_pliku = "" Then
        opisz_blad = "Wartoœæ nazwa_pliku nie mo¿e byæ pusta"
        GoTo blad
    End If
    If kopiuj_z = "" Then
        opisz_blad = "Wartoœæ kopiuj_z nie mo¿e byæ pusta"
        GoTo blad
    End If
    If kopiuj_do = "" Then
        opisz_blad = "Wartoœæ kopiuj_do nie mo¿e byæ pusta"
        GoTo blad
    End If
    If istnieje_folder(kopiuj_z) = False Then
        opisz_blad = "Brak wskazanego folderu Ÿród³owego: " & kopiuj_z
        GoTo blad
    End If
    If istnieje_plik(nazwa_pliku, kopiuj_z) = False Then
        opisz_blad = "Brak wskazanego pliku Ÿród³owego: " & nazwa_pliku
        GoTo blad
    End If
    If istnieje_folder(kopiuj_do) = False Then
        opisz_blad = "Brak wskazanego folderu Ÿród³owego: " & kopiuj_do
        GoTo blad
    End If
    If Right(kopiuj_z, 1) <> "\" Then kopiuj_z = kopiuj_z & "\"
    If Right(kopiuj_do, 1) <> "\" Then kopiuj_do = kopiuj_do & "\"
    If nowa_nazwa = "" Then nowa_nazwa = nazwa_pliku
    objFSO.CopyFile kopiuj_z & nazwa_pliku, kopiuj_do & nowa_nazwa, czy_nadpisac_istniejacy
    kopiuj_plik = True
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    kopiuj_plik = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Function

Function kopiuj_folder(folder_zrodlowy As String, kopiuj_do As String, Optional czy_nadpisac_istniejacy As Boolean) As Boolean
    Dim objFSO As Object
    
    gdzie_blad = "Funkcja: kopiuj_folder"
    On Error GoTo blad
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If folder_zrodlowy = "" Then
        opisz_blad = "Wartoœæ folder_zrodlowy nie mo¿e byæ pusta"
        GoTo blad
    End If
    If kopiuj_do = "" Then
        opisz_blad = "Wartoœæ kopiuj_do nie mo¿e byæ pusta"
        GoTo blad
    End If
    If istnieje_folder(folder_zrodlowy) = False Then
        opisz_blad = "Brak wskazanego folderu Ÿród³owego: " & folder_zrodlowy
        GoTo blad
    End If
    objFSO.CopyFolder folder_zrodlowy, kopiuj_do, czy_nadpisac_istniejacy
    kopiuj_folder = True
    gdzie_blad = old_gdzie_blad
    Exit Function
blad:
    kopiuj_folder = False
    If opisz_blad = "" Then opisz_blad = "Niezdefiniowany. " & Err & ": " & Err.Description
    If log_to_txt(gdzie_blad & vbTab & opisz_blad, "log_" & ActiveWorkbook.Name, ActiveWorkbook.path) = False Then MsgBox gdzie_blad & vbTab & opisz_blad, vbOK, "B³¹d " & ActiveWorkbook.Name
    gdzie_blad = old_gdzie_blad
    ilosc_blad = ilosc_blad + 1
End Function

    
