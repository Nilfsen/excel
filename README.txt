

Lista modyfikacji:

v1000
	1		Wprowadzenie pliku z zestawem makr
	2		Funkcja txt_new - tworzy nowy plik tekstowy
	3		Funkcja txt_add - dodaje do wskazanego pliku now¹ treœæ
	4		Funkcja nowy_folder - tworzy nowy folder
	5		Funkcja istnieje_plik - sprawdza czy dany plik istnieje
	6		Funkcja istnieje_folder - sprawdza czy dany folder istnieje
	7		Funkcja log_to_txt - tworzy plik logu lub dopisuje zdarzenie do istniej¹cego pliku
	8		Procedura Export_makr - exportuje pliki .bas do folderu !archiwum oraz uaktualnia pliki w folderze github


v1001
	1		Nie zapisuje kopi pliku makra.xls w katalogu github, tylko wyeksportowane makra
	2		dodana funkcja zwick_set_param
	3		Dodana publiczna zmienna ilosc_blad do sledzenia ilosci b³êdów w czasie wykonywania procedur
	4		procedura export_makr bêdzie tworzyæ log dodawany do folderu /!archiwum/bie¿¹ca_wersja/
	5		Dodana funkcja kopiuj_plik - kopiuje plik z jednego folderu do drugiego, opcjonalnie zmienia nazwê kopii
	6		procedura export_makr bêdzie tworzyæ log dodawany do folderu github ( usuwa starszy )
	7		nowa funkcja kopiuj_folder
v1002
	1		poprawione tworzenie plików log procedury konwerter_makr ( tworzenie nazwy pliku )
v1003
	1		Ponownie nazwy pliku log…
v1004
	1		i tym razem ostatnie poprawki ( oby ) dla nazw pliku log
v1005
	1		Usuniêty b³¹d gdy u¿ycie funkcji z prawid³owym wynikiem w innym makrze które wygenerowa³o b³¹d powodowa³ wskazanie b³êdnego miejsca b³êdu
	2		Uaktualniony szablon Funkcji
v1006
	1		Usuniêty b³¹d niekopuj¹cego siê logu do folderu github/excel
