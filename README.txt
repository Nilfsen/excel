

Lista modyfikacji:

v1000
	1		Wprowadzenie pliku z zestawem makr
	2		Funkcja txt_new - tworzy nowy plik tekstowy
	3		Funkcja txt_add - dodaje do wskazanego pliku now� tre��
	4		Funkcja nowy_folder - tworzy nowy folder
	5		Funkcja istnieje_plik - sprawdza czy dany plik istnieje
	6		Funkcja istnieje_folder - sprawdza czy dany folder istnieje
	7		Funkcja log_to_txt - tworzy plik logu lub dopisuje zdarzenie do istniej�cego pliku
	8		Procedura Export_makr - exportuje pliki .bas do folderu !archiwum oraz uaktualnia pliki w folderze github


v1001
	1		Nie zapisuje kopi pliku makra.xls w katalogu github, tylko wyeksportowane makra
	2		dodana funkcja zwick_set_param
	3		Dodana publiczna zmienna ilosc_blad do sledzenia ilosci b��d�w w czasie wykonywania procedur
	4		procedura export_makr b�dzie tworzy� log dodawany do folderu /!archiwum/bie��ca_wersja/
	5		Dodana funkcja kopiuj_plik - kopiuje plik z jednego folderu do drugiego, opcjonalnie zmienia nazw� kopii
	6		procedura export_makr b�dzie tworzy� log dodawany do folderu github ( usuwa starszy )
	7		nowa funkcja kopiuj_folder
v1002
	1		poprawione tworzenie plik�w log procedury konwerter_makr ( tworzenie nazwy pliku )
v1003
	1		Ponownie nazwy pliku log�
v1004
	1		i tym razem ostatnie poprawki ( oby ) dla nazw pliku log
v1005
	1		Usuni�ty b��d gdy u�ycie funkcji z prawid�owym wynikiem w innym makrze kt�re wygenerowa�o b��d powodowa� wskazanie b��dnego miejsca b��du
	2		Uaktualniony szablon Funkcji
v1006
	1		Usuni�ty b��d niekopuj�cego si� logu do folderu github/excel
