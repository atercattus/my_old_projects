ini_check_eax proc field:DWORD
; # сравнивает eax с #0. Если равно, то выдается сообщение об ошибке и устанавливается флаг cf. Что-то типа сишной assert().  
	
	cmp		eax,	0
	jne		@F
		.data
			ini_err_begin	db 'Ошибка чтения параметра "',0
			ini_err_end		db '" из файла инициализации!',0
		.code		
		
		; формирую строку из трех частей		
		invoke	strCopy, tmp_str, addr ini_err_begin
		invoke	strCat, tmp_str, field
		invoke	strCat, tmp_str, addr ini_err_end
		
		; вывожу сообщение об ошибке
		invoke	MessageBox, WND, tmp_str, addr AppName, MB_OK or MB_ICONERROR
		
		mov		byte ptr [tmp_str], 0
		
		stc ; произошла ошибка
		ret
@@:

	clc
	ret
ini_check_eax endp

ReadINIFile proc
; # Чтение настроек из INI-файла

	.data
		section	db "Info",0
		k_OLC	db "OLC",0
		k_MLCB	db "MLCB",0
		k_MLCE	db "MLCE",0
		k_TD	db "TD",0
		k_SD	db "StartDir",0
		k_F		db "Format",0
		def_str db 0,0
	.code
		
; One Line Comment
	invoke	GetPrivateProfileString, addr section, addr k_OLC, addr def_str, addr OLC, 20, INIPath
	invoke	ini_check_eax, addr k_OLC
	jc		@F
	invoke	strExt, addr OLC
;	invoke	Delim2Zero, addr OLC ; однострочный комментарий не бывает ПОКА нескольких видов (есть ли смысл вообще???)
		
; Multi Line Comment Begin
	invoke	GetPrivateProfileString, addr section, addr k_MLCB, addr def_str, addr MLCB, 20, INIPath
	invoke	ini_check_eax, addr k_MLCB
	jc		@F
	invoke	strExt, addr OLC
	invoke	Delim2Zero, addr MLCB

; Multi Line Comment End
	invoke	GetPrivateProfileString, addr section, addr k_MLCE, addr def_str, addr MLCE, 20, INIPath
	invoke	ini_check_eax, addr k_MLCE
	jc		@F
	invoke	strExt, addr OLC
	invoke	Delim2Zero, addr MLCE
	
; Text Delimiters
	invoke	GetPrivateProfileString, addr section, addr k_TD, addr def_str, addr TD, 20, INIPath
	invoke	ini_check_eax, addr k_TD
	jc		@F
	invoke	strExt, addr OLC
	
; init directory
	invoke	GetPrivateProfileString, addr section, addr k_SD, addr def_str, StartPath, 100, INIPath
	cmp		eax,	0
	jne		@@StartDir_exists
		; начальный путь не указан. использую текущий.
		invoke	strCopy, StartPath, EXEPath
@@StartDir_exists:
	
; Format of file name
	invoke	GetPrivateProfileString, addr section, addr k_F, addr def_str, addr format, 10, INIPath
	invoke	ini_check_eax, addr k_F
	jc		@F
	
; All OK
	invoke	MessageBox, WND, SADD("Параметры загружены"), addr AppName, MB_OK or MB_ICONINFORMATION
	invoke	EnableMenuItem, hMenu, IDM_START, MF_ENABLED
	ret
		
@@:
; Error	
	invoke	MessageBox, WND, SADD("Ошибка при загрузке параметров!"), addr AppName, MB_OK or MB_ICONERROR
	invoke	EnableMenuItem, hMenu, IDM_START, MF_GRAYED
	ret

ReadINIFile endp

.data?
	INI_file_handle dd ?
.code

WriteToINIFile proc what:DWORD, what_len:DWORD
; # Запись строки what длиной len в файл инициализации.

	LOCAL	wd : DWORD

	invoke	WriteFile, INI_file_handle, what, what_len, addr wd, NULL
	
	ret
WriteToINIFile endp	

.data
	equal_ch db "=",0
	skobka1 db "[",0
	skobka2 db "]"
	CRLF db 0Dh,0Ah,0
.code

WriteField proc key, key_val : DWORD

	invoke	szLen, key
	invoke	WriteToINIFile, key, eax

	invoke	szLen, addr equal_ch
	invoke	WriteToINIFile, addr equal_ch, eax
	
	invoke	CopyZero2Delim, tmp_str, key_val
	invoke	szLen, tmp_str
	invoke	WriteToINIFile, tmp_str, eax
	
	invoke	szLen, addr CRLF
	invoke	WriteToINIFile, addr CRLF, eax

	ret
WriteField endp

WriteINIFile proc
; # Запись параметров в INI-файл

	; Открываю файл
	invoke	CreateFile, INIPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, addr file_error, addr AppName, MB_OK or MB_ICONERROR
		ret
@@:
	mov		INI_file_handle,	eax
	
; ####################### [Info] ###############################
	invoke	strCopy, tmp_str, addr skobka1
	invoke	strCat, tmp_str, addr section
	invoke	strCat, tmp_str, addr skobka2
	invoke	szLen, tmp_str
	invoke	WriteToINIFile, tmp_str, eax

	invoke	WriteField, addr k_OLC, addr OLC
	invoke	WriteField, addr k_MLCB, addr MLCB
	invoke	WriteField, addr k_MLCE, addr MLCE
	invoke	WriteField, addr k_TD, addr TD
	invoke	WriteField, addr k_SD, StartPath
	invoke	WriteField, addr k_F, addr format
	
	; Закрываю файл
	invoke	CloseHandle, INI_file_handle

	ret
WriteINIFile endp