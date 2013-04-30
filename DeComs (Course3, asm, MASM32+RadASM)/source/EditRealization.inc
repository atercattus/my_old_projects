.data

	; размер текста в окне
	old_file_size	dd	0
	; идентификатор текст
	buff			dd	0
	
.data?
	hEdit	HWND	?
	
.const
	EditStyle equ	WS_CHILD or WS_VISIBLE or WS_HSCROLL or WS_VSCROLL or ES_NOHIDESEL or ES_SAVESEL or \
					ES_LEFT or ES_MULTILINE or ES_AUTOHSCROLL or ES_AUTOVSCROLL or ES_SUNKEN
	
.code

Clear proc
; # Очистка содержимого редактора.
	
	invoke	SendMessage, hEdit, EM_SETHANDLE, NULL, 0
	invoke	StatusBarSetText, addr sb_def_text
	
	ret
Clear endp

Edt_Init proc
; # Создание и инициализация редактора.

	invoke	CreateWindowEx, 0, SADD("edit"), NULL, EditStyle, 0, 0, 0, 0, WND, NULL, hInstance, NULL

	mov		hEdit, eax
	
	; запрещаю ввод в редактор
	invoke	SendMessage, hEdit, EM_GETHANDLE, NULL, 0
	invoke	LocalFree, eax
	invoke	SendMessage, hEdit, EM_SETHANDLE, 0, 0

	invoke	StatusBarSetText, addr sb_def_text
	
	mov		buff,	0
	mov		old_file_size,	0
	
	ret
Edt_Init endp

Edt_Done proc
; # Освобождение памяти, выделенной под редактор.

	invoke	LocalFree, buff
	
	invoke	DestroyWindow, hEdit
	
	ret	
Edt_Done endp

DelRange proc From:DWORD,To:DWORD
; # Удаление куска текста из памяти, заданного промежутком [From, To). Диапазон проверяется.

	push	eax
	push	ecx
	push	edx

	; проверка диапазона >>
	mov		eax,	From
	mov		ecx,	To
	mov		edx,	old_file_size
	
	cmp		eax,	edx
	jna		@F
		jmp		@@exit	; From >= Size
@@:

	cmp		ecx,	edx
	jna		@F
		jmp		@@exit	; To >= Size
@@:

	cmp		eax,	ecx
	jb		@F
		jmp		@@exit	; From >= To
@@:
	
	; удаление >>
	
	push	esi
	push	edi

	mov		eax,	buff_ptr
	
	; вычисляю точки входа
	mov		esi,	eax
	mov		edi,	eax
	
	add		edi,	From
	add		esi,	To
	
	; вычисляю длину копируемого блока
	mov		ecx,	old_file_size
	sub		ecx,	To
	inc		ecx	; #0 тоже копирую

	; копирую
	push	ecx
	invoke	RtlMoveMemory, edi, esi, ecx
	pop		ecx
	dec		ecx ; снова забиваю на #0
	
	; пересчитываю позиции
	add		esi,	ecx
	add		edi,	ecx
	
	; пересчет размера буфера >>
	mov		eax,	To
	sub		eax,	From
	sub		old_file_size,	eax
	
	pop		edi
	pop		esi
	
@@exit:
	pop		edx
	pop		ecx
	pop		eax
	
	ret
DelRange endp

.const
	file_error db "Ошибка при работе с файлом!",0
	mem_error db "Ошибка работы с памятью",0
.code

LF2CRLF proto
CheckFileType proto

LoadFromFile proc Path:DWORD
; # Загрузка файла.

	LOCAL	file : DWORD
	LOCAL	file_size : DWORD
	LOCAL	rd : DWORD
	LOCAL	pntr : DWORD

	mov		eax,	buff
	cmp		eax,	0
	je		@F
		invoke	LocalFree, eax
@@:

	; Открываю файл
	invoke	CreateFile, Path, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, addr file_error, addr AppName, MB_OK or MB_ICONERROR
		ret
@@:
	mov		file,	eax

	; Читаю файл в память
	invoke	GetFileSize, eax, NULL
	mov		file_size,	eax
	inc		eax	; в конце идет 0
	
	; ИМЕННО ПЕРЕМЕЩАЕМУЮ, иначе "edit" не хочет толком работать! ... в теории ;)
	invoke	LocalAlloc, LMEM_MOVEABLE, eax
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, addr mem_error, addr AppName, MB_OK or MB_ICONERROR
		invoke	CloseHandle, file
		ret	
@@:
	mov		buff,	eax
	
	invoke	LocalLock, buff
	mov		pntr,	eax
	
	invoke	ReadFile, file, pntr, file_size, addr rd, NULL
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, addr file_error, addr AppName, MB_OK or MB_ICONERROR
		invoke	CloseHandle, file
		ret	
@@:

	; Закрываю файл
	invoke	CloseHandle, file
	
	; В конце содержимого файла должен стоять 0
	mov		eax,	pntr
	add		eax,	file_size
	mov		byte ptr [eax],	0

	invoke	LocalUnlock, buff

	m2m		old_file_size, file_size
	
	; проверка типа файла
	invoke	CheckFileType
	
	invoke	SendMessage, hEdit, EM_SETHANDLE, buff, 0

	invoke	StatusBarSetText, Path
	
	ret
LoadFromFile endp

SaveToFile proc Path:DWORD
; # Сохранение файла.

	LOCAL	file : DWORD
	LOCAL	wd : DWORD
	LOCAL	pntr : DWORD
	
	mov		eax,	buff
	cmp		eax,	0
	jne		@F
		ret ; нечего сохранять
@@:

	; Открываю файл
	invoke	CreateFile, Path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, addr file_error, addr AppName, MB_OK or MB_ICONERROR
		ret
@@:
	mov		file,	eax

	; временно посылаю менеджер памяти подальше
	invoke	LocalLock, buff
	mov		pntr,	eax
	
	; скидываю буфер в файл
	invoke	WriteFile, file, pntr, old_file_size, addr wd, NULL
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, addr file_error, addr AppName, MB_OK or MB_ICONERROR
		invoke	LocalUnlock, buff ; а зачем, если сейчас сносить ?
		invoke	CloseHandle, file
		ret	
@@:

	; Закрываю файл
	invoke	CloseHandle, file
	
	invoke	LocalUnlock, buff ; а зачем ???

	ret
SaveToFile endp

DelConfirm proc From:DWORD,To:DWORD
; # Диалог на удаление комментария (текста, см. DelRange выше).

	invoke	SendMessage, hEdit, EM_SETHANDLE, buff, 0
	
	; выделяю удаляемый текст
	invoke	SendMessage, hEdit, EM_SETSEL, From, To
	; если текст за пределами окна Edit'а, то прокручиваю до него
	invoke	SendMessage, hEdit, EM_SCROLLCARET, 0, 0
	
	invoke	MessageBox, WND, SADD("Удалить данный комментарий?"), addr AppName, MB_YESNOCANCEL or MB_ICONEXCLAMATION
	cmp		eax,	IDYES
	jne		@F
		; идиот согласился вырезать кусок из файла... гы, гы...
		invoke	DelRange, From, To
		mov		eax,	IDYES
@@:
	
	ret
DelConfirm endp

CheckFileType proc
; # Проверяет принадлежность файла к типам "CR/LF" или "LF". В последнем случае выполняет преобразование.

	invoke	LocalLock, buff

@@loop:
	mov		cl,		byte ptr [eax]
	cmp		cl,		0
	je		@@exit
	
	cmp		cl,		0Dh ; CR
	je		@@exit	; нашел CR/LF - преобразование не нужно
					; фиг с ней, с проверкой на наличие LF после CR. Без этого это уже не исходник, а бинарник.
	
	cmp		cl,		0Ah ; LF
	jne		@F
		; нашел LF - необходимо преобразование
		invoke	LocalUnlock, buff
		invoke	LF2CRLF
		ret
@@:

	inc		eax
	jmp		@@loop	
	
@@exit:

	invoke	LocalUnlock, buff
	
	ret
CheckFileType endp

LF2CRLF proc
; # Преобразование однобайтной последовательности "LF" в двухбайтовую "CR LF".
	LOCAL	pntr : DWORD
	LOCAL	new_pntr : DWORD
	LOCAL	new_buff : DWORD
	
	invoke	LocalLock, buff
	mov		pntr,	eax
	
	; считаю количество разделителей ("LF", 0Ah)
	xor		ecx,	ecx ; счетчик
	
@@count_loop:
	mov		dl,		byte ptr [eax]
	cmp		dl,		0
	je		@@count_exit
	
	cmp		dl,		0Ah
	jne		@F
		inc		ecx
@@:
	inc		eax
	jmp		@@count_loop

@@count_exit:

	; Теперь в ecx число байтов LF
	; Новый необходимый размер буфера будет равен Size = Old_Size + ecx {+CR по кол-ву LF}
	mov		eax,	old_file_size
	add		eax,	ecx
	mov		old_file_size,	eax
	inc		eax ; + конечный #0
	
	; создаю новый
	invoke	LocalAlloc, LMEM_MOVEABLE, eax
	mov		new_buff,	eax
	
	invoke	LocalLock, eax
	mov		new_pntr,	eax
	
	push	esi
	push	edi
	
	mov		esi,	pntr
	mov		edi,	new_pntr
	
	; копирую с заменой	
@@replace_loop:
	mov		al,		byte ptr [esi]
	cmp		al,		0
	je		@@replace_exit
	
	.IF (al==0Ah)
		; нашел разделитель - сразу пишу слово
		mov		word ptr [edi],	0A0Dh
		inc		edi ; т.к. записал слово, а не байт
	.ELSE
		mov		byte ptr [edi],	al
	.ENDIF
	
	inc		esi
	inc		edi
	
	jmp		@@replace_loop

@@replace_exit:

	mov		byte ptr [edi],	0

	pop		edi
	pop		esi
	
	; удаляю старый буфер и сохраняю новый
	invoke	LocalUnlock, buff
	invoke	LocalUnlock, new_buff
	
	invoke	LocalFree, buff
	m2m		buff,	new_buff
	
	ret
LF2CRLF endp