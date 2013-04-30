.data?
	; шаблон имени файла
	FileFormat	dd	?
	; формат имени файла для поиска
	format		db 20 dup (?)
	; фиксированный буфер файла
	buff_ptr	dd	?
	; номер найденного шаблона в списке. если не найден, то НЕ МЕНЯЕТСЯ.
	templ_num	dw	? ; можно и байт, на не работает тогда стек
	
	; списки шаблонов
	OLC			db	20 dup (?)
	MLCB		db	20 dup (?)
	MLCE		db	20 dup (?)
	TD			db	20 dup (?)
	
.data
	EOL			db	0Dh,0Ah,0,0
	
.code

equal proc templ:DWORD
; # Поиск вхождения набора шаблонов [templ] в буфер с позиции [esi].
;
;	Вход:
;		templ - указатель на список шаблонов, разделенных #0. в конце еще один #0.
;		esi - указатель на буфер файла.
;	Выход:
;		eax:
;			-1	: дошел до конца буфера файла
;			#0	: дошел до конца списка шаблонов
;			any	: адрес в буфере с которого найден шаблон
;		ecx - длина найденного шаблона
;		edi - указывает на #0 за найденным шаблоном
;		esi - указывает на символ ЗА концом найденного шаблона в буфере файла
;		templ_num - номер найденного шаблона в списке.

	LOCAL	_esi : DWORD
	
	mov		_esi,	esi
	
	push	bx
	xor		bx,		bx ; счетчик templ_num
	
	mov		eax,	templ
	mov		edi,	eax
	xor		ecx,	ecx ; длина текущего шаблона
@@loop:

	mov		al,		byte ptr [esi]	; символ из строки
; ------------------------------------------------------------------------
	mov		ah,		byte ptr [edi]	; символ из шаблона
	cmp		ah,		0				; случаем не конец шаблона ?
	jne		@F
		; конец шаблона - есть совпадение
		mov		eax,	esi
		sub		eax,	ecx
		
		mov		templ_num,	bx
		pop		bx
		
		ret
@@:
; ------------------------------------------------------------------------
	cmp		al,		0				; случаем не конец буфера ?
	jne		@F
		; конец буфера файла - не найдено совпадений
		xor		eax,	eax
		dec		eax				; eax = -1 (дошел до конца буфера файла)
		
		pop		bx
		
		ret
@@:
; ------------------------------------------------------------------------
	cmp		al,		ah
	jne		@F
		inc		esi
		inc		edi
		inc		ecx
		jmp		@@loop
@@:
	
; >>> расхождение ... >>>
	; восстанавливаю буфер файла
	mov		esi,	_esi
	; сбрасываю длину текущего шаблона
	xor		ecx,	ecx
	; увеличиваю счетчик шаблонов
	inc		bx
	
	; ищу конец текущего шаблона
@@:
	inc		edi
	cmp		byte ptr [edi-1],	0
	jne		@B
; <<< расхождение ... <<<
	
	; а не конец ли списка шаблонов ?
	cmp		byte ptr [edi],	0
	jne		@@loop

	; конец, однако...
	; совпадений нет, выхожу
	xor		eax,	eax		; eax = 0 (дошел до конца списка шаблонов)
	
	pop		bx

	ret
equal endp

equal_from_pos proc templ:DWORD, from:DWORD
; # Поиск вхождения набора шаблонов [templ] в буфер по адресу [from] и до конца буфера (#0).
;
;	Вход:
;		templ - указатель на список шаблонов, разделенных #0. в конце еще один #0.
;		from - адрес буфера файла. 
;	Выход:
;		eax:
;			#0	: шаблон не найден
;			any	: адрес в буфере с которого найден шаблон
;		ecx - длина найденного шаблона
;		edi - указывает на #0 за найденным шаблоном
;	Кстати:
;		esi - не меняется

	push	esi
	
	mov		esi,	from
	
@@loop:
	invoke	equal, templ
	cmp		eax,	-1
	jne		@F
		; конец буфера файла - ничего не нашел и выхожу
		pop		esi
		ret
@@:

	cmp		eax,	0		; "test eax, eax" ?
	jne		@F			; если кончился шаблон, то ищу дальше
		inc		esi
		jmp		@@loop
@@:
	
	; что-то найдено - выхожу
	pop		esi 

	ret
equal_from_pos endp

it_is proc templ:DWORD
; # Проверка на наличие в данной позиции данного шаблона.
;
;	Вход:
;		templ - указатель на список шаблонов, разделенных #0. в конце еще один #0.
;		esi - адрес буфера файла.
;	Выход:
;		eax - (!0/0) : совпадает или нет
;	Кстати:
;		esi - не меняется

	push	esi
	invoke	equal,	templ
	pop		esi
	
	cmp		eax,	-1
	jne		@F
		inc		eax
@@:	
	
	ret
it_is endp

ProcessingOneFile proc
; # Обработка файла, уже загруженного в Edit.

	; "номер начального шаблона" должен совпадать с номером закрывающего
	LOCAL	beg_templ_num : WORD

	mov		eax,	old_file_size
	cmp		eax,	0
	jne		@F
		ret
@@:

	push	esi
	push	edi
	
	; Обход содержимого буфера и поиск комментариев или строк
	
	invoke	LocalLock, buff
	mov		buff_ptr,	eax
	
	mov		esi,	eax
	
@@loop:

; #################################### OLC ####################################
	invoke	it_is, addr OLC
	cmp		eax,	0
	je		@@OLC_jump
		mov		edx,	eax
		add		eax,	ecx
		
		invoke	equal_from_pos, addr EOL, eax
		.IF (eax==-1)
			; дошел до конца файла
			; !!! т.к. OLC может быть последним в файле, который может и не содержать соответствующей закрывающей строки,
			; то необходимо проверить, что (edx) не указывает на OLC
			mov		esi,	edx
			invoke	it_is, addr OLC
			cmp		eax,	0
			je		@@isnt_end
				; и правда есть ...
				
				sub		eax,	buff_ptr
				invoke	DelConfirm, edx, eax
				cmp		eax,	IDYES
				jne		@F
					mov		byte ptr [esi], 0 ; а надо ? 
					; пересчитываю размер буфера файла
					mov		eax,	esi
					sub		eax,	buff_ptr
					mov		old_file_size,	eax
@@:

@@isnt_end:

			jmp		@@exit
		.ELSE
			add		eax, ecx	; eax = eax + length(EOL)
			push	eax
			mov		esi, edx

			; теперь блок находится в диапазоне [edx..eax]
			; удаляю его
			mov		ecx, buff_ptr
			sub		edx, ecx
			sub		eax, ecx
			dec		eax
			dec		eax
			
			invoke	DelConfirm, edx, eax
			.IF (eax == IDYES)
				pop		eax
			.ELSE
				pop		esi
				cmp		eax,	IDCANCEL
				jne		@F
					invoke	LocalUnlock, buff
					xor		eax,	eax
					jmp		@@end_func
@@:
			.ENDIF

			jmp	@@continue
		.ENDIF
@@OLC_jump:

; #################################### MLC ####################################
	invoke	it_is, addr MLCB
	cmp		eax,	0
	je		@@MLC_jump
		mov		edx,	eax
		add		eax,	ecx
		
		; сохраняю номер начального шаблона
		m2m		beg_templ_num,	templ_num

; цикл поиска соответствующего шаблона конца многострочного комментария
@@tmpl_num_loop:
		invoke	equal_from_pos, addr MLCE, eax
		.IF (eax==-1)
			; дошел до конца файла
			jmp		@@exit
		.ELSE
		
			; сравниваю номера шаблонов
			mov		si,		beg_templ_num
			cmp		si,		templ_num
			je		@F
				; нашел другой шаблон, ищу дальше
				add		eax,	ecx	; eax = eax + length(MLCE)
				jmp		@@tmpl_num_loop
@@:
			
			; совпадение номеров, продолжаю работу ...
			
			add		eax, ecx	; eax = eax + length(MLCE)
			push	eax
			mov		esi, edx

			; теперь блок находится в диапазоне [edx..eax]
			; удаляю его			
			mov		ecx, buff_ptr
			sub		edx, ecx
			sub		eax, ecx
			
			invoke	DelConfirm, edx, eax
			.IF (eax == IDYES)
				pop		eax
			.ELSE
				pop		esi
				cmp		eax,	IDCANCEL
				jne		@F
					invoke	LocalUnlock, buff
					xor		eax,	eax
					jmp		@@end_func
@@:
			.ENDIF
			
			jmp	@@continue
		.ENDIF
@@MLC_jump:

; #################################### TD #####################################
	invoke	it_is, addr TD
	cmp		eax,	0
	je		@@TD_jump
		mov		edx,	eax
		add		eax,	ecx
		
		invoke	equal_from_pos, addr TD, eax
		.IF (eax==-1)
			; дошел до конца файла
			jmp		@@exit
		.ELSE
			add	eax, ecx	; eax = eax + length(TD)
			mov	esi, eax

			jmp	@@continue
		.ENDIF
@@TD_jump:

; #############################################################################
@@continue:

	; проверка на конец буфера файла
	cmp		byte ptr [esi], 0
	je		@@exit	; если конец, то выхожу
	inc		esi
	jmp		@@loop

@@exit:

	invoke	LocalUnlock, buff

	xor		eax,	eax
	inc		eax

@@end_func:

	pop		edi
	pop		esi
	
	ret
ProcessingOneFile endp

ViewFile proc Path:DWORD
; # Обход файлов в папке.
;
;	Вход:
;		Path - путь к папке, вроде "C:\Windows\"
;	Выход:
;		eax - (0/1) - прервать/продолжить работу

	LOCAL	data : WIN32_FIND_DATA
	LOCAL	fh : DWORD
	
	; перебираю файлы в этой директории
	invoke	FindFirstFile, FileFormat, addr data
	mov		fh,	eax
	
@@loop:
	cmp		fh,	INVALID_HANDLE_VALUE
	je		@@break
	
	mov		eax,	dword ptr [data.cFileName]
	
	cmp		ax,		2Eh		; '.',0
	je	@F
	
	and		eax,	0FFFFFFh
	cmp		eax,	2E2Eh	; '..',0
	je	@F
	
	; в данный момент обрабатывается или папка ЭТОГО уровня или файл

	mov		eax,	data.dwFileAttributes
	and		eax,	FILE_ATTRIBUTE_DIRECTORY
	jnz		@@find_file
	
		invoke	strCopy, tmp_str, addr data.cFileName
		invoke	szLen,	tmp_str
		add		eax,	tmp_str
		mov		word ptr [eax], 003Fh ; искусственно добавляю в конец имени ФАЙЛА '?', чтобы отличать от папок	

		invoke	TreeView_InsertItem, hTmpTreeView, item, tmp_str
	
		invoke	LoadFromFile, addr data.cFileName
		invoke	ProcessingOneFile
		cmp		eax,	0
		jne		@F
			invoke	FindClose, fh
			xor		eax,	eax
			ret
@@:
		invoke	SaveToFile, addr data.cFileName

@@find_file:

	invoke	FindNextFile, fh, addr data
	cmp		eax,	0
	jne		@@loop
	
@@break:

	invoke	FindClose, fh
	
	xor		eax,	eax
	inc		eax

	ret
ViewFile endp

ViewDir proc Path:DWORD
; # Рекурсивный обход папок.
;
;	Вход:
;		Path - путь к папке, вроде "C:\Windows\"
;	Выход:
;		eax - (0/1) - прервать/продолжить работу

	LOCAL	data : WIN32_FIND_DATA
	LOCAL	fh : DWORD
	
	invoke SetCurrentDirectory, Path
	
	; обрабатываю файлы в данной директории
	invoke	ViewFile, Path
	cmp		eax,	0
	jne		@F
		ret
@@:

	; перебираю папки в этой директории	
	invoke	FindFirstFile, SADD("*.*"), addr data
	mov		fh,	eax
	
@@loop:
	cmp		fh,	INVALID_HANDLE_VALUE
	je		@@break
	
	mov		eax,	dword ptr [data.cFileName]
	
	cmp		ax,		2Eh		; '.',0
	je	@F
	
	and		eax,	0FFFFFFh
	cmp		eax,	2E2Eh	; '..',0
	je	@F
	
	; в данный момент обрабатывается или папка ЭТОГО уровня или файл
	mov		eax,	data.dwFileAttributes
	and		eax,	FILE_ATTRIBUTE_DIRECTORY
	jz		@@find_dir
	
		; добавляю подэлемент
		invoke	TreeView_InsertItem, hTmpTreeView, item, addr data.cFileName
		mov		item,	eax
	
		invoke	ViewDir, addr data.cFileName ; рекурсивная обработка вложенной папки
		cmp		eax,	0
		jne		@F
			invoke	FindClose, fh
			xor		eax,	eax
			ret
@@:

@@find_dir:
	invoke	FindNextFile, fh, addr data
	cmp		eax,	0
	jne		@@loop
	
@@break:

	invoke	FindClose, fh
	
	invoke	SetCurrentDirectory, SADD("..")
	
	mov		item,	TreeView_GetParent(hTmpTreeView, item)
	
	xor		eax,	eax
	inc		eax

	ret
ViewDir endp

FileProcess proc Path:DWORD, Format:DWORD ; базовый путь и шаблон для поиска
; # Обработка файлов.
;
;	Вход:
;		Path - начальный путь, вроде "C:\Windows\"
;		Format - шаблон вроде "*.pas"
;	Выход:
;		---

	invoke	ShowWindow, WND, SW_MAXIMIZE

	m2m		FileFormat,	Format
	
	invoke	TreeView_InsertItem, hTmpTreeView, NULL, StartPath
	
	mov		item,	TreeView_GetRoot(hTmpTreeView) ; короче, ^^^ eax ^^^
	
	invoke	ViewDir, Path
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, WND, SADD("Обработка была отменена!"), addr AppName, MB_OK or MB_ICONINFORMATION
@@:
	
	invoke	Clear
	
	invoke	ShowWindow, WND, SW_RESTORE
	
	invoke	FileListDlgShow
	
	ret
FileProcess endp