.data
	hCompDC		HDC		0
	
.data?
	hDC			HDC		?
	Rect		RECT	<>
	hCompBtmp	HBITMAP	?
	hOldBtmp	HBITMAP	?
	
	AreaHeight	DWORD	?

	; толщина карандаша
	CurPenSize	DWORD	?
	
.code

;---------------------------------------------------------------------------------
;                   DeleteBackBuffer
;---------------------------------------------------------------------------------
DeleteBackBuffer proc
	mov		eax,	hCompDC
	cmp		eax,	0
	je		@F
		invoke	SelectObject, hCompDC, hOldBtmp
		invoke	DeleteObject, hCompBtmp
		invoke	DeleteDC, hCompDC
		mov		hCompDC,	0		
@@:
  
    ret
    
DeleteBackBuffer endp

Clear proto

;---------------------------------------------------------------------------------
;                   CreateBackBuffer
;---------------------------------------------------------------------------------
CreateBackBuffer proc hWnd:HWND, W:DWORD, H:DWORD, OldSave:DWORD
LOCAL	_hDC	: HDC
LOCAL	_hCompDC : HDC
LOCAL	_Rect : RECT
LOCAL	_hCompBtmp : HBITMAP
LOCAL	_hOldBtmp : HBITMAP

	mov		_Rect.left, 0
	mov		eax, TBHeight
	mov		_Rect.top, eax
	m2m		_Rect.right, W
	m2m		_Rect.bottom, H
	
	invoke	GetDC, hWnd
	mov		_hDC, eax
	
	invoke	CreateCompatibleDC, _hDC
	mov		_hCompDC, eax
	
	invoke	CreateCompatibleBitmap, _hDC, _Rect.right, _Rect.bottom
	mov		_hCompBtmp,	eax
	
	invoke	SelectObject, _hCompDC, eax
	mov		_hOldBtmp, eax
	
	invoke	ReleaseDC, hWnd, _hDC
	
	cmp		hCompDC,	0
	je		@@neednt_copy
		; очищаю фон 
		invoke	BitBlt, _hCompDC, 0, 0, _Rect.right, _Rect.bottom, NULL, 0, 0, WHITENESS
		
		; если нужно, копирую старое содержимое 
		.IF (OldSave)
			invoke	BitBlt, _hCompDC, 0, 0, Rect.right, Rect.bottom, hCompDC, 0, 0, SRCCOPY
		.ENDIF
		
		; удал€ю старый буфер
		invoke	SelectObject, hCompDC, hOldBtmp
		invoke	DeleteObject, hCompBtmp
		invoke	DeleteDC, hCompDC
		
		; и сохран€ю новые значени€
		m2m		hCompDC,	_hCompDC
		m2m		hCompBtmp,	_hCompBtmp
		m2m		hOldBtmp,	_hOldBtmp
		invoke	RtlMoveMemory, addr Rect, addr _Rect, sizeof RECT
		
		jmp		@F
		
@@neednt_copy:
	; сохран€ю новые значени€
	m2m		hCompDC,	_hCompDC
	m2m		hCompBtmp,	_hCompBtmp
	m2m		hOldBtmp,	_hOldBtmp
	invoke	RtlMoveMemory, addr Rect, addr _Rect, sizeof RECT
	; очищаю фон
	invoke	Clear

@@:		

	mov		eax,	1
	ret

CreateBackBuffer endp

;---------------------------------------------------------------------------------
;                  Clear
;---------------------------------------------------------------------------------
Clear proc

	mov		eax,	hCompDC
	cmp		eax,	0
	je		@F
		invoke	BitBlt, eax, 0, 0, Rect.right, Rect.bottom, NULL, 0, 0, WHITENESS
@@:
	  
    ret
Clear endp

;---------------------------------------------------------------------------------
;                  Flip
;---------------------------------------------------------------------------------
Flip proc

	mov		eax,	hCompDC
	cmp		eax,	0
	je		@@end
		invoke	BitBlt, hDC, 0, TBHeight, Rect.right, AreaHeight, hCompDC, 0, 0, SRCCOPY
@@end:

	ret
Flip endp

;---------------------------------------------------------------------------------
;                  Resize
;---------------------------------------------------------------------------------
Resize proc
	LOCAL R : RECT
	
	invoke	GetClientRect, WND, addr R
	mov		ecx,	R.bottom
	sub		ecx,	TBHeight
	sub		ecx,	SBHeight
	; если высота меньше 0, то будет выводитс€ все (ќЎ»Ѕ ј!)
	cmp		ecx,	0
	jnl		@F
	ret
@@:
	mov		edx,	Rect.bottom
	; если высота выводимой области меньше высоты изображени€, то нужно усекать изображение
	cmp		ecx,	edx
	jna		@F
	xchg	ecx,	edx
@@:

	mov	AreaHeight,	ecx
	
	ret
Resize endp

.data ; строковые ресурсы
	sERROR db "ќшибка!",0
	sINFO db "—ообщение",0
	sGetParamError db "ќшибка получени€ параметров изображени€!",0
	sWriteFileError db "ќшибка записи в файл!",0
	
.code

SetPenParam proto Color:DWORD, W:DWORD
SetBrushColor proto Color:DWORD

LoadFromFile proc path:DWORD
	LOCAL bmp : HBITMAP
	LOCAL DC : HDC
	LOCAL B : BITMAP

	cmp		hCompDC,	0
	jne		@F
		ret
@@:

	; загружаю битмап
	invoke	LoadImage, NULL, path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, 0, SADD("ќшибка при загрузке изображени€!"), addr sERROR, MB_OK or MB_ICONERROR
		ret
@@:
	mov		bmp,	eax
	; получил HBITMAP
	
	; создаю временный контекст
	invoke	CreateCompatibleDC, hCompDC
	mov		DC, eax
	
	; объедин€ю контекст с битмапом
	invoke	SelectObject, DC, bmp

	; получаю размеры битмапа
	invoke	GetObject, bmp, sizeof BITMAP, addr B
	
	.IF (!eax)
		invoke	MessageBox, 0, addr sGetParamError, addr sERROR, MB_OK or MB_ICONERROR
		
		invoke	SelectObject, DC, NULL
		invoke	DeleteObject, bmp
		invoke	DeleteDC, DC
		
		ret
	.ENDIF
	
	; создаю новый буфер
	invoke	CreateBackBuffer, WND, B.bmWidth, B.bmHeight, FALSE
	
	; копирую в него битмап
	invoke	BitBlt, hCompDC, 0, 0, B.bmWidth, B.bmHeight, DC, 0, 0, SRCCOPY
	
	
	; восстанавливаю параметры инструментов
	invoke	SetPenParam, FGColor, CurPenSize
	invoke	SetBrushColor, BGColor
	
	; перерисовываю окно
	invoke	InvalidateRect, WND, NULL, TRUE
	
	; и удал€ю мусор
	invoke	SelectObject, DC, NULL
	
	invoke	DeleteObject, bmp
	
	invoke	DeleteDC, DC
	
	; конец
	invoke	MessageBox, 0, SADD("»зображение загружено!"), addr sINFO, MB_OK or MB_ICONINFORMATION	
	
	ret
LoadFromFile endp

SaveToFile proc path:DWORD
	LOCAL	bfh : BITMAPFILEHEADER
	LOCAL	BI : BITMAPINFO
	LOCAL	f : DWORD
	LOCAL	wb : DWORD
	LOCAL	buff : DWORD
	LOCAL	BMPParam : BITMAP
	LOCAL	BMPSize : DWORD
	

; провер€ю на наличие изображени€ ###############################################################################
	cmp		hCompDC,	0
	jne		@F
		invoke	MessageBox, 0, SADD("ј что мне сохран€ть то?"), addr sERROR, MB_OK or MB_ICONERROR
		ret
@@:

; получаю размеры битмапа #######################################################################################
	invoke	GetObject, hCompBtmp, sizeof BITMAP, addr BMPParam
	.IF (!eax)
		invoke	MessageBox, 0, addr sGetParamError, addr sERROR, MB_OK or MB_ICONERROR
		ret
	.ENDIF
	
; вычисл€ю размер изображени€ в байтах и выдел€ю пам€ть #########################################################
	mov		eax,	BMPParam.bmWidthBytes
	mov		ecx,	BMPParam.bmHeight
	xor		edx,	edx
	mul		ecx
	mov		BMPSize,	eax
	
	invoke	GlobalAlloc, GMEM_FIXED, BMPSize
	.IF (!eax)
		invoke	MessageBox, 0, SADD("ќшибка выделени€ пам€ти!"), addr sERROR, MB_OK or MB_ICONERROR
		ret
	.ENDIF
		
	mov		buff,	eax
	
; получаю данные битмапа ########################################################################################
	invoke	RtlZeroMemory, addr BI, sizeof BITMAPINFO
	mov		BI.bmiHeader.biSize,	sizeof BITMAPINFOHEADER
	m2m		BI.bmiHeader.biWidth,	BMPParam.bmWidth
	m2m		BI.bmiHeader.biHeight,	BMPParam.bmHeight
	m2m		BI.bmiHeader.biPlanes,	BMPParam.bmPlanes
	m2m		BI.bmiHeader.biBitCount,	BMPParam.bmBitsPixel
	mov		BI.bmiHeader.biCompression,	BI_RGB

	invoke	GetDIBits, hCompDC, hCompBtmp, 0, BMPParam.bmHeight, buff, addr BI, DIB_RGB_COLORS
	.IF (!eax)
		invoke	MessageBox, 0, addr sGetParamError, addr sERROR, MB_OK or MB_ICONERROR
		invoke	GlobalFree, buff
		ret
	.ENDIF
	
; создаю файл ###################################################################################################
	invoke	CreateFile, path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL
	cmp		eax,	INVALID_HANDLE_VALUE
	jne		@F
		invoke	MessageBox, 0, SADD("ќшибка при открытии файла!"), addr sERROR, MB_OK or MB_ICONERROR
		invoke	GlobalFree, buff
		ret
@@:
	mov		f,	eax
	

; заполн€ю BITMAPFILEHEADER #####################################################################################
	invoke	RtlZeroMemory, addr bfh, sizeof BITMAPFILEHEADER
	
	mov		bfh.bfType,	4D42h

	mov		eax,	BMPSize
	add		eax,	(sizeof BITMAPFILEHEADER)+(sizeof BITMAPINFOHEADER)
	mov		bfh.bfSize,	eax
	
	mov		bfh.bfOffBits,	54

; записываю BITMAPFILEHEADER в файл #############################################################################
	invoke	WriteFile, f, addr bfh, sizeof BITMAPFILEHEADER, addr wb, NULL
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, 0, addr sWriteFileError, addr sERROR, MB_OK or MB_ICONERROR
		invoke	GlobalFree, buff
		invoke	CloseHandle, f
		ret
@@:

; записываю BITMAPINFOHEADER в файл #############################################################################
	invoke	WriteFile, f, addr BI.bmiHeader, sizeof BITMAPINFOHEADER, addr wb, NULL
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, 0, addr sWriteFileError, addr sERROR, MB_OK or MB_ICONERROR
		invoke	GlobalFree, buff
		invoke	CloseHandle, f
		ret
@@:

; записываю сам битовый массив ##################################################################################
	invoke	WriteFile, f, buff, BMPSize, addr wb, NULL
	cmp		eax,	0
	jne		@F
		invoke	MessageBox, 0, addr sWriteFileError, addr sERROR, MB_OK or MB_ICONERROR
		invoke	GlobalFree, buff
		invoke	CloseHandle, f
		ret
@@:

; освобождаю ресурсы ############################################################################################

	invoke	GlobalFree, buff
	invoke	CloseHandle, f
	invoke	MessageBox, 0, SADD("»зображение сохранено!"), addr sINFO, MB_OK or MB_ICONINFORMATION
		
	ret
SaveToFile endp