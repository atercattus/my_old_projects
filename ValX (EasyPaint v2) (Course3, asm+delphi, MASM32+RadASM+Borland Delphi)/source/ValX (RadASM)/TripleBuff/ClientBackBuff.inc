;##################################################################
;    ...::: Valkiria :::...
;
; Файл     : ClientBackBuff.inc
; Описание : Реализация заднего буфера клиентской области
; Версия   : 11.03
;
;##################################################################
.const
	; минимальный размер стороны изображения, чтобы программа рисовала тень от листа
	MIN_IMAGE_SIZE_WITH_SHADOW	equ	15

.data?
	hBackCompBtmp	HBITMAP	?
	hBackOldBtmp	HBITMAP	?
	
.code

;##################################################################################
;                  CreateClientBackBuff
;##################################################################################
CreateClientBackBuff proc hWnd:DWORD
	LOCAL	_hDC : HDC
	LOCAL	hRgn : HRGN

	cmp		ClientSize.y,	0
	jg		@F
		ret
	@@:
	
	invoke	GetDC,	hWnd
	mov		_hDC,	eax

	invoke	CreateCompatibleDC, _hDC
	mov		hBackDC, eax
	
	mov		eax,	ClientSize.y
	inc		eax
	invoke	CreateCompatibleBitmap, _hDC, ClientSize.x, eax
	mov		hBackCompBtmp,	eax
	
	invoke	SelectObject, hBackDC, eax
	mov		hBackOldBtmp, eax
	
	invoke	ReleaseDC, hWnd, _hDC

	ret
CreateClientBackBuff endp

;##################################################################################
;                  DeleteClientBackBuff
;##################################################################################
DeleteClientBackBuff proc

	invoke	SelectObject, hBackDC, hBackOldBtmp
	invoke	DeleteObject, hBackCompBtmp
	invoke	DeleteDC, hBackDC
	mov		hBackDC,	0
	
	ret
DeleteClientBackBuff endp

;##################################################################################
;                  ResizeClientBackBuff
;##################################################################################
ResizeClientBackBuff proc

	invoke	DeleteClientBackBuff
	invoke	CreateClientBackBuff, WND

	invoke	SelectObject, hBackDC, MyPen
	invoke	DeleteObject, eax
	
	invoke	SelectObject, hBackDC, MyBrush
	invoke	DeleteObject, eax
	
	ret
ResizeClientBackBuff endp

;##################################################################################
;                  FlipClientBackBuff2Wnd
;##################################################################################
DrawMiniWnd proto DC:HDC
FlipClientBackBuff2Wnd proc
	
	mov		eax,	ClientSize.y
	inc		eax	
	invoke	BitBlt, hDC, 0, TOP_DELTA, ClientSize.x, eax, hBackDC, 0, 0, SRCCOPY
	
	ret
FlipClientBackBuff2Wnd endp

;##################################################################################
;                  Resize
;##################################################################################
Resize proc
	LOCAL R : RECT

	invoke	GetClientRect, WND, addr R
	m2m		ClientSize.x,	R.right
	
	mov		eax,	R.bottom
	sub		eax,	TOP_DELTA
	sub		eax,	SBHeight
	mov		ClientSize.y,	eax
	mov		MayFlip,	eax
	; если высота не больше 0, то будет выводится все (ОШИБКА!)
	cmp		eax,	0
	jng		@@err
	
	; рассчитываю регионы для новых размеров области вывода
	invoke	CalcRgns

	invoke	ResizeClientBackBuff
	
	ret

@@err:
	set_false MayFlip

	ret
Resize endp

;##################################################################################
;                  DRAW_PAGE_EDGE
;##################################################################################
DRAW_PAGE_EDGE proc
	.data?
		oP	HPEN	?
		oB	HBRUSH	?
		X1	DWORD	?
		Y1	DWORD	?
		X2	DWORD	?
		Y2	DWORD	?
	.code
	
	push	ebx
	
; Вывод границ
	invoke	GetStockObject, BLACK_PEN
	invoke	SelectObject, hBackDC, eax
	mov		oP,	eax
	
	mov		eax,	LeftTop.x		; eax = left
	mov		ebx,	eax
	add		ebx,	Rect.right		; ebx = right
	
	mov		ecx,	LeftTop.y		; ecx = top
	mov		edx,	ecx
	add		edx,	Rect.bottom		; edx = bottom
	
	dec		eax
	dec		ecx
	inc		ebx
	inc		edx
	
	mov		X1,		eax
	mov		Y1,		ecx
	mov		X2,		ebx
	mov		Y2,		edx
	
	invoke	Rectangle, hBackDC, eax, ecx, ebx, edx

; проверяю размер изображения
	cmp		Rect.right,	MIN_IMAGE_SIZE_WITH_SHADOW
	jnb		@F
		pop		ebx
		jmp		@@MACRO_EXIT
	@@:
	
	cmp		Rect.bottom,	MIN_IMAGE_SIZE_WITH_SHADOW
	jnb		@F
		pop		ebx
		jmp		@@MACRO_EXIT
	@@:

; вывод тени
	invoke	GetStockObject, LTGRAY_BRUSH
	invoke	SelectObject, hBackDC, eax
	mov		oB,	eax
	
	invoke	GetStockObject, NULL_PEN
	invoke	SelectObject, hBackDC, eax
	
	; RIGHT
	mov		eax,	X2
	mov		ecx,	Y1
	add		ecx,	5
	push	ecx			; y2 = Y1+5
	push	eax			; x2 = X2
	
	add		eax,	8
	mov		ecx,	Y2
	inc		ecx
	push	ecx			; y2 = Y2+1
	push	eax			; x2 = X2+8
	
	push	hBackDC
	call	Rectangle
	
	; BOTTOM
	mov		eax,	X1
	mov		ecx,	Y2
	add		eax,	5
	push	ecx			; y2 = Y2
	push	eax			; x2 = X1+5
	
	mov		eax,	X2
	add		ecx,	8
	add		eax,	8
	push	ecx			; y1 = Y2+8
	push	eax			; x1 = X2+8
	
	push	hBackDC
	call	Rectangle
	
; восстановление параметров
	pop		ebx

	invoke	SelectObject, hBackDC, oB
@@MACRO_EXIT:
	invoke	SelectObject, hBackDC, oP
	ret
DRAW_PAGE_EDGE endp

;##################################################################################
;                  DrawClientRect
;##################################################################################
DrawClientRect proc
	mov		eax,	hBackDC
	test	eax,	eax
	jz		@@end
		mov		eax,	ClientSize.y
		cmp		eax,	0
		jl		@F
			inc		eax
			invoke	PatBlt, hBackDC, 0, 0, ClientSize.x, eax, WHITENESS
			invoke	DRAW_PAGE_EDGE
			invoke	Flip
		@@:
@@end:
	ret
DrawClientRect endp

;##################################################################################
;                  CalcRgns
;
; Расчет регионов исходной и приемной областей при блиттинге
;##################################################################################
CalcRgns proc
	push	ebx ; !!!
	
	; MayFlip = true(1)
	set_true MayFlip

; ########### X ###################
; if (LT.x >= 0) { S.x = 0; D.x = LT.x; } else { S.x = |LT.x|; D.x = 0; }
	mov		eax,	LeftTop.x
	cmp		eax,	0
	jl		@@1
		; (LT.x >= 0)
		; S.x = 0; D.x = LT.x;
		xor		edx,	edx		; edx = SourceRect.left
								; eax = DestRect.left
		jmp		@@2
	@@1:
		; (LT.x < 0)
		; S.x = |LT.x|; D.x = 0;
		neg		eax				; eax = |LT.x|
		xor		edx,	edx		; edx = 0
		xchg	eax,	edx		; соблюдаю предназначение регистров
	@@2:
	mov		SourceRect.left,	edx
	mov		DestRect.left,		eax
	
; ########### Y ###################
; if (LT.y >= 0) { S.y = 0; D.y = LT.y; } else { S.y = |LT.y|; D.y = 0; }
	mov		eax,	LeftTop.y
	cmp		eax,	0
	jl		@@3
		; (LT.y >= 0)
		; S.y = 0; D.y = LT.y;
		xor		edx,	edx		; edx = SourceRect.top
								; eax = DestRect.top
		jmp		@@4
	@@3:
		; (LT.y < 0)
		; S.y = |LT.y|; D.y = 0;
		neg		eax				; eax = |LT.y|
		xor		edx,	edx		; edx = 0
		xchg	eax,	edx		; соблюдаю предназначение регистров
	@@4:
	mov		SourceRect.top,	edx
	mov		DestRect.top,	eax
	
; ########### W ###################
; A = ClientSize.w - LT.x
; if (R.w <= A) { S.w = R.w; } else { S.w = A; }
; D.w = S.w;
	; A
	mov		ebx,	ClientSize.x
	sub		ebx,	LeftTop.x
	; S.w
	mov		eax,	Rect.right		; eax{S.w} = R.w
	cmp		eax,	ebx
	jng		@F
		mov		eax,	ebx			; eax{S.w} = A
	@@:
	cmp		eax,	0				; ширина отрицательная
	jg		@F
		pop		ebx
		set_false	MayFlip
	@@:
	mov		SourceRect.right,	eax
	mov		DestRect.right,		eax
	
	add		eax,	DestRect.left		; eax = D.x + D.w
	cmp		eax,	0					; если приемная область за левой границей
	jnl		@F
		pop		ebx
		set_false	MayFlip
	@@:
	
; ########### H ###################
; B = ClientSize.h - LT.y
; if (R.h <= B) { S.h = R.h; } else { S.h = B; }
; D.h = S.h;
	; B
	mov		ebx,	ClientSize.y
	sub		ebx,	LeftTop.y
	; S.h
	mov		eax,	Rect.bottom		; eax{S.h} = R.h
	cmp		eax,	ebx
	jng		@F
		mov		eax,	ebx			; eax{S.h} = B
	@@:
	inc		eax
	cmp		eax,	0				; высота отрицательная
	jg		@F
		pop		ebx
		set_false	MayFlip
	@@:
	mov		SourceRect.bottom,	eax
	mov		DestRect.bottom,	eax
	
	add		eax,	DestRect.top		; eax = D.y + D.h 
	cmp		eax,	0					; если приемная область за верхней границей
	jnl		@F
		pop		ebx
		set_false	MayFlip
	@@:
		
	pop		ebx ; !!!
	
	ret
CalcRgns endp