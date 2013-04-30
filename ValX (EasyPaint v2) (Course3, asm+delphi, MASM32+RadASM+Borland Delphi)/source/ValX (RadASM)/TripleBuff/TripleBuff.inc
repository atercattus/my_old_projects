;##################################################################
;    ...::: Valkiria :::...
;
; Файл     : TripleBuff.inc
; Описание : Основной модуль для реализации тройной буферизации
; Версия   : 24.03
;
;##################################################################
.data
	MyPen		HPEN	0
	MyBrush		HBRUSH	0

.data?
	; приемный контекст
	hDC			HDC		?

	; контекст заднего буфера клиентской области
	hBackDC		HDC		?
	
	; размер изображения
	Rect		RECT	<>
	
	; реальный размер клиентской области (с учетом TB & SB)
	ClientSize	POINT	<>
	; координаты левого верхнего угла изображения в координатах реальной клиентской области (как бы TB нет!)
	LeftTop		POINT	<>
	; размер региона для блиттинга
	SourceRect	RECT	<>
	DestRect	RECT	<>
	; есть ли что отрисовывать (видна ли хоть часть изображения в пределах клиентской области): bool
	MayFlip		DWORD	?

	; толщина карандаша
	CurPenSize	DWORD	?

.const
	; максимальный размер карандаша (минимальный == 1)
	MaxPenSize equ 100
	
	; минимальный размер видимой части по оси Y
	scrollY_min		equ		10
	; минимальный размер видимой части по оси X
	scrollX_min		equ		10
	
.code

; macro
set_false MACRO var
	xor		eax,	eax
	mov		var,	eax
	ret
ENDM

set_true MACRO var
	xor		eax,	eax
	inc		eax
	mov		var,	eax
ENDM

DSNewStr MACRO name,len
	.data
		name db len dup (0)
	.code
ENDM

DRAW_AND_FLIP MACRO
	invoke	DrawClientRect
	invoke	FlipClientBackBuff2Wnd
ENDM

FLIP_AND_FLIP MACRO
	invoke	Flip
	invoke	FlipClientBackBuff2Wnd
ENDM

DrawMiniWnd proto :HDC

; TripleBuff.inc (This file)
	SetPenParam proto Color:DWORD, W:DWORD
	SetBrushColor proto Color:DWORD
	PutPixel proto X:DWORD, Y:DWORD, Color:DWORD
	ScrollX proto d:DWORD
	ScrollY proto d:DWORD	

; ImgMemBuff.inc
	CreateBackBuffer proto hWnd:HWND, W:DWORD, H:DWORD, OldSave:DWORD
	DeleteBackBuffer proto
	Clear proto
	LoadFromFile proto path:DWORD
	SaveToFile proto path:DWORD

; ClientBackBuff.inc
	CreateClientBackBuff proto hWnd:DWORD
	DeleteClientBackBuff proto
	ResizeClientBackBuff proto
	; копирование буфера изображения в буфер клиентской области в памяти
	Flip proto
	; копирование клиенского буфера в контекст окна
	FlipClientBackBuff2Wnd proto
	Resize proto
	DrawClientRect proto
	CalcRgns proto

; модуль для работы с буфером изображения в памяти
include ImgMemBuff.inc
; модуль для работы с задним буфером клиентской области
include ClientBackBuff.inc

;##################################################################################
;                  ScrollX
;##################################################################################
ScrollX proc d:DWORD
	mov		eax,	d
	add		eax,	LeftTop.x			; eax = LT.x`
	
	mov		ecx,	scrollX_min
	sub		ecx,	Rect.right			; ecx = scrollX_min - Rect.w
	cmp		eax,	ecx
	jg		@@1
		xchg	eax,	ecx				; if (LT.x` < scrollX_min - Rect.w) ; за левой границей ...
		jmp		@@2						; ... LT.x` = scrollX_min - Rect.w ; выставляю минимум
@@1:
		mov		edx,	ClientSize.x
		sub		edx,	scrollX_min		; edx = ClientSize.w - scrollX_min
		cmp		edx,	eax
		jnl		@F						; if (ClientSize.w-scrollX_min < LT.x`) ; за правой границей ...
			xchg	eax,	edx			; ... LT.x` = ClientSize.w-scrollX_min ; выставляю максимум
		@@:
@@2:
	
	mov		LeftTop.x,	eax
	
	ret
ScrollX endp

;##################################################################################
;                  ScrollY
;##################################################################################
ScrollY proc d:DWORD
	mov		eax,	d
	add		eax,	LeftTop.y			; eax = LT.y`

	mov		ecx,	scrollY_min
	sub		ecx,	Rect.bottom			; ecx = scrollY_min - Rect.h
	cmp		eax,	ecx
	jg		@@1
		xchg	eax,	ecx				; if (LT.y` < scrollY_min - Rect.h) ; за верхней границей ...
		jmp		@@2						; ... LT.y` = scrollY_min - Rect.h ; выставляю минимум
@@1:
		mov		edx,	ClientSize.y
		sub		edx,	scrollY_min		; edx = ClientSize.h - scrollY_min
		cmp		edx,	eax
		jnl		@F						; if (ClientSize.h-scrollY_min < LT.y`) ; за нижней границей ...
			xchg	eax,	edx			; ... LT.y` = ClientSize.h-scrollY_min ; выставляю максимум
		@@:
@@2:
	
	mov		LeftTop.y,	eax

	ret
ScrollY endp

;##################################################################################
;                  SetPenParam
;##################################################################################
SetPenParam proc Color:DWORD, W:DWORD
	LOCAL num[20]:BYTE

	invoke	CreatePen, PS_SOLID, W, Color
	mov		MyPen,	eax
	invoke	SelectObject, hCompDC, MyPen
	invoke	DeleteObject, eax
	
	invoke	SelectObject, hBackDC, MyPen
	invoke	DeleteObject, eax
	
	; меняю текст в строке состояния
	dsText frmt, " Толщина линии: %lu"
	invoke	wsprintf, addr num, addr frmt, W
	invoke	SendMessage, hStatusBar, SB_SETTEXT, 1, addr num
	
	; сохраняю
	m2m		FGColor,	Color
	
	m2m		CurPenSize,	W
	
	ret
SetPenParam endp

;##################################################################################
;                  SetBrushColor
;##################################################################################
SetBrushColor proc Color:DWORD
	LOCAL	LB : LOGBRUSH

	mov		LB.lbStyle,		BS_SOLID
	m2m		LB.lbColor,		Color	

	invoke	CreateBrushIndirect, addr LB
	mov		MyBrush,	eax
	
	invoke	SelectObject, hCompDC, MyBrush
	invoke	DeleteObject, eax
	
	invoke	SelectObject, hBackDC, MyBrush
	invoke	DeleteObject, eax
	
	; сохраняю
	m2m		BGColor,	Color
	
	ret
SetBrushColor endp

;##################################################################################
;                  PutPixel
;##################################################################################
PutPixel proc X:DWORD, Y:DWORD, Color:DWORD
	LOCAL	pc, bc : DWORD

	mov		eax,	CurPenSize
	.IF (eax==1)
		invoke	SetPixel, hCompDC, X, Y, Color
	.ELSE
	
		m2m		pc,	FGColor
		m2m		bc,	BGColor
		invoke	SetPenParam, Color, CurPenSize
		invoke	SetBrushColor, Color
	
		mov		eax,	CurPenSize
		shr		eax,	1
		mov		ecx,	X
		mov		edx,	Y
		add		ecx,	eax
		add		edx,	eax
		
		push	edx
		push	ecx
		
		neg		eax
		lea		ecx,	[ecx+2*eax]
		lea		edx,	[edx+2*eax]
		
		push	edx
		push	ecx
		
		push	hCompDC
		
		call	Ellipse
		
		invoke	SetPenParam, pc, CurPenSize
		invoke	SetBrushColor, bc
	.ENDIF
	
	ret
PutPixel endp