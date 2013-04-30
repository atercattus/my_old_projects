;##################################################################
;    ...::: Valkiria :::...
;
; ‘айл     : MiniWnd.inc
; ќписание : –еализаци€ окна "Ёскиз"
; ¬ерси€   : 26.03
;
;##################################################################
.data?
	; стандартна€ "пуста€" кисть
	NULL_Brush		HBRUSH	?
	; размеры клиентской области
	MiniWidth		DWORD	?
	MiniHeight		DWORD	?
	
	; задний контекст окна (хранитс€ ”∆≈ отмасштабированна€ копи€)
	MWndDC			HDC		?
	MWndBmp			HBITMAP	?

.const
	MiniWndClassName	db "ValX_MiniWnd",0
	MiniWndAppName		db "Ёскиз",0 ; "Thumbnail",0

	MINI_WND_STYLE		equ	WS_CAPTION or WS_SYSMENU or WS_POPUP
	
.data
	; коэффициент масштабировани€: –еальный–азмер/–азмерЁскиза
	SCALE_KOEF			DWORD	6

.code

MiniWndProc proto hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
DrawMiniWnd proto DC:HDC
ResizeMiniWnd proto goto_RightTop_angle:DWORD 

THUMBNAIL_IMAGE proc
; ѕерерисовка забнего буфера в клиентскую область
	invoke	InvalidateRect, hMiniWnd, NULL, 0
	ret
THUMBNAIL_IMAGE endp

THUMBNAIL_MAKE proc
; Ѕлиттинг с масштабированием исходного изображени€ в задний буфер
	invoke	SetStretchBltMode, MWndDC, 4
	invoke	StretchBlt, MWndDC, 0, 0, MiniWidth, MiniHeight, hCompDC, 0, 0, Rect.right, Rect.bottom, SRCCOPY
	invoke	SetStretchBltMode, MWndDC, 2
	ret
THUMBNAIL_MAKE endp

CreateMiniWnd proc
	LOCAL	wc		: WNDCLASSEX
	LOCAL	R		: RECT
	LOCAL	_hDC	: HDC
	
; —ќ«ƒјё ќ Ќќ Ё— »«ј
	invoke	RtlZeroMemory, addr wc, sizeof WNDCLASSEX
	mov		wc.cbSize, sizeof WNDCLASSEX
	mov		wc.style, CS_OWNDC or CS_HREDRAW or CS_VREDRAW
	mov		wc.lpfnWndProc, offset MiniWndProc
	push	hInstance
	pop		wc.hInstance
	mov		wc.lpszClassName, offset MiniWndClassName
	m2m		wc.hCursor,	StdCursor
	
	invoke	RegisterClassEx, addr wc
	
	invoke	CreateWindowEx,\
				WS_EX_TOOLWINDOW,\
				addr MiniWndClassName,\
				addr MiniWndAppName,\
				MINI_WND_STYLE,\
				; сейчас все выставлю
				CW_USEDEFAULT,\
				CW_USEDEFAULT,\
				CW_USEDEFAULT,\
				CW_USEDEFAULT,\
				WND,\
				NULL,\
				NULL,\
				NULL
	mov		hMiniWnd,	eax
	
	; подгон€ю размер Ёскиза под размер изображени€ и коэффициент масштабировани€
	invoke	ResizeMiniWnd, TRUE
	
; —ќ«ƒјё «јƒЌ»…  ќЌ“≈ —“ ƒЋя Ё— »«ј
	invoke	GetDC,	hMiniWnd
	mov		_hDC,	eax
	
	invoke	CreateCompatibleDC,	_hDC
	mov		MWndDC,		eax
	
	invoke	CreateCompatibleBitmap, _hDC, MiniWidth, MiniHeight
	mov		MWndBmp,	eax
	
	invoke	SelectObject, MWndDC, eax
	invoke	DeleteObject, eax
	
	invoke	ReleaseDC, hMiniWnd, _hDC

; »Ќ»÷»јЋ»«ј÷»я » ќ“ќЅ–ј∆≈Ќ»≈

	invoke	THUMBNAIL_MAKE
	
	invoke	GetStockObject, NULL_BRUSH
	mov		NULL_Brush,		eax
	
	invoke	ShowWindow, hMiniWnd, SW_SHOWDEFAULT
;	invoke	UpdateWindow, hMiniWnd
	
	ret
CreateMiniWnd endp

DestroyMiniWnd proc
	invoke	DeleteObject, MWndBmp
	invoke	DeleteDC, MWndDC
	invoke	DestroyWindow, hMiniWnd
	ret
DestroyMiniWnd endp

MiniWndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
	LOCAL	PS : PAINTSTRUCT

	.IF (uMsg==WM_PAINT)
		invoke	BeginPaint, hMiniWnd, addr PS
		invoke	DrawMiniWnd, eax
		invoke	EndPaint, hMiniWnd, addr PS
		
	.ELSEIF (uMsg==WM_CLOSE)
		; если пытаюсь закрыть, то пр€чу и снимаю галку в меню
		invoke	ShowWindow, hMiniWnd, FALSE
		invoke	CheckMenuItem, hMenu, IDM_WND_THUMBNAIL, MF_BYCOMMAND or MF_UNCHECKED
		; указываю системе, что не надо удал€ть окно
		xor		eax,	eax
		
	.ELSEIF (uMsg==WM_DESTROY)
		invoke	PostQuitMessage, 0
		
	.ELSE
		invoke	DefWindowProc, hWnd, uMsg, wParam, lParam
		ret
	.ENDIF
	
	ret
MiniWndProc endp

DrawMiniWnd proc myDC:HDC
	LOCAL	R1	: RECT
	LOCAL	_oB	: HBRUSH

	invoke	SelectObject, myDC, NULL_Brush
	mov		_oB,	eax

; ‘ќЌ
	; быстренько копирую из заднего буфера в экранный контекст
	invoke	BitBlt, myDC, 0, 0, MiniWidth, MiniHeight, MWndDC, 0, 0, SRCCOPY
	
; ¬»ƒ»ћјя ќЅЋј—“№
; X
; if (LT.x < 0) { MLT.x = -LT.x*(M.w/R.w); } else { MLT.x = 0; }
	mov		eax,	LeftTop.x
	cmp		eax,	0
	jnl		@@1
		neg		eax
		mul		MiniWidth
		mov		ecx,	Rect.right
		test	ecx,	ecx
		jz		@@EXIT
		div		ecx
		jmp		@@2
@@1:
		xor		eax,	eax
@@2:
	mov		R1.left,		eax

; Y
; if (LT.y < 0) { MLT.y = -LT.y*(M.h/R.h); } else { MLT.y = 0; }
	mov		eax,	LeftTop.y
	cmp		eax,	0
	jnl		@@3
		neg		eax
		mul		MiniHeight
		mov		ecx,	Rect.bottom
		test	ecx,	ecx
		jz		@@EXIT
		div		ecx
		jmp		@@4
@@3:
		xor		eax,	eax
@@4:
	mov		R1.top,		eax

; W
; W = max(M.w*(Dest.w*R.w), 2);
	mov		eax,	DestRect.right
	mul		MiniWidth
	mov		ecx,	Rect.right
	div		ecx
	cmp		eax,	2
	jg		@F
		mov		eax,	2
	@@:
	mov		R1.right,	eax
	
; H
; H = max(M.h*(Dest.h/R.h), 2);
	mov		eax,	DestRect.bottom
	mul		MiniHeight
	mov		ecx,	Rect.bottom
	div		ecx
	cmp		eax,	2
	jg		@F
		mov		eax,	2
	@@:
	mov		R1.bottom,	eax
	
; draw visibled rect
	invoke	SetROP2, myDC, R2_NOTXORPEN
	invoke	Rectangle, myDC, R1.left, R1.top, R1.right, R1.bottom
;	invoke	DrawFocusRect, myDC, addr R1
	invoke	SetROP2, myDC, R2_COPYPEN

@@EXIT:

	invoke	SelectObject, myDC, _oB

	ret
DrawMiniWnd endp

ExtendClient2Full MACRO W, H
;  орректировка клиентских размеров W & H до размеров окна Ёскиза
	; w = W + 2*GetSystemMetrics(SM_CXFIXEDFRAME)
	invoke	GetSystemMetrics, SM_CXFIXEDFRAME
	shl		eax,		1
	add		W,			eax	
	
	; h = H + 2*GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYSMCAPTION)
	invoke	GetSystemMetrics, SM_CYFIXEDFRAME
	shl		eax,		1
	add		H,			eax
	invoke	GetSystemMetrics, SM_CYSMCAPTION
	add		H,			eax
ENDM

ResizeMiniWnd proc goto_RightTop_angle : DWORD ; подгон€ть ли под верхний правый угол окна
	LOCAL	R		: RECT
	LOCAL	X, Y,
			W, H	: DWORD
	LOCAL	_hDC	: HDC

	mov		ecx,	SCALE_KOEF

	mov		eax,	Rect.right
	xor		edx,	edx
	div		ecx
	mov		W,			eax
	mov		MiniWidth,	eax
	
	mov		eax,	Rect.bottom
	xor		edx,	edx
	div		ecx
	mov		H,			eax
	mov		MiniHeight,	eax
	
	ExtendClient2Full W, H
	
	.IF (goto_RightTop_angle == 0)
		; если не нужно подгон€ть, то не трогаю X & Y
		invoke	GetWindowRect, hMiniWnd, addr R
		m2m		X,	R.left
		m2m		Y,	R.top
	.ELSE
		; если нужно подгон€ть, то просчитываю X & Y дл€ пересчитанных W & H
		invoke	GetWindowRect, WND, addr R
		mov		eax,	R.right
		sub		eax,	W
		mov		X,		eax
		m2m		Y,		R.top
	
		invoke	GetSystemMetrics, SM_CXFIXEDFRAME
		add		eax,	4
		sub		X,		eax
	
		invoke	GetSystemMetrics, SM_CYFIXEDFRAME
		add		Y,		eax
		invoke	GetSystemMetrics, SM_CYCAPTION
		add		Y,		eax
		add		Y,		4
	.ENDIF
	
	invoke	MoveWindow, hMiniWnd, X, Y, W, H, TRUE
	
; ѕ≈–≈—ќ«ƒјё HBITMAP
	invoke	GetDC,	hMiniWnd
	mov		_hDC,	eax
	
	invoke	DeleteObject, MWndBmp
	
	invoke	CreateCompatibleBitmap, _hDC, MiniWidth, MiniHeight
	mov		MWndBmp,	eax
	
	invoke	SelectObject, MWndDC, eax
	
	invoke	ReleaseDC, hMiniWnd, _hDC
	
	ret
ResizeMiniWnd endp