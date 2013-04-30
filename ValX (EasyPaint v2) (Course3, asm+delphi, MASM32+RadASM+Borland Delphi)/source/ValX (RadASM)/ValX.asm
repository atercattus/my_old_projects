;##################################################################
;    ...::: Valkiria :::...
;
; ‘айл     : ValX.inc
; ќписание : ќсновной модуль программы
; ¬ерси€   : 26.03
;
;##################################################################
.386
.model flat,stdcall
option casemap:none

include \masm32\include\windows.inc
include \masm32\include\user32.inc
include \masm32\include\kernel32.inc
include \masm32\include\gdi32.inc
include \masm32\include\comctl32.inc
include \masm32\include\comdlg32.inc
include \masm32\include\masm32.inc

includelib \masm32\lib\user32.lib
includelib \masm32\lib\kernel32.lib
includelib \masm32\lib\gdi32.lib
includelib \masm32\lib\comctl32.lib
includelib \masm32\lib\masm32.lib
includelib \masm32\lib\comdlg32.lib

include \masm32\macros\macros.asm

.const
	cWndStyle	equ WS_OVERLAPPEDWINDOW or WS_CLIPSIBLINGS or WS_CLIPCHILDREN
	cClassStyle	equ CS_OWNDC
	; начальные размеры клиентской области и изображени€
	cWidth		equ 520
	cHeight		equ 500
	
IDM_FILE				equ 10000
	IDM_NEW				equ 10001
	IDM_OPEN			equ 10002
	IDM_SAVE			equ 10003
	IDM_EXIT			equ 10005

IDM_VIEW				equ 10100
	IDM_WND_INSTRUMS	equ 10101
	IDM_WND_STATUSBAR	equ 10102
	IDM_WND_PARAMS		equ 10103
	IDM_WND_PAL			equ 10104
	IDM_NET				equ 10105
	IDM_WND_THUMBNAIL	equ 10106
	
IDM_IMAGE				equ 10200
	IDM_CLEAR			equ 10201
	IDM_INVERT			equ 10202
	IDM_MATRIX			equ 10203
	IDM_PROCESS			equ 10204
		IDM_COLORFILT	equ 10205
		IDM_THUMBNAIL	equ 10206
		IDM_BROWSER		equ 10205

IDM_INSTRUMENTS			equ 10301
	IDM_PEN				equ 10302
	IDM_LINE			equ 10303
	IDM_RECT			equ 10304
	IDM_ELLIPCE			equ 10305
	IDM_ERASE			equ 10306
	IDM_FILL			equ 10307
	IDM_GETTER			equ 10308

IDM_PARAMS				equ 10401
	IDM_SELFG			equ 10402
	IDM_SELBG			equ 10403

IDM_HELP				equ 10501
	IDM_ABOUT			equ 10502
	
	; приращени€ при скролинге
	SCROLL_MINI	equ	2
	SCROLL_FAST	equ	SCROLL_MINI*10
	
.data
	hInstance	HINSTANCE	0
	; дескриптор главного окна
	WND			HWND		0
	; дескриптор окошка Ёскиз
	hMiniWnd	HWND		0
	; дескриптор строки состо€ни€
	hStatusBar	DWORD	0
	
	ClassName	db "ValX_Application",0
	AppName		db "Valkiria alpha (build 26.03.06 12:34)",0
	MenuName	db "MAINMENU",0
	
	; главное меню 
	hMenu		HMENU		0
	; стандартный курсор на врем€ рисовани€
	StdCursor	HCURSOR		0
	; курсор на врем€ таскани€
	DragCursor	HCURSOR		0
	
	; отступ сверху в клиентской области окна от начала BackClient
	; складываетс€ из высоты ToolBar и высоты ParamBar
	TOP_DELTA	DWORD		0
	
	; текущий примитив и выбранный пункт меню
	OutMode		DWORD	IDM_PEN
	; ранее выбранный пункт (дл€ пипетки)
	PrevMode	DWORD	IDM_PEN
	; предыдуща€ точка вывода временного примитива (в координатах клиентской области)
	OldPos		POINT	<0, 0>
	; начальна€ точка примитива (в координатах клиентской области)
	StartPos	POINT	<0, 0>
	; флаг выводимости ранее временного примитива
	ExistsPrev	BYTE	FALSE
	; флаг о произошедшем изменении размеров окна
	HasBeenResized	BYTE	FALSE

	; какой сейчас режим:
	MODE_DRAW		equ		1		; рисование
	MODE_DRAG		equ		0FFh	; таскание
	MODE_NOTH		equ		0		; ничего-не-деланье
	WHAT_MODE		BYTE	MODE_NOTH
	
	; цвет карандаша
	FGColor		DWORD	0
	; цвет кисти
	BGColor		DWORD	0FFFFFFh
	
.code

include Interface\ParamBar.inc
include TripleBuff\TripleBuff.inc
include Interface\StatusBar.inc
include Interface\ToolBar.inc
include MiniWnd.inc
include ParamDlg.inc
include Mouse.inc
include Menu.inc
include procWinMain.inc

WinMain proto

start:
	invoke	GetModuleHandle, NULL
	mov		hInstance, eax
	
	invoke InitCommonControls
	
	invoke	WinMain
	invoke	ExitProcess, eax

SCROLL_X MACRO d
	invoke	ScrollX, d
	invoke	CalcRgns
	DRAW_AND_FLIP
	invoke	THUMBNAIL_IMAGE
ENDM

SCROLL_Y MACRO d
	invoke	ScrollY, d
	invoke	CalcRgns
	DRAW_AND_FLIP
	invoke	THUMBNAIL_IMAGE
ENDM

MY_COORD_CHECK macro reg
	push	ebx
	
	mov		ebx,	reg
	cmp		bx,		0
	jnl		@F
		sub		ebx,	0FFFFh
	@@:
	mov		reg,	ebx
	
	pop		ebx
ENDM

WndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
	LOCAL	PS : PAINTSTRUCT
	LOCAL	X:DWORD, Y:DWORD

	; ####################### WM_MOUSEMOVE #######################
	.IF (uMsg==WM_MOUSEMOVE)
		.IF (WHAT_MODE == MODE_DRAG)
			invoke	SetCursor, DragCursor ; вот такой изврат в Win32API, чтобы мен€ть курсоры :(
			
			push	ebx
			
			mov		eax,	lParam
			mov		ecx,	eax
			and		eax,	0FFFFh
			shr		ecx,	16
			
			MY_COORD_CHECK	eax
			MY_COORD_CHECK	ecx
			mov		edx,	eax
			mov		ebx,	ecx
			
			sub		eax,	OldPos.x
			sub		ecx,	OldPos.y
			
			mov		OldPos.x,	edx
			mov		OldPos.y,	ebx
			
			push	ecx
			invoke	ScrollX, eax
			pop		ecx
			invoke	ScrollY, ecx
			
			invoke	CalcRgns
			DRAW_AND_FLIP
			invoke	THUMBNAIL_IMAGE

			pop		ebx
			
		.ELSEIF (WHAT_MODE == MODE_DRAW)
			invoke	SetCursor, StdCursor
			
			mov		eax,	lParam
			mov		ecx,	eax
			and		eax,	0FFFFh
			shr		ecx,	16
			
			invoke	MouseMove, eax, ecx, wParam
		.ELSE
			invoke	SetCursor, StdCursor
		.ENDIF
		
	; ####################### WM_PAINT #######################
	.ELSEIF (uMsg==WM_PAINT)
		invoke	BeginPaint, hWnd, addr PS
		mov		hDC, eax
		DRAW_AND_FLIP
		mov		ExistsPrev,	FALSE
		invoke	EndPaint, hWnd, addr PS
		
	; ####################### WM_LBUTTONUP #######################
	.ELSEIF (uMsg==WM_LBUTTONUP)
		invoke	SetCursor, StdCursor
		invoke	ReleaseCapture
		
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		invoke	MouseLUp, eax, ecx

		mov		WHAT_MODE,	MODE_NOTH

	; ####################### WM_LBUTTONDOWN #######################
	.ELSEIF (uMsg==WM_LBUTTONDOWN)
		invoke	SetCapture, hWnd

		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		push	eax ; save X
		push	ecx ; save Y

		invoke	GetKeyState, VK_CONTROL
		and		ax,	8000h
		.IF (ax==0)
			mov		WHAT_MODE,	MODE_DRAW
			invoke	SetCursor, StdCursor
			pop		ecx ; rest Y
			pop		eax ; rest X
			
			invoke	MouseLDown, eax, ecx
		.ELSE
			mov		WHAT_MODE,	MODE_DRAG
			invoke	SetCursor, DragCursor
			pop		ecx ; rest Y
			pop		eax ; rest X
			mov		OldPos.y,	ecx
			mov		OldPos.x,	eax
		.ENDIF

	; ####################### WM_RBUTTONUP #######################
	.ELSEIF (uMsg==WM_RBUTTONUP)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		invoke	MouseRUp, eax, ecx
		
	; ####################### WM_RBUTTONDOWN #######################
	.ELSEIF (uMsg==WM_RBUTTONDOWN)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		invoke	MouseRDown, eax, ecx
		
	; ####################### WM_KEYDOWN #######################
	.ELSEIF (uMsg==WM_KEYDOWN)
		mov		ecx,	wParam
		
		.IF (ecx==VK_ADD)
			mov		eax,	CurPenSize
			cmp		eax,	MaxPenSize
			jnb		@F
				inc		eax
				invoke	SetPenParam,	FGColor,	eax
			@@:
			
		.ELSEIF (ecx==VK_SUBTRACT)
			mov		eax,	CurPenSize
			cmp		eax,	1
			jna		@F
				dec		eax
				invoke	SetPenParam,	FGColor,	eax
			@@:
			
		.ELSEIF (ecx==VK_UP)
			SCROLL_Y	SCROLL_MINI
			
		.ELSEIF(ecx==VK_DOWN)
			SCROLL_Y	-SCROLL_MINI
			
		.ELSEIF (ecx==VK_LEFT)
			SCROLL_X	SCROLL_MINI
			
		.ELSEIF(ecx==VK_RIGHT)
			SCROLL_X	-SCROLL_MINI
			
		.ENDIF

	; ####################### WM_KEYUP #######################
	.ELSEIF (uMsg==WM_KEYUP)
	
		.IF (wParam==VK_CONTROL)
			cmp		WHAT_MODE,	MODE_NOTH
			jne		@F
				mov		WHAT_MODE,	MODE_NOTH
			@@:
		.ENDIF

	; ####################### WM_DRAWITEM #######################
	.ELSEIF (uMsg==WM_DRAWITEM)
		.IF (wParam==StatusBarID)
			invoke	DrawItem, lParam
		.ENDIF
		
	; ####################### WM_MOUSEWHEEL #######################
	.ELSEIF (uMsg==WM_MOUSEWHEEL)
		mov		eax,	wParam
		shr		eax,	16
		
		mov		ecx,	SCROLL_FAST
		
		cmp		eax,	65000
		jb		@F
			neg		ecx
		@@:
		
		push	ecx
		invoke	GetKeyState, VK_SHIFT
		pop		ecx
		and		ax,	8000h
		.IF (ax == 0)
			invoke	ScrollY, ecx
		.ELSE
			invoke	ScrollX, ecx
		.ENDIF
		
		invoke	CalcRgns
		DRAW_AND_FLIP
		invoke	THUMBNAIL_IMAGE
		
	; ####################### WM_SIZE #######################
	.ELSEIF (uMsg==WM_SIZE)
		invoke	SendMessage, hToolBar, TB_AUTOSIZE, 0, 0
		invoke	SendMessage, hStatusBar, WM_SIZE, wParam, lParam
		invoke	Resize
		invoke	ScrollX, 0
		invoke	ScrollY, 0
		invoke	CalcRgns
		DRAW_AND_FLIP
		invoke	THUMBNAIL_IMAGE
		mov		HasBeenResized,	TRUE
		
	; ####################### WM_COMMAND #######################
	.ELSEIF (uMsg==WM_COMMAND)
		invoke	MenuProcess, wParam

	; ####################### WM_GETMINMAXINFO #######################
	.ELSEIF (uMsg==WM_GETMINMAXINFO)
		mov		eax,	lParam
		; ”станавливаю минимальный размер окна
		mov		(MINMAXINFO ptr [eax]).ptMinTrackSize.x,	520
		mov		(MINMAXINFO ptr [eax]).ptMinTrackSize.y,	118

	.ELSEIF (uMsg==WM_SHOWWINDOW)
		mov		HasBeenResized,	FALSE

	; ####################### WM_CLOSE or WM_DESTROY #######################
	.ELSEIF (uMsg==WM_CLOSE || uMsg==WM_DESTROY)
		invoke	MessageBox, hWnd, SADD("¬ыйти?"), SADD("ѕодтверждение"), MB_OKCANCEL or MB_ICONQUESTION
		.IF (eax == IDOK)
			invoke	PostQuitMessage, 0
		.ENDIF

	; ####################### ELSE #######################
	.ELSE
		invoke DefWindowProc, hWnd, uMsg, wParam, lParam
		ret
	.ENDIF

	xor 	eax, eax
	ret
WndProc endp

end start