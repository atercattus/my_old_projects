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

WinMain proto :DWORD
MenuChecked proto :DWORD

.const
	cWndStyle	equ WS_OVERLAPPEDWINDOW
	cClassStyle	equ CS_OWNDC
	; начальные размеры клиентской области и изображения
	cWidth		equ 520
	cHeight		equ 500
	
	IDM_NEW		equ	10002
	IDM_OPEN	equ	10003
	IDM_SAVE	equ	10004
	IDM_EXIT	equ	10006
	
	IDM_INSTRUMENTS	equ 10101
	IDM_PEN		equ 10102
	IDM_LINE	equ 10103
	IDM_RECT	equ 10104
	IDM_ELLIPCE	equ 10105
	IDM_ERASE	equ 10106
	IDM_FILL	equ 10107
	IDM_GETTER	equ 10108
	IDM_CLEAR	equ 10109
	
	IDM_SELFG	equ 10202
	IDM_SELBG	equ 10203
	
	IDM_ABOUT	equ	10302
	
.data
	hInstance	HINSTANCE	0
	WND			HWND		0
	
	ClassName	db "MistClassName",0
	AppName		db "EasyPaint",0
	MenuName	db "MAINMENU",0
	hMenu		HMENU	0
	
	; текущий примитив и выбранный пункт меню
	OutMode		DWORD	IDM_PEN
	; ранее выбранный пункт (для пипетки)
	PrevMode	DWORD	IDM_PEN
	; предыдущая точка вывода временного примитива
	OldPos		POINT	<0, 0>
	; начальная точка примитива
	StartPos	POINT	<0, 0>
	; флаг выводимости ранее временного примитива
	ExistsPrev	BYTE	FALSE
	; флаг о произошедшем изменении размеров окна
	HasBeenResized	BYTE	FALSE
	
	; цвет карандаша
	FGColor		DWORD	0
	; цвет кисти
	BGColor		DWORD	0FFFFFFh
	
.code

include ToolBarRealization.inc
include StatusBarRealization.inc
include DoubleBuff.inc
include ParamDlg.inc
include Draw.inc
include Mouse.inc
include Menu.inc
include procWinMain.inc

start:
	invoke	GetModuleHandle, NULL
	mov		hInstance, eax
	
	invoke InitCommonControls
	
	invoke	WinMain, hInstance
	invoke	ExitProcess, eax

; защита от выхода координат за пределы "листа"
CoordProtect MACRO
	cmp		eax,	Rect.right
	jb		@F
		mov		eax,	Rect.right
@@:

	mov		edx,	Rect.bottom
	add		edx,	TBHeight
	cmp		ecx,	edx
	jna		@F
		mov		ecx,	edx
@@:

	mov		edx,	TBHeight
	cmp		ecx,	edx
	jnb		@F
		mov		ecx,	edx
@@:

endm

WndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
	LOCAL PS : PAINTSTRUCT

	; ####################### WM_MOUSEMOVE #######################
	.IF (uMsg==WM_MOUSEMOVE)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		CoordProtect
		
		invoke	MouseMove, eax, ecx, wParam
		
	; ####################### WM_LBUTTONUP #######################
	.ELSEIF (uMsg==WM_LBUTTONUP)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		CoordProtect
		
		invoke	MouseLUp, eax, ecx

	; ####################### WM_LBUTTONDOWN #######################
	.ELSEIF (uMsg==WM_LBUTTONDOWN)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		CoordProtect
		
		invoke	MouseLDown, eax, ecx

	; ####################### WM_RBUTTONUP #######################
	.ELSEIF (uMsg==WM_RBUTTONUP)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		CoordProtect
		
		invoke	MouseRUp, eax, ecx
		
	; ####################### WM_RBUTTONDOWN #######################
	.ELSEIF (uMsg==WM_RBUTTONDOWN)
		mov		eax,	lParam
		mov		ecx,	eax
		and		eax,	0FFFFh
		shr		ecx,	16
		
		CoordProtect
		
		invoke	MouseRDown, eax, ecx
		
	; ####################### WM_KEYDOWN #######################
	.ELSEIF (uMsg==WM_KEYDOWN)
		mov		eax,	CurPenSize
		mov		ecx,	wParam
		
		.IF (ecx==VK_ADD)
			cmp		eax,	MaxPenSize
			jnb		@F
				inc		eax
			@@:
		.ELSEIF (ecx==VK_SUBTRACT)
			cmp		eax,	1
			jna		@F
				dec		eax
			@@:
		.ELSE
			jmp		@@end
		.ENDIF

		invoke	SetPenParam,	FGColor,	eax
@@end:
	

; ####################### WM_NCLBUTTONUP #######################
.ELSEIF (uMsg==WM_NCLBUTTONUP)
	invoke	MouseLUp, OldPos.x, OldPos.y
		
	; ####################### WM_PAINT #######################
	.ELSEIF (uMsg==WM_PAINT)
		invoke	BeginPaint, hWnd, addr PS
		mov		hDC, eax
		invoke	Flip
		mov		ExistsPrev,	FALSE
		invoke	EndPaint, hWnd, addr PS
		
	; ####################### WM_DRAWITEM #######################
	.ELSEIF (uMsg==WM_DRAWITEM)
		.IF (wParam==StatusBarID)
			invoke	DrawItem, lParam
		.ENDIF
		
	; ####################### WM_MOUSEWHEEL #######################
	.ELSEIF (uMsg==WM_MOUSEWHEEL)
		mov		eax,	wParam
		shr		eax,	16
		
		mov		ecx,	CurPenSize
		
		.IF (eax<65000)
			.IF (ecx<MaxPenSize)
				inc		ecx
			.ENDIF
		.ELSE
			.IF (ecx>1)
				dec		ecx
			.ENDIF
		.ENDIF
		
		mov		CurPenSize,	ecx

		invoke	SetPenParam,	FGColor,	ecx
		

	; ####################### WM_SIZE #######################
	.ELSEIF (uMsg==WM_SIZE)
		invoke	SendMessage, hToolBar, TB_AUTOSIZE, 0, 0
		invoke	SendMessage, hStatusBar, WM_SIZE, wParam, lParam
		invoke	Resize
		invoke	Flip
		mov		HasBeenResized,	TRUE	
		
	; ####################### WM_COMMAND #######################
	.ELSEIF (uMsg==WM_COMMAND)
		invoke	MenuProcess, wParam
		
	.ELSEIF (uMsg==WM_SHOWWINDOW)
		mov		HasBeenResized,	FALSE

	; ####################### WM_CLOSE or WM_DESTROY #######################
	.ELSEIF (uMsg==WM_CLOSE || uMsg==WM_DESTROY)
		invoke	MessageBox, hWnd, SADD("Выйти?"), SADD("Подтверждение"), MB_OKCANCEL or MB_ICONQUESTION
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