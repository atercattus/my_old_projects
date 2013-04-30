.386
.model flat,stdcall
option casemap:none

include \masm32\include\windows.inc
include \masm32\include\user32.inc
include \masm32\include\kernel32.inc
include \masm32\include\comctl32.inc
include \masm32\include\comdlg32.inc
include \masm32\include\shlwapi.inc
include \masm32\include\masm32.inc

includelib \masm32\lib\user32.lib
includelib \masm32\lib\kernel32.lib
includelib \masm32\lib\comctl32.lib
includelib \masm32\lib\comdlg32.lib
includelib \masm32\lib\shlwapi.lib
includelib \masm32\lib\masm32.lib

include \masm32\macros\macros.asm

WinMain proto

.const
	cWndStyle	equ WS_OVERLAPPEDWINDOW
	cClassStyle	equ CS_HREDRAW or CS_VREDRAW
	; начальные размеры клиентской области
	cWidth		equ 500
	cHeight		equ 500
	
	
	IDM_PARAMS	equ	10002
	IDM_START	equ	10003
	IDM_OPTIONS	equ	10004
	IDM_EXIT	equ	10005
	
	IDM_ABOUT	equ	10102
	
.data
	ClassName		db "XerxProg_" ; это не очепятка, '0' не нужен!!
	AppName			db "DeComs",0
	MenuName		db "MAINMENU",0
	INI_file_name	db "Info.ini",0

.data?
	hInstance	HINSTANCE	?
	WND			HWND		?
	hMenu		HMENU		?
	EXEPath		dd			?
	INIPath		dd			?
	; начальный путь обработки
	StartPath	dd			?
	; временная переменная
	tmp_str		dd			?

.code

include strs.inc
include StatusBarRealization.inc
include EditRealization.inc
include EditFileList.inc
include Filer.inc
include INIFile.inc
include ResultWindow.inc
include Menu.inc
include procWinMain.inc

start:

	invoke	GetModuleHandle, NULL
	mov		hInstance, eax
	
	; формирую полный путь к файлу с настройками
	
	invoke	LocalAlloc, LMEM_FIXED, MAX_PATH+1
	mov		EXEPath,	eax

	invoke	LocalAlloc, LMEM_FIXED, MAX_PATH+1
	mov		INIPath,	eax
	
	invoke	LocalAlloc, LMEM_FIXED, MAX_PATH+1
	mov		tmp_str,	eax

	invoke	LocalAlloc, LMEM_FIXED, MAX_PATH+1
	mov		StartPath,	eax

	invoke	GetCurrentDirectory, MAX_PATH, EXEPath
	invoke	check_on_slash, EXEPath
	invoke	strCopy, INIPath, EXEPath
	invoke	strCat, INIPath, addr INI_file_name
	
	invoke InitCommonControls
	
	invoke	WinMain
	
	invoke	LocalFree, StartPath
	invoke	LocalFree, tmp_str
	invoke	LocalFree, EXEPath
	invoke	LocalFree, INIPath
	
	invoke	ExitProcess, eax

WndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM

	mov		eax, uMsg	; так быстрее и меньше
	
	; ####################### WM_SIZE #######################	
	.IF (eax==WM_SIZE)

		; StatusBar
		invoke	SendMessage, hStatusBar, WM_SIZE, wParam, lParam
		
		;Edit
		mov		eax,	lParam
		xor		ecx,	ecx
		mov		cx,		ax
		shr		eax,	16
		sub		eax,	SBHeight
		inc		eax
		invoke	MoveWindow, hEdit, 0, 0, ecx, eax, TRUE
		
	.ELSEIF (eax==WM_CREATE)
		invoke	TreeView_CreateTemporary, WND
		
	; ####################### WM_COMMAND #######################
	.ELSEIF (eax==WM_COMMAND)
		invoke	MenuProcess, wParam
		
	; ####################### WM_CLOSE or WM_DESTROY #######################
	.ELSEIF (eax==WM_CLOSE || eax==WM_DESTROY)
		invoke	MessageBox, hWnd, SADD("Выйти?"), addr AppName, MB_OKCANCEL or MB_ICONASTERISK
		.IF (eax == IDOK)
			invoke	TreeView_DeleteTemporary
			invoke	WriteINIFile
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