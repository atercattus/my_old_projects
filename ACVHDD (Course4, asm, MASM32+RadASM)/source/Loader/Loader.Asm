.386
.model flat,stdcall
option casemap:none

include \masm32\include\kernel32.inc
include \masm32\include\windows.inc
include \masm32\include\shell32.inc
include \masm32\include\user32.inc
include \masm32\include\advapi32.inc
include \masm32\include\comdlg32.inc
include \masm32\include\gdi32.inc
include \masm32\include\w2k\native.inc
include \masm32\include\w2k\devioctl.inc
include \masm32\include\winioctl.inc
include \masm32\macros\MACROS.ASM
include \masm32\macros\Strings.mac

include \masm32\include\shlwapi.inc
includelib \masm32\lib\shlwapi.lib

include \masm32\include\comctl32.inc
includelib \masm32\lib\comctl32.lib

include \masm32\m32lib\MASM32.INC
includelib \masm32\m32lib\MASM32.lib

includelib \masm32\lib\kernel32.lib
includelib \masm32\lib\shell32.lib
includelib \masm32\lib\user32.lib
includelib \masm32\lib\advapi32.lib
includelib \masm32\lib\comdlg32.lib
includelib \masm32\lib\gdi32.lib

.data?
	hInstance	dd ?
	hMainWnd	dd ?
	
.const
	APPLICATION_NAME	db "AterCattus Virtual HDD loader",0
	
$invoke macro Fun:REQ, A:VARARG
  ifb <A>
    invoke Fun
  else
    invoke Fun, A
  endif
  exitm <eax>
endm

.code
ProcessMessages proc hWnd:DWORD
LOCAL lpMsg : MSG
.WHILE TRUE
	invoke	PeekMessage, addr lpMsg, hWnd, 0, 0, PM_NOREMOVE
	.IF (eax)
		invoke	GetMessage, addr lpMsg, NULL, 0, 0
		.BREAK .IF (!eax)
		invoke	DispatchMessage, addr lpMsg
	.ELSE
		.BREAK
	.ENDIF
.ENDW
	
	ret
ProcessMessages endp

return macro code
	mov		eax,	code
	ret
endm

.data?
	mapped	DWORD ?
	driver_path	db MAX_PATH+1 dup (?)
	k32lib	DWORD ?
	SetFileValidData	DWORD ?

include PasswordDlg.inc
include DriverService.inc
include TrayIcon.inc
include NewImageDlg.inc
include DelImageDlg.inc
include MainWndProc.inc

.code

start:
	; загрузка нужной мне функции 
	invoke	LoadLibrary, $CTA0('kernel32.dll')
	.if (eax==NULL)
@@no_win:
			invoke	MessageBox, 0, $CTA0('Недопустимая версия Windows'), addr APPLICATION_NAME, MB_OK or MB_ICONERROR
			invoke	ExitProcess, TRUE
	.endif
	mov		k32lib,		eax
	invoke	GetProcAddress, k32lib, $CTA0('SetFileValidData')
	.if (eax==NULL)
		invoke	FreeLibrary, k32lib
		jmp		@@no_win
	.endif
	mov		SetFileValidData, eax
	
	; проверка на запуск одной копии
	invoke	CreateFileMapping, -1, NULL, PAGE_READONLY, 0, 32, addr APP_CLASS_NAME
	.IF (eax != 0)
		mov		mapped,	eax
		invoke	GetLastError
		.IF (eax == ERROR_ALREADY_EXISTS)
			invoke	CloseHandle, mapped
			invoke	FreeLibrary, k32lib
			invoke	MessageBox, 0, $CTA0('Программа уже запущена'), addr APPLICATION_NAME, MB_OK or MB_ICONERROR
			invoke	ExitProcess, TRUE
		.ENDIF
	.ENDIF
	
	; получаю абсолютный путь к файлу драйвера в директории запуска
	invoke	GetCurrentDirectory, MAX_PATH, addr driver_path
	invoke	szCatStr, addr driver_path, addr DRIVER_FILE
	
	invoke	GetModuleHandle, NULL
	mov		hInstance, eax
	invoke	InitCommonControls
	invoke	WinMain, hInstance, NULL, NULL, SW_SHOWDEFAULT
	invoke	CloseHandle, mapped
	invoke	FreeLibrary, k32lib
	
	invoke	ExitProcess, eax
end start