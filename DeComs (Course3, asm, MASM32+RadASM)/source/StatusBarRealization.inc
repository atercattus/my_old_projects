.const
	StatusBarID	equ	101
	sb_def_text db "Файл не открыт",0
	
.data?
	hStatusBar	DWORD	?
	SBHeight	DWORD	?
	
.code

StatusBarSetText proto text:DWORD

InitStatusBar proc
; # Создание статусной строки.

	LOCAL R : RECT
	
	invoke	CreateStatusWindow, WS_CHILD or WS_VISIBLE or SBS_SIZEGRIP, \
								NULL, \
								WND, StatusBarID
	
	mov	hStatusBar, eax
	
	invoke	GetWindowRect, hStatusBar, addr R
	mov		eax,	R.bottom
	sub		eax,	R.top
	inc		eax
	mov		SBHeight,	eax

	invoke	StatusBarSetText, addr sb_def_text
	
	invoke	ShowWindow, hStatusBar, SW_SHOW
	
	ret
InitStatusBar endp

StatusBarSetText proc text:DWORD
; # Изменение текста строки статуса.

	invoke	SendMessage, hStatusBar, SB_SETTEXT, 0, text
	
	ret
StatusBarSetText endp