;##################################################################
;    ...::: Valkiria :::...
;
; ‘айл     : initApp.inc
; ќписание : »нициализирующие действи€
; ¬ерси€   : 23.03
;
;##################################################################
	; создаю буфер изображени€
	invoke	CreateBackBuffer, WND, cWidth, cHeight, FALSE
	.IF (!eax)
		invoke	MessageBox, WND, SADD("ќшибка при создании буфера цвета!"), SADD("ќшибка"), MB_OK or MB_ICONERROR
		invoke	PostQuitMessage, NULL
		invoke	ExitProcess, 0
	.ENDIF
	
	; создаю буфер клиентской области
	invoke	CreateClientBackBuff, WND
	
	; выставл€ю начальные параметры
	invoke	SetPenParam, FGColor, 1
	invoke	SetBrushColor, BGColor
	
	; выбираю самый быстрый(вроде как) режим работы StretchBlt
	invoke	SetStretchBltMode, hDC, 2
	
	; создаю окно Ёскиза
	invoke	CreateMiniWnd