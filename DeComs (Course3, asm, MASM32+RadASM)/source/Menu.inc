MenuProcess proc wParam:DWORD
; # Обработка сообщений меню.

	mov		eax,	wParam
	; ####################### IDM_ABOUT #########################
	.IF (ax==IDM_ABOUT)
		.data
			about_text db 'Удалялка комментариев из файлов каталога',0
		.code
		invoke	MessageBox, WND, addr about_text, SADD("О программе"), MB_OK or MB_ICONINFORMATION
	

	; ####################### IDM_PARAMS ########################
	.ELSEIF (ax==IDM_PARAMS)
		invoke	ReadINIFile


	; ####################### IDM_START #########################
	.ELSEIF (ax==IDM_START)
		invoke	FileProcess, StartPath, addr format


	; ####################### IDM_OPTIONS #######################
	.ELSEIF (ax==IDM_OPTIONS)
		invoke	ResultDlgShow


	; ####################### IDM_EXIT ##########################
	.ELSEIF (ax==IDM_EXIT)
		invoke	SendMessage, WND, WM_CLOSE, 0, 0
	

	.ENDIF

	ret
MenuProcess endp