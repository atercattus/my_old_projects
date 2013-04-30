.const
	APP_CLASS_NAME	db "AC.ACVHDD_WinClass",0
.code

WndProc proto :HWND, :UINT, :WPARAM, :LPARAM

WinMain proc hInst:HINSTANCE, hPrevInst:HINSTANCE, CmdLine:LPSTR, CmdShow:DWORD
LOCAL	wc		: WNDCLASSEX
LOCAL	msg		: MSG
LOCAL	hWnd	: DWORD

	invoke	RtlZeroMemory, addr wc, sizeof(WNDCLASSEX)
	mov		wc.cbSize,			sizeof(WNDCLASSEX)
	mov		wc.lpfnWndProc,		offset WndProc
	m2m		wc.hInstance,		hInst
	mov		wc.lpszClassName,	offset APP_CLASS_NAME
	
	invoke	RegisterClassEx, addr wc
	
	invoke	CreateWindowEx,	WS_EX_TOPMOST, addr APP_CLASS_NAME, NULL, 0, CW_USEDEFAULT, CW_USEDEFAULT,
							CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInst, NULL

	.while TRUE
		invoke	GetMessage, addr msg, NULL, 0, 0
		.BREAK .IF (!eax)
		invoke	DispatchMessage, addr msg
	.endw
	
	mov	eax, msg.wParam
	ret
WinMain endp

WndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM

	.if uMsg==WM_CREATE
		; создаю иконку в трее
		m2m		hMainWnd, hWnd
		invoke	AC_TrayIconCreate
		
		; ищу загруженный драйвер
		invoke	AC_FindDriverInMem
		.if (eax == FIND_DRIVER_RUNNING)
			; драйвер загружен. нужно получить букву диска, на которой он "висит"
			invoke	AC_GetUsedDriveLetter
			.IF (eax != 0)
				; найден диск, на котором "висит" драйвер
				; сохран€ю букву диска
				mov		byte ptr [DiskLetter],	al
			.ELSE
				; привод не найден. какой-то глюк в системе: драйвер работает, а диска нет!
				invoke	MessageBox, hMainWnd, $CTA0('—бой в системе'), addr APPLICATION_NAME, MB_OK or MB_ICONERROR
				; может быть стоит попытатьс€ выгрузить глючный драйвер и промолчать об ошибке?
				; тогда все будет выгл€деть так, словно драйвер просто не бы еще загружен.
				; да это и логично, т.к. диска то нет. и пользователь никак не видит, что драйвер работает
			.ENDIF 
		.else
			; остальные состо€ни€ мне не важны - все равно диск не загружен
		.endif
		
		; мен€ю иконку в текущее состо€ние драйвера
		invoke	AC_TrayIconMenuItemsCheckState
		
	.elseif uMsg==WM_DESTROY
		invoke	AC_TrayIconDelete
		invoke	PostQuitMessage, NULL
		
	.elseif uMsg==WM_COMMAND
		invoke	AC_TrayIconWndProc_WM_COMMAND, wParam, lParam
		
	.elseif uMsg==WM_TRAY
		invoke	AC_TrayIconWndProc_WM_TRAY, wParam, lParam
		
	.else
		invoke	DefWindowProc, hWnd, uMsg, wParam, lParam
		ret
	.endif
	
	xor	eax,eax
	ret
WndProc endp