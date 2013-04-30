WinMain proc hInst:HINSTANCE
	LOCAL	wc		:WNDCLASSEX
	LOCAL	msg		:MSG
	LOCAL	hAccel	:HACCEL
	LOCAL	w, h,
	        x, y	:DWORD
	LOCAL	R		:RECT
	
	invoke	RtlZeroMemory, addr wc, sizeof WNDCLASSEX
	mov		wc.cbSize, sizeof WNDCLASSEX
	mov		wc.style, cClassStyle
	mov		wc.lpfnWndProc, offset WndProc
	push	hInstance
	pop		wc.hInstance
	invoke	GetStockObject, BLACK_BRUSH
	mov		wc.hbrBackground, eax
	mov		wc.lpszClassName, offset ClassName
	invoke	LoadIcon, hInstance, 102
	mov		wc.hIcon, eax
	mov		wc.hIconSm, eax
	
	invoke	LoadCursor, hInstance, 101
	cmp		eax,	NULL
	jne		@F
		invoke	MessageBox, 0, SADD("Ошибка в структуре программы!"), SADD("Ошибка"), MB_OK or MB_ICONERROR or MB_SYSTEMMODAL
		invoke	ExitProcess, 0
@@:
	mov		wc.hCursor, eax
	
	; РЕГИСТРИРУЮ КЛАСС
	invoke	RegisterClassEx, addr wc
	
	; ВЫЧИСЛЯЮ НУЖНЫЙ РАЗМЕР И ПОЛОЖЕНИЕ НА ЭКРАНЕ
	; w = сWidth + 2*GetSystemMetrics(SM_CXFIXEDFRAME)
	invoke	GetSystemMetrics, SM_CXFIXEDFRAME
	shl		eax,	1
	add		eax,	cWidth
	mov		w,		eax
	; h = сHeight + 2*GetSystemMetrics(SM_CYFIXEDFRAME) + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENUSIZE) +
	;{в идеале} + ToolBar.Height+StatusBar.Height/*но они еще не определены :(, поэтому добавляю позже*/
	invoke	GetSystemMetrics, SM_CYFIXEDFRAME
	shl		eax,	1
	add		eax,	cHeight
	mov		h,		eax
	invoke	GetSystemMetrics, SM_CYCAPTION
	add		h,		eax
	invoke	GetSystemMetrics, SM_CYMENUSIZE
	add		h,		eax
	; x = (GetSystemMetrics(SM_CXSCREEN)-w) >> 1
	invoke	GetSystemMetrics, SM_CXSCREEN
	sub		eax,	w
	shr		eax,	1
	mov		x,		eax
	; y = (GetSystemMetrics(SM_CYSCREEN)-h) >> 1
	invoke	GetSystemMetrics, SM_CYSCREEN
	sub		eax,	h
	shr		eax,	1
	mov		y,		eax
	
	invoke	LoadMenu, hInstance, offset MenuName
	mov		hMenu,	eax
	
	; СОЗДАЮ ОКНО
	invoke	CreateWindowEx,\
				NULL,\
				addr ClassName,\
				addr AppName,\
				cWndStyle,\
				x,\
				y,\
				w,\
				h,\
				NULL,\
				hMenu,\
				hInst,\
				NULL
	mov		WND,	eax

	invoke	InitToolBar
	invoke	InitStatusBar
	
	; пересчитываю высоту окна с учетом панели инструментов и строки состояния
	invoke	GetWindowRect, WND, addr R
	
	mov		eax,	R.bottom
	sub		eax,	R.top
	add		eax,	TBHeight
	add		eax,	SBHeight
	
	mov		ecx,	R.right
	sub		ecx,	R.left
	
	; и изменяю размер окна
	invoke	MoveWindow, WND, R.left, R.top, ecx, eax, FALSE
	
include InitApp.inc

	invoke	LoadAccelerators, hInstance, SADD("ACCEL")
	mov		hAccel,	eax
	
	; ПОКАЗЫВАЮ ОКНО
	invoke	ShowWindow, WND, SW_SHOW
	invoke	UpdateWindow, WND
	
	; ЗАПУСКАЮ ОБРАБОТКУ СООБЩЕНИЙ
	
@@MainLoop:
	invoke	GetMessage, addr msg, NULL, 0, 0
	cmp		eax,	0
	je @@BreakMainLoop
	
	.IF (hDlg)
		invoke	IsDialogMessage, hDlg, addr msg
		.IF (eax)
			jmp short @@MainLoop
		.ENDIF
	.ENDIF

	invoke	TranslateAccelerator, WND, hAccel, addr msg
	cmp		eax,	0
	jne @@MainLoop
	
	invoke	TranslateMessage, addr msg

	invoke	DispatchMessage, addr msg
	
	jmp short @@MainLoop
	
@@BreakMainLoop:

	invoke	DeleteBackBuffer
		
	mov		eax, 0
	ret
WinMain endp