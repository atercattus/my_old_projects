.const
	ToolBarID	equ	100

.data
	hToolBar	DWORD	0
	image		HBITMAP	0
	TBHeight	DWORD	0
	
.data?
	tbb		TBBUTTON	<>

.code

NewButton proc bID:DWORD, cID:DWORD
	m2m tbb.iBitmap,	bID
	m2m tbb.idCommand,	cID
	mov tbb.fsStyle,	TBSTYLE_BUTTON	
	
	invoke SendMessage, hToolBar, TB_ADDBUTTONS, 1, addr tbb
	
	ret
NewButton endp

NewCheck proc bID:DWORD, cID:DWORD
	m2m	tbb.iBitmap,	bID
	m2m tbb.idCommand,	cID
	mov tbb.fsStyle,	TBSTYLE_GROUP or TBSTYLE_CHECKGROUP
	invoke SendMessage, hToolBar, TB_ADDBUTTONS, 1, addr tbb
	
	ret
NewCheck endp

Blank proc
	mov tbb.iBitmap,	-1
	mov tbb.idCommand,	0
	mov tbb.fsStyle,	TBSTYLE_SEP
	invoke SendMessage, hToolBar, TB_ADDBUTTONS, 1, addr tbb
	
	ret
Blank endp

SetBmpBGColor proto

InitToolBar proc
	LOCAL rect	: RECT
	LOCAL tbab	: TBADDBITMAP
	
	invoke	LoadBitmap, hInstance, 100
	cmp		eax,	NULL
	jne		@F
		invoke	MessageBox, 0, SADD("Ошибка в структуре программы!"), SADD("Ошибка"), MB_OK or MB_ICONERROR or MB_SYSTEMMODAL
		invoke	ExitProcess, 0
@@:
	mov		image,	eax
	
	invoke CreateToolbarEx,	WND, \
							WS_CHILD or WS_CLIPSIBLINGS or TBSTYLE_TOOLTIPS, \
                            ToolBarID, \
                            0, \
                            NULL, \
                            NULL, \
                            NULL, \
                            0, \
                            32, \
                            32, \
                            32, \
                            32, \
                            sizeof TBBUTTON
                            
	mov	hToolBar, eax

	invoke	SetBmpBGColor
	
	mov tbab.hInst,	0
	m2m tbab.nID,	image
	invoke SendMessage, hToolBar, TB_ADDBITMAP, 12, addr tbab

	mov	tbb.fsState,	TBSTATE_ENABLED
	mov tbb.dwData,		0
	mov tbb.iString,	0

	; new open save | pen line rect ellipse erase fill get | help | exit
    invoke	Blank
    invoke	NewButton, 0, IDM_NEW
    invoke	NewButton, 1, IDM_OPEN
    invoke	NewButton, 2, IDM_SAVE
    invoke	Blank
    invoke	NewCheck, 3, IDM_PEN
    invoke	NewCheck, 4, IDM_LINE
    invoke	NewCheck, 5, IDM_RECT
    invoke	NewCheck, 6, IDM_ELLIPCE
    invoke	NewCheck, 7, IDM_ERASE
    invoke	NewCheck, 8, IDM_FILL
    invoke	NewCheck, 9, IDM_GETTER
    invoke	Blank
    invoke	NewButton, 10, IDM_ABOUT
    invoke	Blank
    invoke	NewButton, 11, IDM_EXIT

    
    ; нажимаю на кнопку Карандаш
    invoke	SendMessage, hToolBar, TB_CHECKBUTTON, IDM_PEN, 1
    
	invoke	ShowWindow, hToolBar, SW_SHOW
	
	invoke	GetWindowRect, hToolBar, addr rect
	
	mov		TBHeight,	44

	ret
InitToolBar endp

SetBmpBGColor proc
	LOCAL	mDC			: HDC
	LOCAL	hBrush		: HBRUSH
	LOCAL	hOldBmp		: HBITMAP
	LOCAL	hOldBrush	: HBRUSH

	invoke	CreateCompatibleDC, NULL
	mov		mDC, eax

	invoke	SelectObject, mDC, image
	mov		hOldBmp, eax

	invoke	GetSysColor, COLOR_BTNFACE
	invoke	CreateSolidBrush, eax
	mov		hBrush, eax

	invoke	SelectObject, mDC, hBrush
	mov		hOldBrush, eax

	invoke	GetPixel, mDC, 1, 1
	invoke	ExtFloodFill, mDC, 1, 1, eax, FLOODFILLSURFACE

	invoke	SelectObject, mDC, hOldBrush
	invoke	DeleteObject, hBrush

	invoke	SelectObject, mDC, image
	invoke	DeleteDC, mDC

	ret
SetBmpBGColor endp