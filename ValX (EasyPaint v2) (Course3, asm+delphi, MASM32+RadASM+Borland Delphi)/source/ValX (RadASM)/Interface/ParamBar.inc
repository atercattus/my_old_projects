;##################################################################
;    ...::: Valkiria :::...
;
; Файл     : ParamBar.inc
; Описание : Реализация панели параметров
; Версия   : 23.03
;
;##################################################################
.const
	ParamBarID	equ	102

.data
	hParamBar	DWORD	0
	
.code

;InitParamBar proc
;	LOCAL rect	: RECT
;	LOCAL tbab	: TBADDBITMAP
;	
;	invoke	LoadBitmap, hInstance, 100
;	cmp		eax,	NULL
;	jne		@F
;		invoke	MessageBox, 0, SADD("Ошибка в структуре программы!"), SADD("Ошибка"), MB_OK or MB_ICONERROR or MB_SYSTEMMODAL
;		invoke	ExitProcess, 0
;@@:
;	mov		image,	eax
;	
;	invoke CreateToolbarEx,	WND, \
;							WS_CHILD or TBSTYLE_TOOLTIPS, \
;                            ToolBarID, \
;                            0, \
;                            NULL, \
;                            NULL, \
;                            NULL, \
;                            0, \
;                            32, \
;                            32, \
;                            32, \
;                            32, \
;                            sizeof TBBUTTON
;                            
;	mov	hToolBar, eax
;
;	invoke	SetBmpBGColor
;	
;	mov tbab.hInst,	0
;	m2m tbab.nID,	image
;	invoke SendMessage, hToolBar, TB_ADDBITMAP, 12, addr tbab
;
;	mov	tbb.fsState,	TBSTATE_ENABLED
;	mov tbb.dwData,		0
;	mov tbb.iString,	0
;
;	; new open save | pen line rect ellipse erase fill get | help | exit
;    invoke	Blank
;    invoke	NewButton, 0, IDM_NEW
;    invoke	NewButton, 1, IDM_OPEN
;    invoke	NewButton, 2, IDM_SAVE
;    invoke	Blank
;    invoke	NewCheck, 3, IDM_PEN
;    invoke	NewCheck, 4, IDM_LINE
;    invoke	NewCheck, 5, IDM_RECT
;    invoke	NewCheck, 6, IDM_ELLIPCE
;    invoke	NewCheck, 7, IDM_ERASE
;    invoke	NewCheck, 8, IDM_FILL
;    invoke	NewCheck, 9, IDM_GETTER
;    invoke	Blank
;    invoke	NewButton, 10, IDM_ABOUT
;    invoke	Blank
;    invoke	NewButton, 11, IDM_EXIT
;
;    
;    ; нажимаю на кнопку Карандаш
;    invoke	SendMessage, hToolBar, TB_CHECKBUTTON, IDM_PEN, 1
;    
;	invoke	ShowWindow, hToolBar, SW_SHOW
;	
;	invoke	GetWindowRect, hToolBar, addr rect
;	
;	add		TOP_DELTA,	44
;
;	ret
;InitParamBar endp
