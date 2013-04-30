.const
	StatusBarID	equ	101
	
.data
	hStatusBar	DWORD	0
	SBHeight	DWORD	0
	
.code

SetPenParam proto Color:DWORD, W:DWORD

InitStatusBar proc
	LOCAL R : RECT
	LOCAL ws[2]: DWORD
	
	invoke	CreateStatusWindow, WS_CHILD or WS_VISIBLE or SBS_SIZEGRIP, \
								NULL, \
								WND, StatusBarID
                            
	mov	hStatusBar, eax
	
	invoke	GetWindowRect, hStatusBar, addr R
	mov		eax,	R.bottom
	sub		eax,	R.top
	inc		eax
	mov		SBHeight,	eax
	
	mov		[ws + 0],	40
	mov		[ws + 4],	-1
	
	invoke	SendMessage, hStatusBar, SB_SETPARTS, 2, addr ws
	
	invoke	SendMessage, hStatusBar, SB_SETTEXT, 0 or SBT_OWNERDRAW, NULL
	invoke	SendMessage, hStatusBar, SB_SETTEXT, 1 or 0, NULL
	
	invoke	ShowWindow, hStatusBar, SW_SHOW
	
	ret
InitStatusBar endp

DrawItem proc lParam:DWORD
	LOCAL DC : HDC
	LOCAL b, ob : HBRUSH

	mov		ecx,	lParam
	mov		eax,	(DRAWITEMSTRUCT ptr [ecx]).itemID
	.IF (eax==0)
		mov		eax,	(DRAWITEMSTRUCT ptr [ecx]).hdc
		mov		DC,		eax
		
		; FG
		invoke	CreateSolidBrush, FGColor
		mov		b,	eax
		
		invoke	SelectObject, DC, b
		mov		ob,	eax
		
		invoke	Rectangle, DC, 5, 5, 17, 17
		
		invoke	SelectObject, DC, ob
		invoke	DeleteObject, b
		
		; BG
		invoke	CreateSolidBrush, BGColor
		mov		b,	eax
		
		invoke	SelectObject, DC, b
		mov		ob,	eax
		
		invoke	Rectangle, DC, 20, 5, 32, 17
		
		invoke	SelectObject, DC, ob
		invoke	DeleteObject, b
		
	.endif
	
	ret

DrawItem endp