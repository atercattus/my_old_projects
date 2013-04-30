.const
	MaxPenSize equ 100
	
.code

SetPenParam proc Color:DWORD, W:DWORD
	LOCAL num[20]:BYTE

	invoke	CreatePen, PS_SOLID, W, Color
	invoke	SelectObject, hCompDC, eax
	invoke	DeleteObject, eax
	
	; мен€ю текст в строке состо€ни€
	szText frmt, " “олщина линии: %lu"
	invoke	wsprintf, addr num, addr frmt, W
	invoke	SendMessage, hStatusBar, SB_SETTEXT, 1, addr num
	
	; сохран€ю
	m2m		FGColor,	Color
	
	m2m		CurPenSize,	W
	
	ret
SetPenParam endp

SetBrushColor proc Color:DWORD
LOCAL	LB : LOGBRUSH

	mov		LB.lbStyle,		BS_SOLID
	m2m		LB.lbColor,		Color	

	invoke	CreateBrushIndirect, addr LB
	invoke	SelectObject, hCompDC, eax
	invoke	DeleteObject, eax
	
	; сохран€ю
	m2m		BGColor,	Color
	
	ret
SetBrushColor endp

PutPixel proc X:DWORD, Y:DWORD, Color:DWORD
	LOCAL	pc, bc : DWORD

	mov		eax,	CurPenSize
	.IF (eax==1)
		invoke	SetPixel, hCompDC, X, Y, Color
	.ELSE
	
		m2m		pc,	FGColor
		m2m		bc,	BGColor
		invoke	SetPenParam, Color, CurPenSize
		invoke	SetBrushColor, Color
	
		mov		eax,	CurPenSize
		shr		eax,	1
		mov		ecx,	X
		mov		edx,	Y
		add		ecx,	eax
		add		edx,	eax
		
		push	edx
		push	ecx
		
		neg		eax
		lea		ecx,	[ecx+2*eax]
		lea		edx,	[edx+2*eax]
		
		push	edx
		push	ecx
		
		push	hCompDC
		
		call	Ellipse
		
		invoke	SetPenParam, pc, CurPenSize
		invoke	SetBrushColor, bc
	.ENDIF
	
	ret
PutPixel endp