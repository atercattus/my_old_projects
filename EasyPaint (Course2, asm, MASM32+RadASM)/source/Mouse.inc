Ymod MACRO reg, val
	mov		reg,	val
	sub		reg,	TBHeight
ENDM

ROP MACRO
	.IF (FGColor==0FFFFFFh)
		push	R2_XORPEN
	.ELSE
		push	R2_NOTXORPEN
	.ENDIF

	push	hCompDC
	call	SetROP2
ENDM

MouseMove proc X:DWORD, Y:DWORD, K:DWORD

	; нажата правая кнопка мыши?
	mov		eax,	K
	and		eax,	MK_RBUTTON
	cmp		eax,	0
	je		@F
	.IF (OutMode==IDM_GETTER)
		; пипетка фона
		Ymod eax, Y
		invoke	GetPixel, hCompDC, X, eax
		mov		BGColor,	eax
		invoke	InvalidateRect, hStatusBar, NULL, TRUE
	.ENDIF
@@:

	; нажата левая кнопка мыши?
	mov		eax,	K
	and		eax,	MK_LBUTTON
	cmp		eax,	0
	jne		@F
		ret
@@:

	cmp		HasBeenResized,	TRUE
	jne		@F
		ret
@@:	
	
	mov		eax,	OutMode
	.IF (eax<=IDM_INSTRUMENTS || eax>=IDM_CLEAR)
		ret
	.ENDIF
	
	;##############################################################################
	.IF (eax==IDM_PEN)
		Ymod eax, Y
		invoke	PutPixel, X, eax, FGColor
		invoke	Flip
		
	;##############################################################################
	.ELSEIF (eax==IDM_ERASE)
		Ymod eax, Y
		invoke	PutPixel, X, eax, BGColor
		invoke	Flip
		
	;##############################################################################
	.ELSEIF (eax==IDM_GETTER)
		Ymod eax, Y
		invoke	GetPixel, hCompDC, X, eax
		mov		FGColor,	eax
		invoke	SetPenParam, eax, CurPenSize
		invoke	InvalidateRect, hStatusBar, NULL, TRUE
		
	;##############################################################################
	.ELSEIF (eax==IDM_LINE)
		ROP

		cmp		ExistsPrev,	FALSE
		je		@F
			Ymod eax, StartPos.y
			invoke	MoveToEx, hCompDC, StartPos.x, eax, NULL
			Ymod eax, OldPos.y
			invoke	LineTo, hCompDC, OldPos.x, eax
		@@:
	
		Ymod eax, StartPos.y
		invoke	MoveToEx, hCompDC, StartPos.x, eax, NULL
		Ymod eax, Y
		invoke	LineTo, hCompDC, X, eax
		
	;##############################################################################
	.ELSEIF (eax==IDM_RECT)
		ROP

		cmp		ExistsPrev,	FALSE
		je		@F
			Ymod eax, StartPos.y
			Ymod ecx, OldPos.y
			invoke	Rectangle, hCompDC, StartPos.x, eax, OldPos.x, ecx
		@@:
	
		Ymod eax, StartPos.y
		Ymod ecx, Y
		invoke	Rectangle, hCompDC, StartPos.x, eax, X, ecx
		
	;##############################################################################
	.ELSEIF (eax==IDM_ELLIPCE)
		ROP

		cmp		ExistsPrev,	FALSE
		je		@F
			Ymod eax, StartPos.y
			Ymod ecx, OldPos.y
			invoke	Ellipse, hCompDC, StartPos.x, eax, OldPos.x, ecx
		@@:
	
		Ymod eax, StartPos.y
		Ymod ecx, Y
		invoke	Ellipse, hCompDC, StartPos.x, eax, X, ecx
		
	.ENDIF
	
	m2m		OldPos.x,	X
	m2m		OldPos.y,	Y
	mov		ExistsPrev,	TRUE
	
	invoke	SetROP2, hCompDC, R2_COPYPEN
	
	invoke	Flip
	
	ret
MouseMove endp

MenuChecked proto ID:DWORD

MouseLUp proc X:DWORD, Y:DWORD

	cmp		HasBeenResized,	TRUE
	jne		@F
		mov	HasBeenResized,	FALSE
		ret
@@:

	mov		eax,	OutMode
	
	.IF (eax==IDM_LINE)
		Ymod eax, StartPos.y
		invoke	MoveToEx, hCompDC, StartPos.x, eax, NULL
		Ymod eax, Y
		invoke	LineTo, hCompDC, X, eax
		invoke	Flip
		
	.ELSEIF (eax==IDM_RECT)
		Ymod eax, StartPos.y
		Ymod ecx, Y
		invoke	Rectangle, hCompDC, StartPos.x, eax, X, ecx
		invoke	Flip
		
	.ELSEIF (eax==IDM_ELLIPCE)
		Ymod eax, StartPos.y
		Ymod ecx, Y
		invoke	Ellipse, hCompDC, StartPos.x, eax, X, ecx
		invoke	Flip
		
	.ELSEIF (eax==IDM_GETTER)
		; восстанавливаю прежний инструмент
		invoke	MenuChecked,	PrevMode
	.ENDIF
	
	mov		ExistsPrev, FALSE

	ret
MouseLUp endp

MouseLDown proc X:DWORD, Y:DWORD
	LOCAL	bc : DWORD
	LOCAL	R : RECT
	LOCAL	P1 : POINT
	LOCAL	P2 : POINT

	mov		HasBeenResized,	FALSE
	
	mov		eax,	OutMode

	.IF (eax==IDM_PEN)
		Ymod eax, Y
		invoke	PutPixel, X, eax, FGColor
		invoke	Flip
		
	.ELSEIF (eax==IDM_ERASE)
		Ymod eax, Y
		invoke	PutPixel, X, eax, BGColor
		invoke	Flip
		
	.ELSEIF (eax==IDM_GETTER)
		Ymod eax, Y
		invoke	GetPixel, hCompDC, X, eax
		mov		FGColor,	eax
		invoke	InvalidateRect, hStatusBar, NULL, TRUE
		
	.ELSEIF (eax==IDM_FILL)
		Ymod eax, Y
		mov		Y, eax
		
		m2m		bc,	BGColor
		invoke	SetBrushColor, FGColor
		
		invoke	GetPixel, hCompDC, X, Y
		invoke	ExtFloodFill, hCompDC, X, Y, eax, FLOODFILLSURFACE
		invoke	Flip
		
		invoke	SetBrushColor, bc
		
	.ELSE
		m2m		StartPos.x,	X
		m2m		StartPos.y,	Y
		mov		ExistsPrev,	FALSE
	.ENDIF

	ret
MouseLDown endp

MouseRUp proc X:DWORD, Y:DWORD

	cmp		OutMode,	IDM_GETTER
	jne		@F
		; восстанавливаю прежний инструмент
		invoke	MenuChecked,	PrevMode
@@:

	ret
MouseRUp endp

MouseRDown proc X:DWORD, Y:DWORD
	LOCAL	bc : DWORD
	
	cmp		HasBeenResized,	TRUE
	jne		@F
		ret
@@:
	
	.IF (OutMode==IDM_GETTER)
		Ymod eax, Y
		invoke	GetPixel, hCompDC, X, eax
		mov		BGColor,	eax
		invoke	InvalidateRect, hStatusBar, NULL, TRUE
		
	.ELSEIF (OutMode==IDM_FILL)
		Ymod eax, Y
		mov		Y, eax
		
		m2m		bc,	BGColor
		invoke	SetBrushColor, BGColor
		
		invoke	GetPixel, hCompDC, X, Y
		invoke	ExtFloodFill, hCompDC, X, Y, eax, FLOODFILLSURFACE
		invoke	Flip
		
		invoke	SetBrushColor, bc
		
	.ENDIF

	ret
MouseRDown endp