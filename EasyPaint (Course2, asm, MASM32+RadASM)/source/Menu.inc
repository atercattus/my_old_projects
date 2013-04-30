SelectColor proc lpColor:DWORD
	LOCAL CC : CHOOSECOLOR
	LOCAL CustomColors[16] : DWORD
	
	invoke	RtlZeroMemory, addr CustomColors, 4*16

    mov		CC.lStructSize, sizeof CHOOSECOLOR
    push	WND
    pop		CC.hwndOwner
    push	hInstance
    pop		CC.hInstance
    m2m		CC.rgbResult, [lpColor]
    lea		eax, CustomColors
    mov		CC.lpCustColors,   eax
    mov		CC.Flags, CC_FULLOPEN or CC_RGBINIT
    mov		CC.lCustData, 0
    mov		CC.lpfnHook, 0
    m2m		CC.lpTemplateName, NULL

    invoke	ChooseColor, addr CC
	.IF (eax)
		mov		eax,	CC.rgbResult
		mov		ecx,	lpColor
		mov		[ecx],	eax
		mov		eax,	1
		ret
	.ENDIF
	
	xor		eax,	eax
	ret
SelectColor endp

MenuChecked proc ID:DWORD

	invoke	CheckMenuItem, hMenu, OutMode, MF_BYCOMMAND or MF_UNCHECKED
	
	.IF (ID==IDM_GETTER) ; беру пипетку
		; сохраняю старый инструмент
		.IF (OutMode != IDM_GETTER)
			m2m		PrevMode,	OutMode
		.ENDIF
	.ENDIF
	
	mov		eax,	ID
	mov		OutMode,	eax
	
	; нажимаю на соответствующую кнопку
	invoke	SendMessage, hToolBar, TB_CHECKBUTTON, ID, 1
	
	invoke	CheckMenuItem, hMenu, ID, MF_BYCOMMAND
	
	.IF (eax==MF_CHECKED)
		push	MF_BYCOMMAND or MF_UNCHECKED
 	.ELSE
 		push	MF_BYCOMMAND or MF_CHECKED
	.ENDIF
	
	push	ID
	push	hMenu
	call	CheckMenuItem
	
	ret

MenuChecked endp

MenuProcess proc wParam:DWORD
	LOCAL fname : DWORD
	LOCAL buff[128]: BYTE

mov		eax,	wParam
.IF (ax==IDM_ABOUT)
	.data
		text db 'Курсовая работа по курсу "Системное программирование"',13,10
		     db 13,10,"Xerx, 13.05.05",0
	.code
	invoke	MessageBox, WND, addr text, SADD("О программе"), MB_OK or MB_ICONINFORMATION



.ELSEIF (ax>=IDM_PEN && ax<=IDM_GETTER)
	; выбран один из инструментов, нужно изменить начинку меню
	and		eax,	0FFFFh
	invoke	MenuChecked, eax
	mov		ExistsPrev,	FALSE


	
.ELSEIF (ax==IDM_CLEAR)
	invoke	MessageBox, WND, SADD("Очистить изображение?"), SADD("Подтверждение"), MB_OKCANCEL or MB_ICONQUESTION
	.IF (eax==IDOK)
		invoke	Clear
		invoke	Flip
	.ENDIF
	

	
.ELSEIF (ax==IDM_OPEN)
	mov		HasBeenResized,	TRUE
	.data
		path db "BMP-файлы",0,"*.bmp",0,0
	.code

	mov		fname, OpenFileDlg(WND, hInstance, "Загрузка изображения", addr path)
	cmp		BYTE PTR [eax], 0
	je		@F
		invoke	LoadFromFile, fname	
@@:



.ELSEIF (ax==IDM_SAVE)
	; fname и path из открытия
	mov fname, SaveFileDlg(WND, hInstance, "Сохранение изображения", addr path)
	cmp BYTE PTR [eax], 0
	je @F
		invoke	SaveToFile, fname
@@:



.ELSEIF (ax==IDM_NEW)
	invoke	DlgShow



.ELSEIF (ax==IDM_SELFG)
	invoke	SelectColor, addr FGColor
	invoke	SetPenParam, FGColor, CurPenSize
	invoke	InvalidateRect, hStatusBar, NULL, TRUE
	


.ELSEIF (ax==IDM_SELBG)
	invoke	SelectColor, addr BGColor
	invoke	SetBrushColor, BGColor
	invoke	InvalidateRect, hStatusBar, NULL, TRUE
	
	

.ELSEIF (ax==IDM_EXIT)
	invoke	SendMessage, WND, WM_CLOSE, 0, 0
		
.ENDIF

	ret
MenuProcess endp