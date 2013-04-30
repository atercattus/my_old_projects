.const
	IDC_STC1 equ 2001
	IDC_STC2 equ 2002
	IDC_EDT1 equ 2003
	IDC_EDT2 equ 2004
	IDC_CHK1 equ 2005
	IDC_BTN1 equ 2006
	IDC_BTN2 equ 2007

	cDlgName db "ParamsDlg",0
	cDlgClass db "PARAMSDLGCLASS",0
	
.data
	hDlg 		DWORD 	0
	CheckBoxVal	DWORD	0
	 
.code

DlgWndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
	LOCAL buff[8] : BYTE

	.IF (uMsg==WM_SHOWWINDOW)
		; ширина
		invoke	ltoa, Rect.right, addr buff
		invoke	SetDlgItemText, hWnd, IDC_EDT1, addr buff
		; высота
		invoke	ltoa, Rect.bottom, addr buff
		invoke	SetDlgItemText, hWnd, IDC_EDT2, addr buff
	
		invoke	GetDlgItem, hWnd, IDC_EDT1
		invoke	SetFocus, eax
		
	.ELSEIF (uMsg==WM_CLOSE)
		invoke	EndDialog, hWnd, NULL
		mov		hDlg, NULL
		
	.ELSEIF (uMsg==WM_COMMAND)
		mov	eax, wParam
		mov	edx, eax
		shr	edx, 16
		.IF (dx==BN_CLICKED)
			.IF (eax==IDC_BTN2)
				invoke	SendMessage, hWnd, WM_CLOSE, NULL, NULL
				
			.ELSEIF (eax==IDC_CHK1)
				mov		eax,	CheckBoxVal
				not		eax
				mov		CheckBoxVal,	eax
				
			.ELSEIF (eax==IDC_BTN1)
			
				push	CheckBoxVal
				
				invoke	GetDlgItemText, hWnd, IDC_EDT2, addr buff, 6
				invoke	ustr2dw, addr buff
				push	eax
				
				invoke	GetDlgItemText, hWnd, IDC_EDT1, addr buff, 6
				invoke	ustr2dw, addr buff
				push	eax
				
				push	WND
				
				call	CreateBackBuffer
				
				invoke	SendMessage, hWnd, WM_CLOSE, NULL, NULL
				
				invoke	Resize
				
				invoke	InvalidateRect, WND, NULL, TRUE
								
			.ENDIF
		.ENDIF
	.ELSE
		mov	eax, FALSE
		ret
	.ENDIF
	
	mov	eax, TRUE
	ret
DlgWndProc endp

DlgShow proc
	
	invoke	DialogBoxParam, hInstance, addr cDlgName, WND, addr DlgWndProc, NULL
	mov		hDlg,	eax
	
	ret
DlgShow endp