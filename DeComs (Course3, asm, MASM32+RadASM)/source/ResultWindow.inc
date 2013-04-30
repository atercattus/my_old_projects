.const
	IDC_EDT1 equ 2002
	IDC_EDT2 equ 2004
	IDC_EDT3 equ 2006
	IDC_EDT4 equ 2008
	IDC_EDT5 equ 2010
	IDC_EDT6 equ 2012

	IDC_BTN1 equ 2013
	IDC_BTN2 equ 2014

	cDlgName db "ParamsDlg",0
	
.data
	hDlg 		DWORD 	0
	 
.code

Z2D proc text:DWORD, ID:DWORD
; # Преобразую внутренний формат хранения в нормально выводимый. И вывод в соответствующее поле.

	; заменяю одиночные #0 на (,)
	invoke	CopyZero2Delim, tmp_str, text
	
	; вывожу в поле
	
	invoke	SetDlgItemText, hDlg, ID, tmp_str
	
	ret	
Z2D endp

D2Z proc ID:DWORD, text:DWORD
; # Преобразую строку в соответствующем поле во внутренний формат хранения.

	invoke	GetDlgItemText, hDlg, ID, tmp_str, MAX_PATH
	invoke	strCopy, text, tmp_str
	invoke	Delim2Zero, text		
	
	ret
D2Z endp

DlgWndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
; # Обработчик диалогового окна.

	.IF (uMsg==WM_INITDIALOG)
		m2m		hDlg,	hWnd
	
	.ELSEIF (uMsg==WM_SHOWWINDOW)
	
		; Загружаю ...	
		invoke	Z2D,	addr OLC,		IDC_EDT1
		invoke	Z2D,	addr MLCB,		IDC_EDT2
		invoke	Z2D,	addr MLCE,		IDC_EDT3
		invoke	Z2D,	addr TD,		IDC_EDT4
		invoke	Z2D,	addr format,	IDC_EDT5
		invoke	SetDlgItemText, hDlg, IDC_EDT6, StartPath
		
		invoke	GetDlgItem, hWnd, IDC_EDT1
		invoke	SetFocus, eax
		
	.ELSEIF (uMsg==WM_CLOSE)
@@WM_CLOSE:
		invoke	EndDialog, hWnd, NULL
		mov		hDlg, NULL
		
	.ELSEIF (uMsg==WM_COMMAND)
		mov	eax, wParam
		mov	edx, eax
		shr	edx, 16
		.IF (dx==BN_CLICKED)
			.IF (eax==IDC_BTN2)
				jmp		@@WM_CLOSE
				
			.ELSEIF (eax==IDC_BTN1)
			
				; Сохраняю ...				
				invoke	D2Z, IDC_EDT1, addr OLC
				invoke	D2Z, IDC_EDT2, addr MLCB
				invoke	D2Z, IDC_EDT3, addr MLCE
				invoke	D2Z, IDC_EDT4, addr TD
				invoke	D2Z, IDC_EDT5, addr format
				invoke	GetDlgItemText, hDlg, IDC_EDT6, StartPath, MAX_PATH
				
				jmp		@@WM_CLOSE
				
			.ENDIF
		.ENDIF
	.ELSE
		xor		eax,	eax
		ret
	.ENDIF
	
	mov	eax, TRUE
	ret
DlgWndProc endp

ResultDlgShow proc
; # Вызов диалога настройки параметров. 

	invoke	DialogBoxParam, hInstance, addr cDlgName, WND, addr DlgWndProc, NULL
	
	ret
ResultDlgShow endp