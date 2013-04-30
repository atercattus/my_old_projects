.const
	IDC_PASSEDIT	equ 1002
	IDC_DRIVELETTER	equ 1006
	IDC_CONFIRM		equ 1003
	IDC_PCANCEL		equ 1004
	
	cPasswordDlgName db "PASSDLG",0
	
.data?
	password_buff	db 101 dup (?)
	drive_letter	db ?

.code

FillDriveLetters proc hWndComboBox:DWORD
LOCAL	DriveMask : DWORD
LOCAL	cur_char[2] : BYTE

	mov		word ptr [cur_char],	0041h ; 'A',0
	mov		DriveMask, $invoke(GetLogicalDrives)
	
	push	ebx
	mov		ebx,	03FFFFFFh ; 2^26 (букв латинского алфавита)
	.WHILE (ebx != 0)
		mov		eax,	DriveMask		
		and		eax,	1
		shr		DriveMask,	1
		shr		ebx,	1
		
		.IF (eax == 0)
			invoke	SendMessage, hWndComboBox, CB_ADDSTRING, 0, addr cur_char
		.ENDIF
		
		inc	byte ptr [cur_char]
	.ENDW
	pop		ebx

	invoke	SendMessage, hWndComboBox, CB_SETCURSEL, 0, 0
	
	ret
FillDriveLetters endp

PasswordDlgWndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM

	.IF (uMsg==WM_SHOWWINDOW)
		invoke	SendDlgItemMessage, hWnd, IDC_PASSEDIT, EM_EXLIMITTEXT, 0, 100
		invoke	SetFocus, $invoke(GetDlgItem, hWnd, IDC_PASSEDIT)
		
	.ELSEIF (uMsg == WM_INITDIALOG)
		invoke	FillDriveLetters, $invoke(GetDlgItem, hWnd, IDC_DRIVELETTER)
		
	.ELSEIF (uMsg==WM_CLOSE)
		invoke	EndDialog, hWnd, NULL
		
	.ELSEIF (uMsg==WM_COMMAND)
		mov	eax, wParam
		mov	edx, eax
		shr	edx, 16
		.IF (dx==BN_CLICKED)
			.IF (eax == IDC_PCANCEL)
				mov		byte ptr [password_buff], 0
				invoke	SendMessage, hWnd, WM_CLOSE, NULL, NULL
			
			.ELSEIF (eax == IDC_CONFIRM)
				; буква диска
				invoke	GetDlgItemText, hWnd, IDC_DRIVELETTER, addr password_buff, 101
				mov		al,	[password_buff]
				mov		[drive_letter], al
				
				; пароль
				invoke	GetDlgItemText, hWnd, IDC_PASSEDIT, addr password_buff, 101
				invoke	trim, addr password_buff
				invoke	szLen, addr password_buff
				.IF (eax < 6)
					; длина недостаточна
					invoke	MessageBox, hWnd, $CTA0('Пароль должен содержать не менее 6 символов'), NULL, MB_OK or MB_ICONERROR
				.ELSE
					invoke	SendMessage, hWnd, WM_CLOSE, NULL, NULL
				.ENDIF
			.ENDIF
		.ENDIF
	.ELSE
		mov	eax, FALSE
		ret
	.ENDIF
	
	mov	eax, TRUE
	ret
PasswordDlgWndProc endp

PasswordDlgShow proc
	
	mov		byte ptr [password_buff], 0
	mov		drive_letter, 0
	invoke	DialogBoxParam, hInstance, addr cPasswordDlgName, hMainWnd, addr PasswordDlgWndProc, NULL
	lea		eax,	password_buff
	
	ret
PasswordDlgShow endp