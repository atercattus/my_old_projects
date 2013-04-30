.const
	IDC_DEL_IMAGEPATH	equ 1002
	IDC_DEL_CHANGEPATH	equ 1003
	IDC_DEL_DELETE		equ 1004
	IDC_DEL_CANCEL		equ 1005
	
	MAX_IMAGE_SIZE	equ 4096
	
	cDelImageDlgName db "DELIMAGEDLG",0
	
.data?
	hDelProgressBar	DWORD ?
	
.code

AC_CreateNewImage proto imgsize:DWORD, imgpath:DWORD 

DelImageDlgWndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
LOCAL	hFile			: DWORD
LOCAL	ofn				: OPENFILENAME
LOCAL	Sectors			: DWORD
LOCAL	BytesWritten	: DWORD

	.IF (uMsg==WM_INITDIALOG)
		invoke	CreateWindowEx,	NULL, $CTA0('msctls_progress32'), NULL, WS_CHILD or WS_VISIBLE or PBS_SMOOTH,\
								6, 70, 462, 15, hWnd, 1, hInstance, NULL
		mov		hDelProgressBar,	eax
		; диалог можно закрыть
		mov		MayClose,	TRUE

	.ELSEIF (uMsg==WM_SHOWWINDOW)
		invoke	GetDlgItem, hWnd, IDC_DEL_IMAGEPATH
		invoke	SetFocus, eax
		
	.ELSEIF (uMsg==WM_CLOSE)
		.if (MayClose==TRUE)
			invoke	EndDialog, hWnd, NULL
		.endif
		
	.ELSEIF (uMsg==WM_COMMAND)
		mov	eax, wParam
		mov	edx, eax
		shr	edx, 16
		.IF (dx==BN_CLICKED)
			.IF (eax == IDC_DEL_CANCEL)
				invoke	SendMessage, hWnd, WM_CLOSE, NULL, NULL
				
			.ELSEIF (eax == IDC_DEL_CHANGEPATH)
				; выбор файла образа
				invoke	RtlZeroMemory, addr ofn, sizeof OPENFILENAME
				invoke	RtlZeroMemory, addr BufFileName, MAXSIZE
				
				mov		ofn.lStructSize,	sizeof OPENFILENAME
				m2m		ofn.hwndOwner,		hMainWnd
				m2m		ofn.hInstance,		hInstance
				mov		ofn.lpstrFilter,	offset OPEN_IMG_FILTER
				mov		ofn.lpstrFile,		offset BufFileName
				mov		ofn.nMaxFile,		MAXSIZE
				mov		ofn.Flags,			OFN_FILEMUSTEXIST or OFN_PATHMUSTEXIST or OFN_LONGNAMES or \
											OFN_EXPLORER or OFN_HIDEREADONLY
											
				invoke	GetOpenFileName, addr ofn
				or		eax,	eax
				jz		@@DIDWP_exit
					invoke	GetDlgItem, hWnd, IDC_DEL_IMAGEPATH
					invoke	SetDlgItemText, hWnd, IDC_DEL_IMAGEPATH, addr BufFileName
				
			.ELSEIF (eax == IDC_DEL_DELETE)
					invoke	GetDlgItem, hWnd, IDC_DEL_IMAGEPATH
					invoke	GetDlgItemText, hWnd, IDC_DEL_IMAGEPATH, addr BufFileName, MAXSIZE
					
					invoke	CreateFile, addr BufFileName, GENERIC_WRITE, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0
					.if (eax == INVALID_HANDLE_VALUE)
						; ошибка при монопольном открытии файла
						invoke	MessageBox,	hMainWnd, $CTA0('Ошибка монопольного открытия файла'),
											addr APPLICATION_NAME, MB_OK or MB_ICONERROR
						jmp		@@DIDWP_exit
					.endif
					mov		hFile,	eax
					
					invoke	LocalAlloc, LMEM_FIXED, 1024*1024
					mov		Sectors,	eax
					
					push	ebx
					
					; запрет закрытия диалога
					mov		MayClose,	FALSE
					
					invoke	GetFileSize, hFile, NULL
					shr		eax,	20 ; размер в Мб
					mov		ebx,	eax 
					; задание начального состояния ProgressBar'а
					shl		eax,	16
					invoke	SendMessage, hDelProgressBar, PBM_SETRANGE, 0, eax
					invoke	SendMessage, hDelProgressBar, PBM_SETPOS, 0, 0
					invoke	SendMessage, hDelProgressBar, PBM_SETSTEP, 1, 0
					
					.while (ebx > 0)
						invoke	WriteFile, hFile, Sectors, 1024*1024, addr BytesWritten, NULL
						dec		ebx
						; тут бы хорошо бы как-нибудь информировать о текущем состоянии дел
						invoke	SendMessage, hDelProgressBar, PBM_STEPIT, 0, 0
						invoke	ProcessMessages, hWnd
					.endw
					invoke	SendMessage, hDelProgressBar, PBM_SETPOS, 0, 0
					pop		ebx
					
					; разрешение закрытия диалога
					mov		MayClose,	TRUE
					
					invoke	LocalFree, Sectors
					
					invoke	CloseHandle, hFile
					invoke	DeleteFile, addr BufFileName
					
					invoke	MessageBox, hWnd, $CTA0('Файл образа был удален'), addr APPLICATION_NAME, MB_OK or MB_ICONINFORMATION 
			.ENDIF
		.ENDIF
	.ELSE
		return(FALSE)
	.ENDIF

@@DIDWP_exit:

	return(TRUE)
DelImageDlgWndProc endp

DelImageDlgShow proc
	
	invoke	DialogBoxParam, hInstance, addr cDelImageDlgName, hMainWnd, addr DelImageDlgWndProc, NULL
	
	ret
DelImageDlgShow endp