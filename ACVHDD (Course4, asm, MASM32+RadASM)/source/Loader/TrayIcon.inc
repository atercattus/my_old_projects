.const
	IDI_TRAY			equ	1
	WM_TRAY				equ	WM_USER+123
	
	IDI_TRAYCONN		equ	104
	IDI_TRAYDISCONN		equ	105
	IDI_TRAYNOTEXIST	equ	106
	
	TRAYICON_STRING_LOADDRIVER		db	"Загрузить образ диска",0
	IDM_LOADDRIVER					equ	1001
	TRAYICON_STRING_UNLOADDRIVER	db	"Выгрузить диск",0
	IDM_UNLOADDRIVER				equ	1002
	TRAYICON_STRING_CREATEIMAGE		db	"Создать образ диска",0
	IDM_CREATEIMAGE					equ	1003
	TRAYICON_STRING_DELETEIMAGE		db	"Удалить образ диска",0
	IDM_DELETEIMAGE					equ	1004
	TRAYICON_STRING_EXIT			db	"Выход",0
	IDM_EXIT						equ	1005

.data?
	hTrayPopup		dd ?
	hTrayConnIcon	dd ?
	hTrayDiscIcon	dd ?
	hTrayNotExIcon	dd ?

.code

AC_TrayIconCreate proc
LOCAL	nid : NOTIFYICONDATA

	mov		hTrayConnIcon,  $invoke(LoadIcon, hInstance, IDI_TRAYCONN)
	mov		hTrayDiscIcon,  $invoke(LoadIcon, hInstance, IDI_TRAYDISCONN)
	mov		hTrayNotExIcon, $invoke(LoadIcon, hInstance, IDI_TRAYNOTEXIST)
	
	mov		nid.cbSize,				sizeof(NOTIFYICONDATA)
	m2m		nid.hwnd,				hMainWnd
	mov		nid.uID,				IDI_TRAY
	mov		nid.uFlags,				NIF_ICON or NIF_MESSAGE or NIF_TIP
	mov		nid.uCallbackMessage,	WM_TRAY
	m2m		nid.hIcon,				hTrayNotExIcon
	invoke	lstrcpy, addr nid.szTip, addr APPLICATION_NAME
	
	invoke	Shell_NotifyIcon, NIM_ADD, addr nid
	
	mov		hTrayPopup, $invoke(CreatePopupMenu)
	invoke	AppendMenu, hTrayPopup, MF_STRING, IDM_LOADDRIVER, addr TRAYICON_STRING_LOADDRIVER
	invoke	AppendMenu, hTrayPopup, MF_STRING, IDM_UNLOADDRIVER, addr TRAYICON_STRING_UNLOADDRIVER
	invoke	AppendMenu, hTrayPopup, MF_SEPARATOR, IDM_EXIT, NULL
	invoke	AppendMenu, hTrayPopup, MF_STRING, IDM_CREATEIMAGE, addr TRAYICON_STRING_CREATEIMAGE
	invoke	AppendMenu, hTrayPopup, MF_STRING, IDM_DELETEIMAGE, addr TRAYICON_STRING_DELETEIMAGE
	invoke	AppendMenu, hTrayPopup, MF_SEPARATOR, IDM_EXIT, NULL
	invoke	AppendMenu, hTrayPopup, MF_STRING, IDM_EXIT, addr TRAYICON_STRING_EXIT

	ret
AC_TrayIconCreate endp

AC_TrayIconDelete proc
LOCAL	nid : NOTIFYICONDATA

	invoke	DestroyMenu, hTrayPopup

	mov		nid.cbSize,				sizeof(NOTIFYICONDATA)
	m2m		nid.hwnd,				hMainWnd
	mov		nid.uID,				IDI_TRAY

	invoke	Shell_NotifyIcon, NIM_DELETE, addr nid
	
	ret
AC_TrayIconDelete endp

AC_TrayIconChange proc hNewIcon:DWORD
LOCAL	nid : NOTIFYICONDATA

	mov		nid.cbSize,				sizeof(NOTIFYICONDATA)
	m2m		nid.hwnd,				hMainWnd
	mov		nid.uID,				IDI_TRAY
	mov		nid.uFlags,				NIF_ICON
	m2m		nid.hIcon,				hNewIcon
	
	invoke	Shell_NotifyIcon, NIM_MODIFY, addr nid

	ret
AC_TrayIconChange endp

AC_TrayIconWndProc_WM_TRAY proc wParam:DWORD, lParam:DWORD
LOCAL	pt : POINT

	.if wParam==IDI_TRAY
		.if (lParam==WM_RBUTTONDOWN)
			invoke	GetCursorPos, addr pt
			invoke	SetForegroundWindow, hMainWnd
			invoke	TrackPopupMenu, hTrayPopup, TPM_LEFTALIGN, pt.x, pt.y, NULL, hMainWnd, NULL
		.endif
	.endif

	ret
AC_TrayIconWndProc_WM_TRAY endp

AC_Cmd_LoadDriver proto
AC_Cmd_UnloadDriver proto
AC_Cmd_CreateImage proto
AC_Cmd_DeleteImage proto
NewImageDlgShow proto
DelImageDlgShow proto
AC_TrayIconMenuItemsCheckState proto

AC_TrayIconWndProc_WM_COMMAND proc wParam:DWORD, lParam:DWORD

	mov		eax,	wParam
	.if (ax == IDM_EXIT)
		invoke	SendMessage, hMainWnd, WM_CLOSE, 0, 0
		
	.elseif (ax == IDM_LOADDRIVER)
		invoke	AC_Cmd_LoadDriver, NULL
;		.if (eax == -1)
;			; nothing do
;		.elseif (eax != FALSE)
;		.else
;			; вроде как это сообщение всегда дублирует реальное сообщение об ошибке. так что нафиг его.
;;			invoke	MessageBox, hMainWnd, $CTA0("Во время загрузки драйвера диска произошел сбой"), addr APPLICATION_NAME, MB_OK or MB_ICONERROR
;		.endif
		invoke	AC_TrayIconMenuItemsCheckState
	
	.elseif (ax == IDM_UNLOADDRIVER)
		invoke	AC_Cmd_UnloadDriver
		invoke	AC_TrayIconMenuItemsCheckState
;		.if (eax == TRUE)
;			invoke	AC_TrayIconChange, hTrayNotExIcon
;		.endif
	
	.elseif (ax == IDM_CREATEIMAGE)
		invoke	NewImageDlgShow
	
	.elseif (ax == IDM_DELETEIMAGE)
		invoke	DelImageDlgShow
	
	.endif

	ret
AC_TrayIconWndProc_WM_COMMAND endp

; выставляет пункты меню в ENABLED/DISABLED в зависимости от состояния системы
AC_TrayIconMenuItemsCheckState proc
	invoke	AC_FindDriverInMem
	.if (eax == FIND_DRIVER_NOT_FOUND)
		invoke	AC_TrayIconChange, hTrayNotExIcon
		invoke	EnableMenuItem, hTrayPopup, IDM_LOADDRIVER, MF_ENABLED
		invoke	EnableMenuItem, hTrayPopup, IDM_UNLOADDRIVER, MF_DISABLED or MF_GRAYED
		
	.elseif (eax == FIND_DRIVER_STOPPED)
		invoke	AC_TrayIconChange, hTrayDiscIcon
		invoke	EnableMenuItem, hTrayPopup, IDM_LOADDRIVER, MF_ENABLED
		invoke	EnableMenuItem, hTrayPopup, IDM_UNLOADDRIVER, MF_DISABLED or MF_GRAYED
		
	.else ; (eax == FIND_DRIVER_RUNNING)
		invoke	AC_TrayIconChange, hTrayConnIcon
		invoke	EnableMenuItem, hTrayPopup, IDM_LOADDRIVER, MF_DISABLED or MF_GRAYED
		invoke	EnableMenuItem, hTrayPopup, IDM_UNLOADDRIVER, MF_ENABLED
	.endif

	ret
AC_TrayIconMenuItemsCheckState endp