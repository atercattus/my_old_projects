.const

	IDC_LST equ 2001
	IDC_BTN equ 2002

	cFileListDlgName	db "FileListDlg",0
	
.data?
	hFileList 		DWORD 	?
	hTreeView		DWORD	?
	hTmpTreeView	DWORD	?
.code

TreeView_TmpSwap proto

FileListDlgWndProc proc hWnd:HWND, uMsg:UINT, wParam:WPARAM, lParam:LPARAM
; # Обработчик диалогового окна.

	.IF (uMsg==WM_INITDIALOG)
		m2m		hFileList,	hWnd
		invoke	GetDlgItem, hWnd, IDC_LST
		mov		hTreeView,	eax
		
		invoke	TreeView_TmpSwap	
	
	.ELSEIF (uMsg==WM_SHOWWINDOW)
		invoke	SetFocus, hTreeView
		
	.ELSEIF (uMsg==WM_CLOSE)
@@WM_CLOSE:
		invoke	EndDialog, hWnd, NULL
		mov		hFileList, NULL
		
	.ELSEIF (uMsg==WM_COMMAND)
		mov	eax, wParam
		mov	edx, eax
		shr	edx, 16
		cmp		dx,	BN_CLICKED
		jne		@F
			cmp		eax,	IDC_BTN
			je		@@WM_CLOSE
@@:
	.ELSE
		xor		eax,	eax
		ret
	.ENDIF
	
	mov	eax, TRUE
	ret
FileListDlgWndProc endp

FileListDlgShow proc
; # Вызов диалога настройки параметров.

	invoke	DialogBoxParam, hInstance, addr cFileListDlgName, WND, addr FileListDlgWndProc, NULL
	
	ret
FileListDlgShow endp

TreeView_DeleteItem macro hwnd, hitem
	invoke	SendMessage, hwnd, TVM_DELETEITEM, 0, hitem
endm

; наборчик различных макросов.

TreeView_DeleteAllItems macro hwnd
	invoke	SendMessage, hwnd, TVM_DELETEITEM, 0, TVI_ROOT
endm

TreeView_GetCount macro hwnd
	invoke	SendMessage, hwnd, TVM_GETCOUNT, 0, 0
	exitm <eax>
endm

TreeView_SetBkColor macro hwnd, clr
	invoke	SendMessage, hwnd, TVM_SETBKCOLOR, 0, clr
endm

TreeView_SetTextColor macro hwnd, clr
	invoke	SendMessage, hwnd, TVM_SETTEXTCOLOR, 0, clr
endm

TreeView_GetNextItem macro hwnd, hitem, code
	invoke	SendMessage, hwnd, TVM_GETNEXTITEM, code, hitem
endm

TreeView_GetChild macro hwnd, hitem
	TreeView_GetNextItem	hwnd, hitem, TVGN_CHILD
endm

TreeView_GetParent macro hwnd, hitem
	TreeView_GetNextItem	hwnd, hitem, TVGN_PARENT
	exitm <eax>
endm

TreeView_GetNext macro hwnd, hitem
	TreeView_GetNextItem	hwnd, hitem, TVGN_NEXT
endm

TreeView_GetPrev macro hwnd, hitem
	TreeView_GetNextItem	hwnd, hitem, TVGN_PREVIOUS
endm

TreeView_GetRoot macro hwnd
	TreeView_GetNextItem	hwnd, NULL, TVGN_ROOT
	exitm <eax>
endm

TreeView_InsertItem proc hwnd:DWORD, parent:DWORD, text:DWORD
; # Добавление элемента.

	LOCAL	param : TV_INSERTSTRUCT
	
	invoke	RtlZeroMemory, addr param, sizeof(TV_INSERTSTRUCT)
	
	m2m		param.hParent,			parent
	mov		param.hInsertAfter,		TVI_LAST

	mov		param.item._mask,		TVIF_TEXT or TVIF_STATE
	m2m		param.item.pszText,		text
	mov		param.item.state,		TVIS_EXPANDED
	mov		param.item.stateMask,	TVIS_EXPANDED
	
	invoke	SendMessage, hwnd, TVM_INSERTITEM, 0, addr param
	
	ret
TreeView_InsertItem endp

TreeView_CreateTemporary proc where:DWORD
; # Создание временного "дерева".
	
	invoke	CreateWindowEx, 0, SADD("SysTreeView32"), SADD("TVTC"), TVS_HASLINES, 0, 0, 1, 1, where, NULL, NULL, NULL
	mov		hTmpTreeView,	eax	 
	
	ret
TreeView_CreateTemporary endp

TreeView_DeleteTemporary proc
; # Удаление временного "дерева".

	invoke	DestroyWindow, hTmpTreeView
	
	ret
TreeView_DeleteTemporary endp

TreeView_GetItemText proc hwnd, hitem, text, text_len : DWORD
; # Получение строки указанного элемента.

	LOCAL	item : TV_ITEM
	
	invoke	RtlZeroMemory, addr item, sizeof(TV_ITEM)
	
	mov		item._mask,			TVIF_TEXT
	m2m		item.hItem,			hitem
	
	m2m		item.pszText,		text
	m2m		item.cchTextMax,	text_len
	
	invoke	SendMessage, hwnd, TVM_GETITEM, 0, addr item
	
	ret
TreeView_GetItemText endp

.data?
	; указатель на текущий добавляемый элемент	
	item		dd	?
.code

TreeView_RecurseAdding proc first : DWORD;, level : DWORD
; # Рекурсивный обход "дерева".

	LOCAL	tmp_item : DWORD
	
	TreeView_GetChild hTmpTreeView, first
	cmp		eax,	NULL
	je		@@enumerate_exit
	
	mov		tmp_item,	eax
	
@@enumerate_loop:

	; добавляю текущее имя 
	invoke	TreeView_GetItemText, hTmpTreeView, tmp_item, tmp_str, MAX_PATH

	invoke	szLen,	tmp_str
	add		eax,	tmp_str
	dec		eax
	.IF (byte ptr [eax] == '?')
		mov		byte ptr [eax], 0
	.ELSE
		invoke	check_on_slash, tmp_str
	.ENDIF
	
	invoke	TreeView_InsertItem, hTreeView, item, tmp_str
	
	; обрабатываю подразделы
	push	item
	m2m		item,	eax
	invoke	TreeView_RecurseAdding, tmp_item
	pop		item
	
	; перехожу к следующему
	TreeView_GetNext hTmpTreeView, tmp_item
	cmp		eax,	0
	je		@@enumerate_exit
	
	mov		tmp_item,	eax
	
	jmp		@@enumerate_loop

@@enumerate_exit:
	
	ret
TreeView_RecurseAdding endp

TreeView_TmpSwap proc
; # Копирование "дерева" из временного списка в видимый.

	TreeView_DeleteAllItems(hTreeView)
	
	invoke	TreeView_InsertItem, hTreeView, NULL, StartPath
	mov		item,	eax
	
	invoke	TreeView_RecurseAdding, TreeView_GetRoot(hTmpTreeView);, 1
	
	TreeView_DeleteAllItems(hTmpTreeView)

	ret
TreeView_TmpSwap endp