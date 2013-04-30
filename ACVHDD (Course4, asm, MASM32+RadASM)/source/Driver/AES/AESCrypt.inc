include AES\boxes-fst.inc
include AES\consts.inc
include AES\usefull_functions.inc
include AES\EncodeDecode.inc

.data?
	sys_init	DWORD		?		; инициализирована-ли система (когда TRUE)
	keyInst_E	keyInstance <>		; ключ для шифрования
	keyInst_D	keyInstance <>		; ключ для расшифровывания
	cipherInst	cipherInstance <>	; шифровщик
	
.const
	; число блоков (16 байт) в минимальном блоке чтения/записи диска (512 байт)
	MIN_BLOCK_COUNT		equ 512/BYTSPERBLOCK	; = 32
	MIN_BLOCK_COUNT_POW	equ 5					; 32 блока = 2^5 

.code

InitAES proto passwrd:DWORD
DoneAES proto
EncodeBuffer proto src:DWORD, dst:DWORD, block_count:DWORD
DecodeBuffer proto src:DWORD, dst:DWORD, block_count:DWORD

InitAES proc passwrd:DWORD
LOCAL	keyMaterial[320] : BYTE

	.if (sys_init == TRUE)
		return(1)
	.endif
	
	push	esi
	push	edi
	
	invoke	populateKeyMaterial, passwrd, addr keyMaterial
	
	invoke	makeKey, addr keyInst_E, DIR_ENCRYPT, 256, addr keyMaterial
	.if (eax != TRUE)
		return(-2)
	.endif
	
	invoke	makeKey, addr keyInst_D, DIR_DECRYPT, 256, addr keyMaterial
	.if (eax != TRUE)
		return(-3)
	.endif
	
	mov		eax,	offset cipherInst
	assume	eax : ptr cipherInstance
	lea		eax,	[eax].IV
	assume	eax : ptr nothing
	
	pop		edi
	pop		esi
	
	mov		sys_init,	TRUE
	
	return(0)
InitAES endp

DoneAES proc
	.if (sys_init != TRUE)
		return(1)
	.endif
	mov		sys_init,	FALSE
	return(0)
DoneAES endp

EncodeBuffer proc src:DWORD, dst:DWORD, block_count:DWORD
LOCAL	block[4] : DWORD
LOCAL	b, i, j : DWORD
LOCAL	cI_IV : DWORD
LOCAL	keySched : DWORD

	.if (sys_init != TRUE)
		return(1)
	.endif
	
	push	ebx
	push	esi
	push	edi
	
	mov		ebx,	offset cipherInst
	assume	ebx : ptr cipherInstance
	lea		ebx,	[ebx].IV
	assume	ebx : ptr nothing
	mov		cI_IV,	ebx
	
	lea		ebx,	keyInst_E
	assume	ebx : ptr keyInstance
	lea		ebx,	[ebx].keySched
	assume	ebx : ptr nothing
	mov		keySched,	ebx
	
	mov		esi,	src
	mov		edi,	dst
	lea		ebx,	block
	
	mov		eax,	block_count
	shr		eax,	MIN_BLOCK_COUNT_POW
	mov		i,		eax ; число блоков передачи данных (512 байтных)
	.while (i > 0)
		; инициализация ключа перед кодированием одного 512-байтного блока
		push	esi
		push	edi
		mov		edi,	cI_IV
		mov		esi,	keySched
		esi2edi
		add		esi,	16
		add		edi,	16
		esi2edi
		pop		edi
		pop		esi
		
		mov		j,	MIN_BLOCK_COUNT
		.while (j > 0)
;			invoke	copy_block, ebx, cI_IV
;			mov		eax,		[esi+0]
;			xor		[ebx+0],	eax
;			mov		eax,		[esi+4]
;			xor		[ebx+4],	eax
;			mov		eax,		[esi+8]
;			xor		[ebx+8],	eax
;			mov		eax,		[esi+12]
;			xor		[ebx+12],	eax
;			
;			invoke	rijndaelEncrypt, ebx, edi, keySched
;			
;			invoke	copy_block, cI_IV, edi
			
			invoke	rijndaelEncrypt, esi, edi, keySched
			
			add		esi,	16
			add		edi,	16
			
			dec		j
		.endw
		
		dec i
	.endw
	
	pop		edi
	pop		esi
	pop		ebx

	return(0)
EncodeBuffer endp

DecodeBuffer proc src:DWORD, dst:DWORD, block_count:DWORD
LOCAL	block[4], block2[4] : DWORD
LOCAL	b, i, j : DWORD
LOCAL	cI_IV : DWORD
LOCAL	keySched : DWORD

	.if (sys_init != TRUE)
		return(1)
	.endif

	push	ebx
	push	esi
	push	edi

	mov		ebx,	offset cipherInst
	assume	ebx : ptr cipherInstance
	lea		ebx,	[ebx].IV
	assume	ebx : ptr nothing
	mov		cI_IV,	ebx
	
	lea		ebx,	keyInst_D
	assume	ebx : ptr keyInstance
	lea		ebx,	[ebx].keySched
	assume	ebx : ptr nothing
	mov		keySched,	ebx
	
	mov		esi,	src
	mov		edi,	dst
	lea		ebx,	block
	lea		edx,	block2
	
	mov		eax,	block_count
	shr		eax,	MIN_BLOCK_COUNT_POW
	mov		i,		eax ; число блоков передачи данных (512 байтных)
	.while (i > 0)
		; инициализация ключа перед декодированием одного 512-байтного блока
		push	esi
		push	edi
		mov		edi,	cI_IV
		mov		esi,	keySched
		esi2edi
		add		esi,	16
		add		edi,	16
		esi2edi
		pop		edi
		pop		esi
		
		mov		j,	MIN_BLOCK_COUNT
		.while (j > 0)
;			invoke	rijndaelDecrypt, esi, edx, keySched
;			
;			invoke	copy_block, ebx, cI_IV
;			mov		eax,		[edx+0]
;			xor		[ebx+0],	eax
;			mov		eax,		[edx+4]
;			xor		[ebx+4],	eax
;			mov		eax,		[edx+8]
;			xor		[ebx+8],	eax
;			mov		eax,		[edx+12]
;			xor		[ebx+12],	eax
;			
;			invoke	copy_block, cI_IV, esi
;			
;			invoke	copy_block, edi, ebx
			
			invoke	rijndaelDecrypt, esi, edi, keySched
			
			add		esi,	16
			add		edi,	16
			
			dec	j
		.endw
		
		dec	i
	.endw

	pop		edi
	pop		esi
	pop		ebx

	return(0)
DecodeBuffer endp