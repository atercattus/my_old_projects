.code

return macro code
	mov		eax,	code
	ret
endm

m2m macro a, b
	push	b
	pop		a
endm

my_strncpy proc dst:DWORD, src:DWORD, maxlen:DWORD
	push	esi
	push	edi
	push	ecx
	
	mov		esi,	src
	mov		edi,	dst
	mov		ecx,	maxlen
@@:
	mov		al,		byte ptr [esi]
	or		al,		al
	jz		@F
	mov		byte ptr [edi],	al
	inc		esi
	inc		edi
	dec		ecx
	jnz		@B
@@:
	mov		byte ptr [edi], 0
	
	pop		ecx
	pop		edi
	pop		esi
	
	return(0)
my_strncpy endp

my_memset macro dst, len, val
; никакие регистры не сохран€ютс€, т.к. используетс€ единожды в InitAES()
; перед самым выходом ѕ≈–≈ƒ восстановлением из стека ESI и EDI 

	mov		ecx,	len
	mov		edi,	dst
	mov		al,		val
	cld
	rep stosb
endm

copy_block proc dst:DWORD, src:DWORD; len не надо, т.к. ¬—≈√ƒј копируетс€ 16 байт

	push	esi
	push	edi
	
	mov		ecx,	4 ; 4DWORD
	mov		esi,	src
	mov		edi,	dst
	cld
	rep movsd
	
	pop		edi
	pop		esi
	
	ret
copy_block endp

;static BYTE convertKey( char k )
;{
;	int kval = (k % 16);
;	if ( kval < 10 )	return( '0' + kval );
;	else				return( kval - 10 + 'A' );
;}
;
;static void populateKeyMaterial( char* text, BYTE* key )
;{
;	char* textscanner = text;
;	int i;
;	for ( i = 0; i < MAX_KEY_SIZE; i++ )
;	{
;		*key++ = convertKey( *textscanner );
;		textscanner++;
;		if ( *textscanner == 0 )
;		{
;			textscanner = text;
;		}
;	} // for i
;	*key = 0;
;}

StrLen proc item:DWORD
    push    ebx
    mov     eax,item               ; get pointer to string
    lea     edx,[eax+3]            ; pointer+3 used in the end
  @@:
    mov     ebx,[eax]              ; read first 4 bytes
    add     eax,4                  ; increment pointer
    lea     ecx,[ebx-01010101h]    ; subtract 1 from each byte
    not     ebx                    ; invert all bytes
    and     ecx,ebx                ; and these two
    and     ecx,80808080h    
    jz      @B                     ; no zero bytes, continue loop
    test    ecx,00008080h          ; test first two bytes
    jnz     @F
    shr     ecx,16                 ; not in the first 2 bytes
    add     eax,2
  @@:
    shl     cl,1                   ; use carry flag to avoid branch
    sbb     eax,edx                ; compute length
    pop     ebx

    ret
StrLen endp

populateKeyMaterial proc text:DWORD, key:DWORD
; скорость не важна
	
	invoke	StrLen, text
	mov		dl,		al		; длина парол€ - это вектор инициализации
	
	mov		esi,	text
	mov		edi,	key
	mov		ecx,	MAX_KEY_SIZE

	mov		al,		[esi]
	
@@pKM_loop:
	
	xor		dl,		al
	mov		[edi],	dl
	
	inc		esi
	inc		edi
	
	mov		al,		[esi]
	or		al,		al
	jnz		@F
		mov		esi,	text
		mov		al,		[esi]
	@@:
	
	dec		ecx
	jnz		@@pKM_loop
	
	mov		byte ptr [edi], 0
	
	return(0)
populateKeyMaterial endp

;void InvMixColumn(word8 a[4][4], word8 BC)
;{ /* Mix the four bytes of every column in a linear way */
;	int j;
;	for (j = 0; j < BC; j++)
; 		*((word32*)a[j]) =	*((word32*)U1[a[j][0]]) ^
;							*((word32*)U2[a[j][1]]) ^
;							*((word32*)U3[a[j][2]]) ^
;							*((word32*)U4[a[j][3]]);
;}

; void InvMixColumn(word8 a[4][4], word8 BC)
InvMixColumn proc a:DWORD; , BC:BYTE == 4!!!
; скорость не важна
; можно «Ќј„»“≈Ћ№Ќќ оптимизировать!

	push	ebx
	
	xor		ecx,	ecx
	mov		cl,		4
	mov		edi,	a

@@IMC_loop:
	; EDI == &a[j]
	
	xor		eax,	eax
	mov		al,		[edi+0]
	mov		edx,	[offset U1+eax*4]
	
	mov		al,		[edi+1]
	xor		edx,	[offset U2+eax*4]
	
	mov		al,		[edi+2]
	xor		edx,	[offset U3+eax*4]
	
	mov		al,		[edi+3]
	xor		edx,	[offset U4+eax*4]
	
	mov		[edi],	edx
	add		edi,	4
	
	dec		ecx
	jnz		@@IMC_loop
	
	pop		ebx	
	
	return(0)
InvMixColumn endp

;int rijndaelKeySched (word8 k[MAXKC][4], word8 W[MAXROUNDS+1][4][4])
;{
;	/* Calculate the necessary round keys
;	 * The number of calculations depends on keyBits and blockBits
;	 */ 
;	int j, r, t, rconpointer = 0;
;	word8 tk[MAXKC][4];
;	int KC = ROUNDS - 6;
;	
;	for(j = KC-1; j >= 0; j--)
;		*((word32*)tk[j]) = *((word32*)k[j]);
;	r = 0;
;	t = 0;
;	/* copy values into round key array */
;	for(j = 0; (j < KC) && (r < (ROUNDS+1)); )
;	{
;		for (; (j < KC) && (t < 4); j++, t++)
;			*((word32*)W[r][t]) = *((word32*)tk[j]);
;		if (t == 4)
;		{
;			r++;
;			t = 0;
;		}
;	}
;		
;	while (r < (ROUNDS+1))
;	{ /* while not enough round key material calculated */
;		/* calculate new values */
;		tk[0][0] ^= S[tk[KC-1][1]];
;		tk[0][1] ^= S[tk[KC-1][2]];
;		tk[0][2] ^= S[tk[KC-1][3]];
;		tk[0][3] ^= S[tk[KC-1][0]];
;		tk[0][0] ^= rcon[rconpointer++];
;
;		KC = ROUNDS-6 = 14-6 = 8 (!!!) всегда дл€ моей реализации
;
;		if (KC != 8)
;		{
;			for (j = 1; j < KC; j++)
;				*((word32*)tk[j]) ^= *((word32*)tk[j-1]);
;		}
;		else
;		{
;			for (j = 1; j < KC/2; j++)
;				*((word32*)tk[j]) ^= *((word32*)tk[j-1]);
;			tk[KC/2][0] ^= S[tk[KC/2 - 1][0]];
;			tk[KC/2][1] ^= S[tk[KC/2 - 1][1]];
;			tk[KC/2][2] ^= S[tk[KC/2 - 1][2]];
;			tk[KC/2][3] ^= S[tk[KC/2 - 1][3]];
;			for (j = KC/2 + 1; j < KC; j++)
;				*((word32*)tk[j]) ^= *((word32*)tk[j-1]);
;		}
;		/* copy values into round key array */
;		for (j = 0; (j < KC) && (r < (ROUNDS+1)); )
;		{
;			for (; (j < KC) && (t < 4); j++, t++)
;			{
;				*((word32*)W[r][t]) = *((word32*)tk[j]);
;			}
;			if (t == 4)
;			{
;				r++;
;				t = 0;
;			} // if
;		} // for j
;	} // while r		
;
;	return 0;
;}

IE macro vector, index
; получение смещени€ index элемента массива vector типа DWORD
; т.е. return &((DWORD*)vector)[index]
	push	ebx
	lea		ebx,	vector
	mov		eax,	index
	shl		eax,	2
	add		eax,	ebx
	pop		ebx
	exitm	<eax>
endm

; int rijndaelKeySched (word8 k[MAXKC][4], word8 W[MAXROUNDS+1][4][4])
rijndaelKeySched proc k:DWORD, W:DWORD
; вызываетс€ один раз при монтировании диска - наплевал на оптимизацию во им€ пон€тности
LOCAL	j, r, t, rconptr : DWORD
LOCAL	tk[MAXKC] : DWORD
LOCAL	KC : DWORD

	push	ebx
	
;	int j, r, t, rconpointer = 0;
;	word8 tk[MAXKC][4];
;	int KC = ROUNDS - 6;
	mov		rconptr,	0
	m2m		KC,			ROUNDS
	sub		KC,			6
	
;	for(j = KC-1; j >= 0; j--)
;		*((word32*)tk[j]) = *((word32*)k[j]);
	m2m		j,	KC
	dec		j
	
	.while (TRUE)
		cmp		j,	0
		jnl		@F
			.break
		@@:
		
		mov		eax,	j
		shl		eax,	2
		
		mov		ecx,	k
		add		ecx,	eax
		mov		edx,	[ecx]
		
		lea		ecx,	tk
		add		ecx,	eax
		mov		[ecx],	edx
		
		dec		j
	.endw

;	r = 0;
;	t = 0;
	xor		eax,	eax
	mov		r,		eax
	mov		t,		eax

	mov		edx,	ROUNDS
	inc		edx				; EDX == ROUNDS+1

	;	/* copy values into round key array */
	;	for(j = 0; (j < KC) && (r < (ROUNDS+1)); )
	;	{
	;		for (; (j < KC) && (t < 4); j++, t++)
	;			*((word32*)W[r][t]) = *((word32*)tk[j]);
	;		if (t == 4)
	;		{
	;			r++;
	;			t = 0;
	;		}
	;	}
	mov		ecx,	eax		; ECX == 0
	.while ((ecx < KC) && (r < edx))
		.while ((ecx < KC) && (t < 4))
			mov		esi,	IE(tk, ecx)
			
			mov		eax,	r
			shl		eax,	2
			add		eax,	t
			
			mov		edi,	W
			shl		eax,	2
			add		edi,	eax
			
			m2m		[edi],	[esi]
			
			inc		ecx
			inc		t
		.endw
		.if (t == 4)
			inc		r
			mov		t,	0
		.endif
	.endw

;	while (r < (ROUNDS+1))
	.while (r < edx)
		; tk[0][0] ^= S[tk[KC-1][1]];
		; tk[0][1] ^= S[tk[KC-1][2]];
		; tk[0][2] ^= S[tk[KC-1][3]];
		; tk[0][3] ^= S[tk[KC-1][0]];
		mov		eax,	KC
		dec		eax
		mov		eax,	IE(tk, eax)
		mov		eax,	[eax]
		ror		eax,	8
		mov		esi,	eax		; esi == (tk[KC-1] ror 8)
		lea		edi,	S		; edi == S
		
		xor		ebx,	ebx
		
		and		eax,	0FFh
		mov		bl,		byte ptr [edi+eax]
		ror		ebx,	8
		shr		esi,	8
		
		mov		eax,	esi
		and		eax,	0FFh
		mov		bl,		byte ptr [edi+eax]
		ror		ebx,	8
		shr		esi,	8
		
		mov		eax,	esi
		and		eax,	0FFh
		mov		bl,		byte ptr [edi+eax]
		ror		ebx,	8
		shr		esi,	8
		
		mov		eax,	esi
		and		eax,	0FFh ; ???
		mov		bl,		byte ptr [edi+eax]
		ror		ebx,	8
		
		lea		eax,	tk
		xor		[eax],	ebx
		
		; tk[0][0] ^= rcon[rconpointer++];
		lea		ecx,	rcon
		add		ecx,	rconptr
		inc		rconptr
		mov		cl,		[ecx]
		xor		byte ptr [eax], cl
		
		;	for (j = 1; j < KC/2; j++)
		;		*((word32*)tk[j]) ^= *((word32*)tk[j-1]);
		mov		j,		1
		mov		ecx,	KC
		shr		ecx,	1
		.while (j < ecx)
			mov		esi,	IE(tk, j)
			mov		eax,	[esi-4]
			xor		[esi],	eax
			
			inc		j
		.endw
		
		;	tk[KC/2][0] ^= S[tk[KC/2 - 1][0]];
		;	tk[KC/2][1] ^= S[tk[KC/2 - 1][1]];
		;	tk[KC/2][2] ^= S[tk[KC/2 - 1][2]];
		;	tk[KC/2][3] ^= S[tk[KC/2 - 1][3]];
		mov		eax,	KC
		shr		eax,	1
		mov		esi,	IE(tk, eax)
		mov		edi,	esi			; EDI == &tk[KC/2]
		sub		esi,	4			; ESI == &tk[KC/2 - 1]
		lea		ebx,	S			; EBX == &S
		xor		ecx,	ecx			; S[...] accumulator		
		xor		eax,	eax			; tk[...] for S[...] indexing
		
		mov		al,		byte ptr [esi+0]
		mov		cl,		byte ptr [ebx+eax]
		ror		ecx,	8
		
		mov		al,		byte ptr [esi+1]
		mov		cl,		byte ptr [ebx+eax]
		ror		ecx,	8
		
		mov		al,		byte ptr [esi+2]
		mov		cl,		byte ptr [ebx+eax]
		ror		ecx,	8
		
		mov		al,		byte ptr [esi+3]
		mov		cl,		byte ptr [ebx+eax]
		ror		ecx,	8

		xor		[edi],	ecx
		
		;	for (j = KC/2 + 1; j < KC; j++)
		;		*((word32*)tk[j]) ^= *((word32*)tk[j-1]);
		mov		eax,	KC
		mov		ecx,	eax
		shr		eax,	1
		inc		eax
		mov		j,		eax
		.while (j < ecx)
			mov		esi,	IE(tk, j)
			mov		eax,	[esi-4]
			xor		[esi],	eax
			
			inc		j
		.endw
		
		;	/* copy values into round key array */
		;	for(j = 0; (j < KC) && (r < (ROUNDS+1)); )
		;	{
		;		for (; (j < KC) && (t < 4); j++, t++)
		;			*((word32*)W[r][t]) = *((word32*)tk[j]);
		;		if (t == 4)
		;		{
		;			r++;
		;			t = 0;
		;		}
		;	}
		xor		ecx,	ecx
		.while ((ecx < KC) && (r < edx))
			.while ((ecx < KC) && (t < 4))
				mov		esi,	IE(tk, ecx)
				
				mov		eax,	r
				shl		eax,	2
				add		eax,	t
				
				mov		edi,	W
				shl		eax,	2
				add		edi,	eax
				
				m2m		[edi],	[esi]
				
				inc		ecx
				inc		t
			.endw
			.if (t == 4)
				inc		r
				mov		t,	0
			.endif
		.endw
	.endw
	
	pop		ebx	
	
	return(0)
rijndaelKeySched endp

;int rijndaelKeyEnctoDec (int keyBits, word8 W[MAXROUNDS+1][4][4])
;{
;	int r;
;
;	for (r = 1; r < ROUNDS; r++) {
;		InvMixColumn(W[r], 4);
;	}
;	return 0;
;}
; int rijndaelKeyEnctoDec (word8 W[MAXROUNDS+1][4][4])
rijndaelKeyEnctoDec proc W:DWORD
	push	ebx
	
	mov		ebx,	1
	.while (ebx < ROUNDS)
		
		mov		eax,	ebx
		shl		eax,	4
		add		eax,	W
		invoke	InvMixColumn, eax
		
		inc		ebx
	.endw
	pop		ebx
	
	return(0)
rijndaelKeyEnctoDec endp

;int makeKey(keyInstance *key, BYTE direction, int keyLen, char *keyMaterial)
;{
;	word8 k[MAXKC][4];
;	int i, j, t;
;
;	if ( keyMaterial )
;	{
;		strncpy(key->keyMaterial, keyMaterial, keyLen/4);
;	}
;
;	ROUNDS = keyLen/32 + 6;
;
;	/* initialize key schedule: */
;	for(i = 0; i < key->keyLen/8; i++)
;	{
;		t = key->keyMaterial[2*i];
;		// ....
;		t = key->keyMaterial[2*i+1];
;		// ....
;
;		k[i / 4][i % 4] = (word8) j;
;	}
;	rijndaelKeySched (k, key->keySched);
;	if (direction == DIR_DECRYPT)
;		rijndaelKeyEnctoDec (key->keyLen, key->keySched);
;
;	return TRUE;
;}

GetKeyField macro field
	push	ebx
	mov		ebx,	key
	assume	ebx : ptr keyInstance
	lea		eax,	[ebx].field
	assume	ebx : ptr nothing
	pop		ebx
endm

;int makeKey(keyInstance *key, BYTE direction, int keyLen, char *keyMaterial)
makeKey proc key:DWORD, direction:DWORD, keyLen:DWORD, keyMaterial:DWORD
LOCAL i : DWORD
LOCAL k[MAXKC] : DWORD
LOCAL pc : BYTE ; prev_char

	;	if ( keyMaterial )
	;		strncpy(key->keyMaterial, keyMaterial, keyLen/4);
	.if (keyMaterial)
		; в keyMaterial могут быть нулевые байты,
		; поэтому нельз€ копировать как sz
		GetKeyField(keyMaterial)
		mov		ecx,	keyLen
		shr		ecx,	2
		push	esi
		push	edi
		mov		esi,	keyMaterial
		mov		edi,	eax
		shr		ecx,	2
		cld
		rep movsd
		pop		edi
		pop		esi
		; invoke	strncpy, eax, keyMaterial, ecx
	.endif
	
	; ROUNDS = keyLen/32 + 6;
	mov		eax,	keyLen
	shr		eax,	5
	add		eax,	6
	mov		ROUNDS,	eax		; ROUNDS == 256/32 + 6 == 8 + 6 = 14 (!!!) всегда
	
	; for (i = 0; i < key->keyLen/8; i++)
	mov		i,		0
	mov		edx,	keyLen
	shr		edx,	3		; EDX == keyLen/8 = 256/8 = 32 (!!!) всегда
	
	mov		pc, 	10111001b ; вектор инициализации
	
	.while (i < edx)
		;	t = key->keyMaterial[2*i];
		GetKeyField(keyMaterial)
		mov		ecx,	i
		mov		ax,		word ptr [eax+2*ecx] ; сразу читаю и [2*i] и [2*i+1]
		
		xor		pc,		al
		xor		pc,		ah
		mov		al,		pc
		
		; k[i / 4][i % 4] = (word8) j;
		lea		ecx,	k
		add		ecx,	i
		mov byte ptr [ecx], al
		
		inc		i
	.endw
	
	;	rijndaelKeySched (k, key->keySched);
	;	if (direction == DIR_DECRYPT)
	;		rijndaelKeyEnctoDec (key->keyLen, key->keySched);
	GetKeyField(keySched)
	push	eax
	invoke	rijndaelKeySched, addr k, eax
	pop		eax
	.if (direction == DIR_DECRYPT)
		invoke	rijndaelKeyEnctoDec, eax
	.endif

	return(TRUE)
makeKey endp