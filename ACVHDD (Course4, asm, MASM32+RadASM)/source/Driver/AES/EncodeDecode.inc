esi2edi macro
; 4DWORD из ESI копируются в EDI
; с "rep movsd" мне меньше нравится
	mov		eax,		[esi+0]
	mov		ecx,		[esi+4]
	mov		[edi+0],	eax
	mov		[edi+4],	ecx
	mov		eax,		[esi+8]
	mov		ecx,		[esi+12]
	mov		[edi+8],	eax
	mov		[edi+12],	ecx
endm

esiXedi macro
; [edi] ^= [esi] последовательно для 4DWORD
	mov		eax,		[esi+0]
	xor		[edi+0],	eax
	mov		eax,		[esi+4]
	xor		[edi+4],	eax
	mov		eax,		[esi+8]
	xor		[edi+8],	eax
	mov		eax,		[esi+12]
	xor		[edi+12],	eax
endm

GTV macro index, T
; Get Tx item Value
; index == r*4+c
; T in [T1, T2, T3, T4]
; temp in esi
	xor		eax,	eax
	mov		al,		[esi+index]
	mov		eax,	[offset T + eax*4]
endm

b_SetUp_Encode proc
; выполняет код, закоменченный ниже:
; необходимое состояние регистров: ESI == &temp, EDI == &b
; меняет: EAX, ECX
;	*((word32*)b)			=	*((word32*)T1[temp[0][0]])	^ *((word32*)T2[temp[1][1]]) ^
;								*((word32*)T3[temp[2][2]])	^ *((word32*)T4[temp[3][3]]);
;	*((word32*)(b+4))		=	*((word32*)T1[temp[1][0]])	^ *((word32*)T2[temp[2][1]]) ^
;								*((word32*)T3[temp[3][2]])	^ *((word32*)T4[temp[0][3]]);
;	*((word32*)(b+8))		=	*((word32*)T1[temp[2][0]])	^ *((word32*)T2[temp[3][1]]) ^
;								*((word32*)T3[temp[0][2]])	^ *((word32*)T4[temp[1][3]]);
;	*((word32*)(b+12))		=	*((word32*)T1[temp[3][0]])	^ *((word32*)T2[temp[0][1]]) ^
;								*((word32*)T3[temp[1][2]])	^ *((word32*)T4[temp[2][3]]);	

	;	*((word32*)b)	=	*((word32*)T1[temp[0][0]])	^ *((word32*)T2[temp[1][1]]) ^
	;						*((word32*)T3[temp[2][2]])	^ *((word32*)T4[temp[3][3]]);
	xor		ecx,		ecx
	GTV		0*4+0, T1
	xor		ecx,		eax
	GTV		1*4+1, T2
	xor		ecx,		eax
	GTV		2*4+2, T3
	xor		ecx,		eax
	GTV		3*4+3, T4
	xor		ecx,		eax
	mov		[edi+0],	ecx
	
	;	*((word32*)(b+4))		=	*((word32*)T1[temp[1][0]])	^ *((word32*)T2[temp[2][1]]) ^
	;								*((word32*)T3[temp[3][2]])	^ *((word32*)T4[temp[0][3]]);
	xor		ecx,		ecx
	GTV		1*4+0, T1
	xor		ecx,		eax
	GTV		2*4+1, T2
	xor		ecx,		eax
	GTV		3*4+2, T3
	xor		ecx,		eax
	GTV		0*4+3, T4
	xor		ecx,		eax
	mov		[edi+4],	ecx

;	*((word32*)(b+8))		=	*((word32*)T1[temp[2][0]])	^ *((word32*)T2[temp[3][1]]) ^
;								*((word32*)T3[temp[0][2]])	^ *((word32*)T4[temp[1][3]]);
	xor		ecx,		ecx
	GTV		2*4+0, T1
	xor		ecx,		eax
	GTV		3*4+1, T2
	xor		ecx,		eax
	GTV		0*4+2, T3
	xor		ecx,		eax
	GTV		1*4+3, T4
	xor		ecx,		eax
	mov		[edi+8],	ecx

;	*((word32*)(b+12))		=	*((word32*)T1[temp[3][0]])	^ *((word32*)T2[temp[0][1]]) ^
;								*((word32*)T3[temp[1][2]])	^ *((word32*)T4[temp[2][3]]);
	xor		ecx,		ecx
	GTV		3*4+0, T1
	xor		ecx,		eax
	GTV		0*4+1, T2
	xor		ecx,		eax
	GTV		1*4+2, T3
	xor		ecx,		eax
	GTV		2*4+3, T4
	xor		ecx,		eax
	mov		[edi+12],	ecx

	ret
b_SetUp_Encode endp

b_SetUp_Decode proc
; выполняет код, закоменченный ниже:
; необходимое состояние регистров: ESI == &temp, EDI == &b
; меняет: EAX, ECX
;	*((word32*)b)			=	*((word32*)T5[temp[0][0]])	^ *((word32*)T6[temp[3][1]]) ^
;								*((word32*)T7[temp[2][2]])	^ *((word32*)T8[temp[1][3]]);
;	*((word32*)(b+4))		=	*((word32*)T5[temp[1][0]])	^ *((word32*)T6[temp[0][1]]) ^
;								*((word32*)T7[temp[3][2]])	^ *((word32*)T8[temp[2][3]]);
;	*((word32*)(b+8))		=	*((word32*)T5[temp[2][0]])	^ *((word32*)T6[temp[1][1]]) ^
;								*((word32*)T7[temp[0][2]])	^ *((word32*)T8[temp[3][3]]);
;	*((word32*)(b+12))		=	*((word32*)T5[temp[3][0]])	^ *((word32*)T6[temp[2][1]]) ^
;								*((word32*)T7[temp[1][2]])	^ *((word32*)T8[temp[0][3]]);

;	*((word32*)b)			=	*((word32*)T5[temp[0][0]])	^ *((word32*)T6[temp[3][1]]) ^
;								*((word32*)T7[temp[2][2]])	^ *((word32*)T8[temp[1][3]]);
	xor		ecx,		ecx
	GTV		0*4+0, T5
	xor		ecx,		eax
	GTV		3*4+1, T6
	xor		ecx,		eax
	GTV		2*4+2, T7
	xor		ecx,		eax
	GTV		1*4+3, T8
	xor		ecx,		eax
	mov		[edi+0],	ecx
	
;	*((word32*)(b+4))		=	*((word32*)T5[temp[1][0]])	^ *((word32*)T6[temp[0][1]]) ^
;								*((word32*)T7[temp[3][2]])	^ *((word32*)T8[temp[2][3]]);
	xor		ecx,		ecx
	GTV		1*4+0, T5
	xor		ecx,		eax
	GTV		0*4+1, T6
	xor		ecx,		eax
	GTV		3*4+2, T7
	xor		ecx,		eax
	GTV		2*4+3, T8
	xor		ecx,		eax
	mov		[edi+4],	ecx

;	*((word32*)(b+8))		=	*((word32*)T5[temp[2][0]])	^ *((word32*)T6[temp[1][1]]) ^
;								*((word32*)T7[temp[0][2]])	^ *((word32*)T8[temp[3][3]]);
	xor		ecx,		ecx
	GTV		2*4+0, T5
	xor		ecx,		eax
	GTV		1*4+1, T6
	xor		ecx,		eax
	GTV		0*4+2, T7
	xor		ecx,		eax
	GTV		3*4+3, T8
	xor		ecx,		eax
	mov		[edi+8],	ecx

;	*((word32*)(b+12))		=	*((word32*)T5[temp[3][0]])	^ *((word32*)T6[temp[2][1]]) ^
;								*((word32*)T7[temp[1][2]])	^ *((word32*)T8[temp[0][3]]);
	xor		ecx,		ecx
	GTV		3*4+0, T5
	xor		ecx,		eax
	GTV		2*4+1, T6
	xor		ecx,		eax
	GTV		1*4+2, T7
	xor		ecx,		eax
	GTV		0*4+3, T8
	xor		ecx,		eax
	mov		[edi+12],	ecx

	ret
b_SetUp_Decode endp

libs_T1 macro x, y, i
; выполняет:
; b[i] = T1[temp[x][y]][1];
; требование к регистрам: esi == &b, edi == &temp
	xor		edx,		edx
	mov		dl,			[edi+x*4+y]
	mov		dl,			[T1 + edx*4 + 1]
	mov		[esi+i],	dl
endm

libs_S5 macro x, y, i
; выполняет:
; b[i] = S5[temp[x][y]];
; требование к регистрам: esi == &b, edi == &temp
	xor		edx,		edx
	mov		dl,			[edi+x*4+y]
	mov		dl,			[S5 + edx]
	mov		[esi+i],	dl
endm

.code

;{
;	int r;
;	word8 temp[4][4];
;
;	*((word32*)temp[0])		=	*((word32*)a)				^ *((word32*)rk[0][0]);
;	*((word32*)temp[1])		=	*((word32*)(a+4))			^ *((word32*)rk[0][1]);
;	*((word32*)temp[2])		=	*((word32*)(a+8))			^ *((word32*)rk[0][2]);
;	*((word32*)temp[3])		=	*((word32*)(a+12))			^ *((word32*)rk[0][3]);
;	*((word32*)b)			=	*((word32*)T1[temp[0][0]])	^ *((word32*)T2[temp[1][1]]) ^
;								*((word32*)T3[temp[2][2]])	^ *((word32*)T4[temp[3][3]]);
;	*((word32*)(b+4))		=	*((word32*)T1[temp[1][0]])	^ *((word32*)T2[temp[2][1]]) ^
;								*((word32*)T3[temp[3][2]])	^ *((word32*)T4[temp[0][3]]);
;	*((word32*)(b+8))		=	*((word32*)T1[temp[2][0]])	^ *((word32*)T2[temp[3][1]]) ^
;								*((word32*)T3[temp[0][2]])	^ *((word32*)T4[temp[1][3]]);
;	*((word32*)(b+12))		=	*((word32*)T1[temp[3][0]])	^ *((word32*)T2[temp[0][1]]) ^
;								*((word32*)T3[temp[1][2]])	^ *((word32*)T4[temp[2][3]]);
;								
;	for(r = 1; r < ROUNDS-1; r++)
;	{
;		*((word32*)temp[0])	=	*((word32*)b)				^ *((word32*)rk[r][0]);
;		*((word32*)temp[1])	=	*((word32*)(b+4))			^ *((word32*)rk[r][1]);
;		*((word32*)temp[2])	=	*((word32*)(b+8))			^ *((word32*)rk[r][2]);
;		*((word32*)temp[3])	=	*((word32*)(b+12))			^ *((word32*)rk[r][3]);
;		*((word32*)b)		=	*((word32*)T1[temp[0][0]])	^ *((word32*)T2[temp[1][1]]) ^
;								*((word32*)T3[temp[2][2]])	^ *((word32*)T4[temp[3][3]]);
;		*((word32*)(b+4))	=	*((word32*)T1[temp[1][0]])	^ *((word32*)T2[temp[2][1]]) ^
;								*((word32*)T3[temp[3][2]])	^ *((word32*)T4[temp[0][3]]);
;		*((word32*)(b+8))	=	*((word32*)T1[temp[2][0]])	^ *((word32*)T2[temp[3][1]]) ^
;								*((word32*)T3[temp[0][2]])	^ *((word32*)T4[temp[1][3]]);
;		*((word32*)(b+12))	=	*((word32*)T1[temp[3][0]])	^ *((word32*)T2[temp[0][1]]) ^
;								*((word32*)T3[temp[1][2]])	^ *((word32*)T4[temp[2][3]]);
;	}
;	
;   /* last round is special */
;	*((word32*)temp[0])		=	*((word32*)b)				^ *((word32*)rk[ROUNDS-1][0]);
;	*((word32*)temp[1])		=	*((word32*)(b+4))			^ *((word32*)rk[ROUNDS-1][1]);
;	*((word32*)temp[2])		=	*((word32*)(b+8))			^ *((word32*)rk[ROUNDS-1][2]);
;	*((word32*)temp[3])		=	*((word32*)(b+12))			^ *((word32*)rk[ROUNDS-1][3]);
;
;	b[0]	= T1[temp[0][0]][1];
;	b[1]	= T1[temp[1][1]][1];
;	b[2]	= T1[temp[2][2]][1];
;	b[3]	= T1[temp[3][3]][1];
;	b[4]	= T1[temp[1][0]][1];
;	b[5]	= T1[temp[2][1]][1];
;	b[6]	= T1[temp[3][2]][1];
;	b[7]	= T1[temp[0][3]][1];
;	b[8]	= T1[temp[2][0]][1];
;	b[9]	= T1[temp[3][1]][1];
;	b[10]	= T1[temp[0][2]][1];
;	b[11]	= T1[temp[1][3]][1];
;	b[12]	= T1[temp[3][0]][1];
;	b[13]	= T1[temp[0][1]][1];
;	b[14]	= T1[temp[1][2]][1];
;	b[15]	= T1[temp[2][3]][1];
;	*((word32*)b)		^=	*((word32*)rk[ROUNDS][0]);
;	*((word32*)(b+4))	^=	*((word32*)rk[ROUNDS][1]);
;	*((word32*)(b+8))	^=	*((word32*)rk[ROUNDS][2]);
;	*((word32*)(b+12))	^=	*((word32*)rk[ROUNDS][3]);
;}

;int rijndaelEncrypt (word8 a[16], word8 b[16], word8 rk[MAXROUNDS+1][4][4])
rijndaelEncrypt proc uses eax ecx edx a:DWORD, b:DWORD, rk:DWORD
LOCAL	temp[4] : DWORD

	push	ebx
	push	esi
	push	edi

	;	*((word32*)temp[0])	=	*((word32*)a)		^ *((word32*)rk[0][0]);
	;	*((word32*)temp[1])	=	*((word32*)(a+4))	^ *((word32*)rk[0][1]);
	;	*((word32*)temp[2])	=	*((word32*)(a+8))	^ *((word32*)rk[0][2]);
	;	*((word32*)temp[3])	=	*((word32*)(a+12))	^ *((word32*)rk[0][3]);
	
	mov		esi,		a
	lea		edi,		temp		; EDI = &temp
	esi2edi
	mov		esi,		rk			; ESI = &rk
	esiXedi
	
	xchg	esi,	edi				; ESI == &temp
	mov		edi,	b				; EDI == &b
	
	;	*((word32*)b)			=	*((word32*)T1[temp[0][0]])	^ *((word32*)T2[temp[1][1]]) ^
	;								*((word32*)T3[temp[2][2]])	^ *((word32*)T4[temp[3][3]]);
	;	*((word32*)(b+4))		=	*((word32*)T1[temp[1][0]])	^ *((word32*)T2[temp[2][1]]) ^
	;								*((word32*)T3[temp[3][2]])	^ *((word32*)T4[temp[0][3]]);
	;	*((word32*)(b+8))		=	*((word32*)T1[temp[2][0]])	^ *((word32*)T2[temp[3][1]]) ^
	;								*((word32*)T3[temp[0][2]])	^ *((word32*)T4[temp[1][3]]);
	;	*((word32*)(b+12))		=	*((word32*)T1[temp[3][0]])	^ *((word32*)T2[temp[0][1]]) ^
	;								*((word32*)T3[temp[1][2]])	^ *((word32*)T4[temp[2][3]]);
	invoke	b_SetUp_Encode
	
	mov		ebx,	1			; EBX == r
	mov		edx,	ROUNDS
	dec		edx					; EDX == ROUNDS-1

	; esi == &temp
	; edi == &b
	; ebx == r
	; edx == ROUNDS-1
	
	; for (r = 1; r < ROUNDS-1; r++)
	.while (ebx < edx)
		;		*((word32*)temp[0])	=	*((word32*)b)				^ *((word32*)rk[r][0]);
		;		*((word32*)temp[1])	=	*((word32*)(b+4))			^ *((word32*)rk[r][1]);
		;		*((word32*)temp[2])	=	*((word32*)(b+8))			^ *((word32*)rk[r][2]);
		;		*((word32*)temp[3])	=	*((word32*)(b+12))			^ *((word32*)rk[r][3]);
		xchg	esi,	edi
		esi2edi
		mov		ecx,	esi		; сохраняю ESI
		mov		esi,	rk
		mov		eax,	ebx
		shl		eax,	4
		add		esi,	eax
		esiXedi
		mov		esi,	ecx		; восстанавливаю ESI
		xchg	esi,	edi
		
		;	*((word32*)b)			=	*((word32*)T1[temp[0][0]])	^ *((word32*)T2[temp[1][1]]) ^
		;								*((word32*)T3[temp[2][2]])	^ *((word32*)T4[temp[3][3]]);
		;	*((word32*)(b+4))		=	*((word32*)T1[temp[1][0]])	^ *((word32*)T2[temp[2][1]]) ^
		;								*((word32*)T3[temp[3][2]])	^ *((word32*)T4[temp[0][3]]);
		;	*((word32*)(b+8))		=	*((word32*)T1[temp[2][0]])	^ *((word32*)T2[temp[3][1]]) ^
		;								*((word32*)T3[temp[0][2]])	^ *((word32*)T4[temp[1][3]]);
		;	*((word32*)(b+12))		=	*((word32*)T1[temp[3][0]])	^ *((word32*)T2[temp[0][1]]) ^
		;								*((word32*)T3[temp[1][2]])	^ *((word32*)T4[temp[2][3]]);	
		invoke	b_SetUp_Encode
		
		inc		ebx
	.endw

	; esi == &temp
	; edi == &b
	; edx == ROUNDS-1
	
	;   /* last round is special */
	;	*((word32*)temp[0])		=	*((word32*)b)				^ *((word32*)rk[ROUNDS-1][0]);
	;	*((word32*)temp[1])		=	*((word32*)(b+4))			^ *((word32*)rk[ROUNDS-1][1]);
	;	*((word32*)temp[2])		=	*((word32*)(b+8))			^ *((word32*)rk[ROUNDS-1][2]);
	;	*((word32*)temp[3])		=	*((word32*)(b+12))			^ *((word32*)rk[ROUNDS-1][3]);

	xchg	esi,	edi
	esi2edi
	mov		ecx,	esi			; save ESI
	mov		esi,	rk
	mov		eax,	edx
	shl		eax,	4
	add		esi,	eax
	esiXedi
	mov		esi,	ecx			; rest ESI
	
	; esi == &b
	; edi == &temp

	;	b[0]	= T1[temp[0][0]][1];
	;	b[1]	= T1[temp[1][1]][1];
	;	b[2]	= T1[temp[2][2]][1];
	;	b[3]	= T1[temp[3][3]][1];
	libs_T1 0, 0, 0
	libs_T1 1, 1, 1
	libs_T1 2, 2, 2
	libs_T1 3, 3, 3
	;	b[4]	= T1[temp[1][0]][1];
	;	b[5]	= T1[temp[2][1]][1];
	;	b[6]	= T1[temp[3][2]][1];
	;	b[7]	= T1[temp[0][3]][1];
	libs_T1 1, 0, 4
	libs_T1 2, 1, 5
	libs_T1 3, 2, 6
	libs_T1 0, 3, 7
	;	b[8]	= T1[temp[2][0]][1];
	;	b[9]	= T1[temp[3][1]][1];
	;	b[10]	= T1[temp[0][2]][1];
	;	b[11]	= T1[temp[1][3]][1];
	libs_T1 2, 0, 8
	libs_T1 3, 1, 9
	libs_T1 0, 2, 10
	libs_T1 1, 3, 11
	;	b[12]	= T1[temp[3][0]][1];
	;	b[13]	= T1[temp[0][1]][1];
	;	b[14]	= T1[temp[1][2]][1];
	;	b[15]	= T1[temp[2][3]][1];
	libs_T1 3, 0, 12
	libs_T1 0, 1, 13
	libs_T1 1, 2, 14
	libs_T1 2, 3, 15
	
	;	*((word32*)b)		^=	*((word32*)rk[ROUNDS][0]);
	;	*((word32*)(b+4))	^=	*((word32*)rk[ROUNDS][1]);
	;	*((word32*)(b+8))	^=	*((word32*)rk[ROUNDS][2]);
	;	*((word32*)(b+12))	^=	*((word32*)rk[ROUNDS][3]);
	mov		edi,	ROUNDS
	shl		edi,	4
	add		edi,	rk
	xchg	esi,	edi
	esiXedi
	
	pop		edi
	pop		esi
	pop		ebx
	ret
rijndaelEncrypt endp

;int rijndaelDecrypt (word8 a[16], word8 b[16], word8 rk[MAXROUNDS+1][4][4])
;{
;	int r;
;   word8 temp[4][4];
;	
;
;	*((word32*)temp[0])		=	*((word32*)a)				^ *((word32*)rk[ROUNDS][0]);
;	*((word32*)temp[1])		=	*((word32*)(a+4))			^ *((word32*)rk[ROUNDS][1]);
;	*((word32*)temp[2])		=	*((word32*)(a+8))			^ *((word32*)rk[ROUNDS][2]);
;	*((word32*)temp[3])		=	*((word32*)(a+12))			^ *((word32*)rk[ROUNDS][3]);
;
;	*((word32*)b)			=	*((word32*)T5[temp[0][0]])	^ *((word32*)T6[temp[3][1]]) ^
;								*((word32*)T7[temp[2][2]])	^ *((word32*)T8[temp[1][3]]);
;	*((word32*)(b+4))		=	*((word32*)T5[temp[1][0]])	^ *((word32*)T6[temp[0][1]]) ^
;								*((word32*)T7[temp[3][2]])	^ *((word32*)T8[temp[2][3]]);
;	*((word32*)(b+8))		=	*((word32*)T5[temp[2][0]])	^ *((word32*)T6[temp[1][1]]) ^
;								*((word32*)T7[temp[0][2]])	^ *((word32*)T8[temp[3][3]]);
;	*((word32*)(b+12))		=	*((word32*)T5[temp[3][0]])	^ *((word32*)T6[temp[2][1]]) ^
;								*((word32*)T7[temp[1][2]])	^ *((word32*)T8[temp[0][3]]);
;
;	for(r = ROUNDS-1; r > 1; r--)
;	{
;		*((word32*)temp[0])	=	*((word32*)b)				^ *((word32*)rk[r][0]);
;		*((word32*)temp[1])	=	*((word32*)(b+4))			^ *((word32*)rk[r][1]);
;		*((word32*)temp[2])	=	*((word32*)(b+8))			^ *((word32*)rk[r][2]);
;		*((word32*)temp[3])	=	*((word32*)(b+12))			^ *((word32*)rk[r][3]);
;
;		*((word32*)b)			=	*((word32*)T5[temp[0][0]])	^ *((word32*)T6[temp[3][1]]) ^
;									*((word32*)T7[temp[2][2]])	^ *((word32*)T8[temp[1][3]]);
;		*((word32*)(b+4))		=	*((word32*)T5[temp[1][0]])	^ *((word32*)T6[temp[0][1]]) ^
;									*((word32*)T7[temp[3][2]])	^ *((word32*)T8[temp[2][3]]);
;		*((word32*)(b+8))		=	*((word32*)T5[temp[2][0]])	^ *((word32*)T6[temp[1][1]]) ^
;									*((word32*)T7[temp[0][2]])	^ *((word32*)T8[temp[3][3]]);
;		*((word32*)(b+12))		=	*((word32*)T5[temp[3][0]])	^ *((word32*)T6[temp[2][1]]) ^
;									*((word32*)T7[temp[1][2]])	^ *((word32*)T8[temp[0][3]]);
;	}
;	/* last round is special */
;
;	*((word32*)temp[0])		=	*((word32*)b)				^ *((word32*)rk[1][0]);
;	*((word32*)temp[1])		=	*((word32*)(b+4))			^ *((word32*)rk[1][1]);
;	*((word32*)temp[2])		=	*((word32*)(b+8))			^ *((word32*)rk[1][2]);
;	*((word32*)temp[3])		=	*((word32*)(b+12))			^ *((word32*)rk[1][3]);
;
;	b[0]	= S5[temp[0][0]];
;	b[1]	= S5[temp[3][1]];
;	b[2]	= S5[temp[2][2]];
;	b[3]	= S5[temp[1][3]];
;	b[4]	= S5[temp[1][0]];
;	b[5]	= S5[temp[0][1]];
;	b[6]	= S5[temp[3][2]];
;	b[7]	= S5[temp[2][3]];
;	b[8]	= S5[temp[2][0]];
;	b[9]	= S5[temp[1][1]];
;	b[10]	= S5[temp[0][2]];
;	b[11]	= S5[temp[3][3]];
;	b[12]	= S5[temp[3][0]];
;	b[13]	= S5[temp[2][1]];
;	b[14]	= S5[temp[1][2]];
;	b[15]	= S5[temp[0][3]];
;	*((word32*)b)		^=	*((word32*)rk[0][0]);
;	*((word32*)(b+4))	^=	*((word32*)rk[0][1]);
;	*((word32*)(b+8))	^=	*((word32*)rk[0][2]);
;	*((word32*)(b+12))	^=	*((word32*)rk[0][3]);
;
;	return 0;
;}

;int rijndaelDecrypt (word8 a[16], word8 b[16], word8 rk[MAXROUNDS+1][4][4])
rijndaelDecrypt proc a:DWORD, b:DWORD, rk:DWORD
LOCAL	temp[4] : DWORD
	
	push	ebx
	push	esi
	push	edi
	push	edx
	
	;	*((word32*)temp[0])		=	*((word32*)a)				^ *((word32*)rk[ROUNDS][0]);
	;	*((word32*)temp[1])		=	*((word32*)(a+4))			^ *((word32*)rk[ROUNDS][1]);
	;	*((word32*)temp[2])		=	*((word32*)(a+8))			^ *((word32*)rk[ROUNDS][2]);
	;	*((word32*)temp[3])		=	*((word32*)(a+12))			^ *((word32*)rk[ROUNDS][3]);
	mov		esi,		a
	lea		edi,		temp
	esi2edi
	mov		esi,		ROUNDS
	shl		esi,		4
	add		esi,		rk
	esiXedi
	
	xchg	esi,	edi				; ESI == &temp
	mov		edi,	b				; EDI == &b

	;	*((word32*)b)			=	*((word32*)T5[temp[0][0]])	^ *((word32*)T6[temp[3][1]]) ^
	;								*((word32*)T7[temp[2][2]])	^ *((word32*)T8[temp[1][3]]);
	;	*((word32*)(b+4))		=	*((word32*)T5[temp[1][0]])	^ *((word32*)T6[temp[0][1]]) ^
	;								*((word32*)T7[temp[3][2]])	^ *((word32*)T8[temp[2][3]]);
	;	*((word32*)(b+8))		=	*((word32*)T5[temp[2][0]])	^ *((word32*)T6[temp[1][1]]) ^
	;								*((word32*)T7[temp[0][2]])	^ *((word32*)T8[temp[3][3]]);
	;	*((word32*)(b+12))		=	*((word32*)T5[temp[3][0]])	^ *((word32*)T6[temp[2][1]]) ^
	;								*((word32*)T7[temp[1][2]])	^ *((word32*)T8[temp[0][3]]);
	invoke	b_SetUp_Decode
	
	mov		ebx,	ROUNDS
	dec		ebx					; EBX == ROUNDS-1
	
	; esi == &temp
	; edi == &b
	; ebx == r
	
	; for (r = ROUNDS-1; r > 1; r--)
	.while (ebx > 1)
		;		*((word32*)temp[0])	=	*((word32*)b)				^ *((word32*)rk[r][0]);
		;		*((word32*)temp[1])	=	*((word32*)(b+4))			^ *((word32*)rk[r][1]);
		;		*((word32*)temp[2])	=	*((word32*)(b+8))			^ *((word32*)rk[r][2]);
		;		*((word32*)temp[3])	=	*((word32*)(b+12))			^ *((word32*)rk[r][3]);
		xchg	esi,	edi
		esi2edi
		mov		ecx,	esi		; сохраняю ESI
		mov		esi,	rk
		mov		eax,	ebx
		shl		eax,	4
		add		esi,	eax
		esiXedi
		mov		esi,	ecx		; восстанавливаю ESI
		xchg	esi,	edi
		
		;		*((word32*)b)			=	*((word32*)T5[temp[0][0]])	^ *((word32*)T6[temp[3][1]]) ^
		;									*((word32*)T7[temp[2][2]])	^ *((word32*)T8[temp[1][3]]);
		;		*((word32*)(b+4))		=	*((word32*)T5[temp[1][0]])	^ *((word32*)T6[temp[0][1]]) ^
		;									*((word32*)T7[temp[3][2]])	^ *((word32*)T8[temp[2][3]]);
		;		*((word32*)(b+8))		=	*((word32*)T5[temp[2][0]])	^ *((word32*)T6[temp[1][1]]) ^
		;									*((word32*)T7[temp[0][2]])	^ *((word32*)T8[temp[3][3]]);
		;		*((word32*)(b+12))		=	*((word32*)T5[temp[3][0]])	^ *((word32*)T6[temp[2][1]]) ^
		;									*((word32*)T7[temp[1][2]])	^ *((word32*)T8[temp[0][3]]);
		invoke	b_SetUp_Decode
		
		dec		ebx
	.endw
	
	; esi == &temp
	; edi == &b

	;	/* last round is special */
	;	*((word32*)temp[0])		=	*((word32*)b)				^ *((word32*)rk[1][0]);
	;	*((word32*)temp[1])		=	*((word32*)(b+4))			^ *((word32*)rk[1][1]);
	;	*((word32*)temp[2])		=	*((word32*)(b+8))			^ *((word32*)rk[1][2]);
	;	*((word32*)temp[3])		=	*((word32*)(b+12))			^ *((word32*)rk[1][3]);
	
	xchg	esi,	edi
	esi2edi
	mov		ecx,	esi			; save ESI
	mov		esi,	rk
	add		esi,	1*4*4		; &rk[1]
	esiXedi
	mov		esi,	ecx			; rest ESI
	
	; esi == &b
	; edi == &temp
	
	;	b[0]	= S5[temp[0][0]];
	;	b[1]	= S5[temp[3][1]];
	;	b[2]	= S5[temp[2][2]];
	;	b[3]	= S5[temp[1][3]];
	libs_S5 0, 0, 0
	libs_S5 3, 1, 1
	libs_S5 2, 2, 2
	libs_S5 1, 3, 3
	;	b[4]	= S5[temp[1][0]];
	;	b[5]	= S5[temp[0][1]];
	;	b[6]	= S5[temp[3][2]];
	;	b[7]	= S5[temp[2][3]];
	libs_S5 1, 0, 4
	libs_S5 0, 1, 5
	libs_S5 3, 2, 6
	libs_S5 2, 3, 7
	;	b[8]	= S5[temp[2][0]];
	;	b[9]	= S5[temp[1][1]];
	;	b[10]	= S5[temp[0][2]];
	;	b[11]	= S5[temp[3][3]];
	libs_S5 2, 0, 8
	libs_S5 1, 1, 9
	libs_S5 0, 2, 10
	libs_S5 3, 3, 11
	;	b[12]	= S5[temp[3][0]];
	;	b[13]	= S5[temp[2][1]];
	;	b[14]	= S5[temp[1][2]];
	;	b[15]	= S5[temp[0][3]];
	libs_S5 3, 0, 12
	libs_S5 2, 1, 13
	libs_S5 1, 2, 14
	libs_S5 0, 3, 15
	
	;	*((word32*)b)		^=	*((word32*)rk[0][0]);
	;	*((word32*)(b+4))	^=	*((word32*)rk[0][1]);
	;	*((word32*)(b+8))	^=	*((word32*)rk[0][2]);
	;	*((word32*)(b+12))	^=	*((word32*)rk[0][3]);
	mov		edi,	rk
	xchg	esi,	edi
	esiXedi

	pop		edx
	pop		edi
	pop		esi
	pop		ebx
	ret
rijndaelDecrypt endp

.code