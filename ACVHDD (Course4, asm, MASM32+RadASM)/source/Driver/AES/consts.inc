.const
	MAXKC			equ 256/32
	MAXROUNDS		equ 14
	
	DIR_ENCRYPT		equ 0		; Are we encrpyting?
	DIR_DECRYPT		equ 1		; Are we decrpyting?
	BITSPERBLOCK	equ 128		; Default number of bits in a cipher block
	BYTSPERBLOCK	equ BITSPERBLOCK/8 ; Default number of byts in a cipher block
	
	MAX_KEY_SIZE	equ 64		; # of ASCII char's needed to represent a key
	MAX_IV_SIZE		equ 32		; # bytes needed to represent an IV
	
; The structure for key information
keyInstance struct
	keyMaterial		BYTE MAX_KEY_SIZE+1 DUP (?) 		; Raw key data in ASCII, e.g., user input or KAT values
	keySched		BYTE (MAXROUNDS+1)*(4*4) DUP (?)	; key schedule
keyInstance ends

; The structure for cipher information
cipherInstance struct
	IV				BYTE MAX_IV_SIZE DUP (?)			; A possible Initialization Vector for ciphering
cipherInstance ends

.data?
	ROUNDS		DWORD	?
	
.code 