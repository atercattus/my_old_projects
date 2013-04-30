#ifndef PALETTE_H
#define PALETTE_H

#include <windows.h>

#define MAX_ITER 2000
#define MIN_ITER 10
#define STD_ITER 100

// INTERFACE
char* DEF_PALETTE_FILE = "default.pal";
COLORREF palette[MAX_ITER];

bool SavePalette(char *path);
bool LoadPalette(char *path);
void GenerateRandomPalette( bool full_rand, DWORD from );
void GenerateDefaultPalette();
void SetCurMaxIter( DWORD _ );

extern HWND hMainWnd;

// IMPLEMENTATION

// Ќ≈Ћ№«я ћ≈Ќя“№ Ќ≈ѕќ—–≈ƒ—“¬≈ЌЌќ!
DWORD CUR_MAX_ITER = STD_ITER;

// сейчас используетс€ стандартна€ палитра?
bool CurUsingDefPalette = true;
// сейв текущего full_rand
bool saved_full_rand;
// число уже заполненных в случайной палитре €чеек
// при увеличении после уменьшени€ не генерит то, что уже и так есть
DWORD how_much_used = 0;

bool SavePalette(char *path) {
	HANDLE hFile = CreateFile( path, FILE_WRITE_DATA, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if ( hFile == INVALID_HANDLE_VALUE ) {
		return false;
	}
	
	DWORD nw;
	WriteFile( hFile, palette, MAX_ITER*sizeof(COLORREF), &nw, NULL );
	
	CloseHandle( hFile );
	
	char buff[MAX_PATH+100];
	wsprintf( buff, "ѕалитра сохранена в файл %s.", path );
	MessageBox( hMainWnd, buff, NULL, MB_ICONINFORMATION );
	
	return true;
}

bool LoadPalette(char *path) {
	HANDLE hFile = CreateFile( path, FILE_READ_DATA, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if ( hFile == INVALID_HANDLE_VALUE ) {
		return false;
	}
	
	DWORD nw;
//	DWORD cmi;
//	ReadFile( hFile, &cmi, sizeof(CUR_MAX_ITER), &nw, NULL );
//	SetCurMaxIter( cmi );
	ReadFile( hFile, palette, MAX_ITER*sizeof(COLORREF), &nw, NULL );
	
	CloseHandle( hFile );
	
	char buff[MAX_PATH+100];
	wsprintf( buff, "ѕалитра загружена из файла %s.", path );
	MessageBox( hMainWnd, buff, NULL, MB_ICONINFORMATION );
	
	return true;
}

_declspec(naked) DWORD mrand() {
	__asm RDTSC;
	__asm xor eax, edx;
	__asm ret;
}

_declspec(naked) DWORD EAXrand() {
	__asm RDTSC;
	__asm ret;
}

_declspec(naked) DWORD EDXrand() {
	__asm RDTSC;
	__asm xchg eax, edx;
	__asm ret;
}

void GenerateRandomPalette( bool full_rand, DWORD from ) {
	for (int i=from; i<CUR_MAX_ITER; i++) {
		if ( full_rand )	palette[i] = mrand();
		else				palette[i] = RGB( mrand() % 256, mrand() % 256, mrand() % 256 );
	}
	
	saved_full_rand = full_rand;
	how_much_used = CUR_MAX_ITER;
	CurUsingDefPalette = false;
}

void GenerateDefaultPalette() {
	int	addR = 15 << 16,
		addG = 4 << 8,
		addB = -2,
		r = 50 << 16,
		g = 50 << 8,
		b = 50;
		
	for (int i=0; i<CUR_MAX_ITER; i++) {
		palette[i] = r | g | b;
		r += addR, g += addG, b += addB;
		if (r < 0)
			addR = -addR, r = -r;
		else if (r > (0xF0 << 16))
			addR = -addR, r += addR;
		if (g < 0)
			addG = -addG, g = -g;
		else if (g > (0xF0 << 8))
			addG = -addG, g += addG;
		if (b < 0)
			addB = -addB, b = -b;
		else if (b > 0xF0)
			addB = -addB, b += addB;
	}
	
	CurUsingDefPalette = true;
	how_much_used = 0;
}

void SetCurMaxIter( DWORD _ ) {

	DWORD _CUR_MAX_ITER = CUR_MAX_ITER;

	CUR_MAX_ITER = (_ > MAX_ITER) ? MAX_ITER : (_ < MIN_ITER) ? MIN_ITER : _;
	
	if ( CurUsingDefPalette ) {
		GenerateDefaultPalette();
	}
	else {
		if (_CUR_MAX_ITER < CUR_MAX_ITER) {	// увеличение
			
			if ( how_much_used < CUR_MAX_ITER ) {
				// не вс€ —≈…„ј— используема€ часть палитры инициализированна
				GenerateRandomPalette( saved_full_rand, how_much_used );
				how_much_used = CUR_MAX_ITER;
			}
			else {
				// увеличение в пределах уже созданной палитры
			}
		}
		else {
			// уменьшение с сохранением палитры
		}
	}
}

#endif // PALETTE_H