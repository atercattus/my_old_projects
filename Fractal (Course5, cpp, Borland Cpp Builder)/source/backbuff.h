#ifndef BACKBUFF_H
#define BACKBUFF_H

// INTERFACE

#include <windows.h>

DWORD backbuff_init( DWORD W, DWORD H, HDC _hDC );
void backbuff_done();
void backbuff_flip( HDC hDestDC );


// IMPLEMENTATION

HDC			hBackDC = NULL;
DWORD		buffWidth, buffHeight, lineDiff;
HBITMAP		hBitmap = NULL;
BITMAPINFO	bitmap_info;
BYTE		*lpColor = NULL;

DWORD backbuff_init( DWORD W, DWORD H, HDC _hDC ) {

	if (!(GetDeviceCaps(_hDC, RASTERCAPS) & RC_BITBLT))
	{
		// в качестве HDC передано что-то другое, либо система не поддерживает BitBlt
		return (DWORD)-1;
	}
	
	// создаю свой совместимый
	hBackDC = CreateCompatibleDC( _hDC );
	if (!hBackDC) return (DWORD)-2;

	// формирую BITMAPINFO:
	buffWidth = W;
	buffHeight = H;
//	lineDiff = (((3*buffWidth) + 3) & (~3)) - (3*buffWidth);

	memset(&bitmap_info, 0, sizeof(BITMAPINFO));

	bitmap_info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmap_info.bmiHeader.biWidth = buffWidth;
	bitmap_info.bmiHeader.biHeight = buffHeight;
	bitmap_info.bmiHeader.biPlanes = 1;
	bitmap_info.bmiHeader.biBitCount = 32;
	bitmap_info.bmiHeader.biCompression = BI_RGB;
	
	hBitmap = CreateDIBSection(hBackDC, &bitmap_info, DIB_RGB_COLORS, (void**)&lpColor, NULL, 0);
	if (!lpColor || !hBitmap)
	{
    	DeleteDC( hBackDC );
    	return (DWORD)-3;
	}
	
	SelectObject( hBackDC, hBitmap );

	return (DWORD)0;
}

void backbuff_done() {

	if (!hBackDC) return;

	DeleteObject( hBitmap );
	DeleteDC( hBackDC );
	hBackDC = NULL;
}

void backbuff_flip( HDC hDestDC ) {

	if (!hBackDC) return;

	BitBlt(hDestDC, 0, 0, buffWidth, buffHeight, hBackDC, 0, 0, SRCCOPY);
}

#endif // BACKBUFF_H