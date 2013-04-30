#include <windows.h>
#include "backbuff.h"
#include "palette.h"

char* szTitle = "Лаба номер 3";
const char* szWindowClass = "Laba3Class";

const double SCALE_DIFF = 0.5;
const double ITER_DIFF = 10;
const double SHIFT_DIFF = 0.5;

// область вывода фрактала
double vr_x, vr_y, vr_w, vr_h;

#define VR_X	-2
#define VR_Y	-2
#define VR_W	4
#define VR_H	4

const int WM_MOUSEWHEEL = 0x020A;

HINSTANCE hInstance;
HWND	hMainWnd;

HCURSOR	hCursorWait;
HCURSOR	hCursorArrow;

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
void paintMandelbrot_Slow(HDC hDC, int w, int h, double dx, double dy, double left, double first);
void d2a(double x, char* a);

void main() {
	hInstance = GetModuleHandle( NULL );
	
	GenerateDefaultPalette();
	
//	char buff[100];
//	d2a( -123.456, buff );
	
	hCursorWait = LoadCursor(NULL, IDC_WAIT);
	hCursorArrow = LoadCursor(NULL, IDC_ARROW);
	
	WNDCLASSEX wc;
	wc.cbSize			= sizeof(WNDCLASSEX);
	wc.style			= 0;
	wc.lpfnWndProc		= WndProc;
	wc.cbClsExtra		= 0;
	wc.cbWndExtra		= 0;
	wc.hInstance		= hInstance;
	wc.hIcon			= LoadIcon( hInstance, IDI_APPLICATION );
	wc.hCursor			= LoadCursor( NULL, IDC_ARROW );
	wc.hbrBackground	= 0;
	wc.lpszMenuName		= NULL;
	wc.lpszClassName	= szWindowClass;
	wc.hIconSm			= wc.hIcon;
	RegisterClassEx(&wc);
	
	hMainWnd = CreateWindow(	szWindowClass, szTitle, WS_OVERLAPPEDWINDOW | WS_VISIBLE,
								CW_USEDEFAULT, CW_USEDEFAULT, 600, 600,
								NULL, NULL, hInstance, NULL );
								
	SendMessage( hMainWnd, WM_KEYDOWN, VK_F1, 0 );

	MSG msg;
	while ( GetMessage( &msg, NULL, 0, 0 ) )
	{
		TranslateMessage( &msg );
		DispatchMessage( &msg );
	}
}

// for DEBUG
int __stdcall WinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	main();
}

void paintFractal( HWND hWnd ) {
	RECT rcSize;
	GetClientRect( hWnd, &rcSize );
	
	SetCursor(hCursorWait);
	
	DWORD gtc = GetTickCount();
	
	paintMandelbrot_Slow(	hBackDC, rcSize.right, rcSize.bottom,
							vr_w, vr_h, vr_x, vr_y
							);
							
	gtc = GetTickCount() - gtc;
/*	
	char v1[20], v2[20], v3[20], v4[20];
	d2a( vr_x, (char*)v1 );
	d2a( vr_y, (char*)v2 );
	d2a( vr_w, (char*)v3 );
	d2a( vr_h, (char*)v4 );
*/	
	char buff[1000];
//	wsprintf( buff, "X=%s Y=%s W=%s H=%s : %i мс", v1, v2, v3, v4, gtc );
	wsprintf( buff, "Просчет за %i мс. Окно (%ix%i). Число итераций %i", gtc, rcSize.right, rcSize.bottom, CUR_MAX_ITER );
	SetWindowText( hWnd, buff );
	
	SetCursor(hCursorArrow);
}

void ScaleToXY( int x, int y, int w, int h, bool ZoomIn, bool OnlyShift ) {

	if ( ZoomIn ) {
		vr_x = ( double(x)   * vr_w / double(w) ) + vr_x - vr_w/2.0;
		vr_y = ( double(h-y) * vr_h / double(h) ) + vr_y - vr_h/2.0;
	}
	
	if ( !OnlyShift ) {
		double _vr_w = vr_w;
		double _vr_h = vr_h;
		
		int dS = (ZoomIn) ? 1 : -1; // delta scale
		
		vr_w *= (1 - dS*SCALE_DIFF);
		vr_h *= (1 - dS*SCALE_DIFF);
		
		vr_x += (_vr_w - vr_w) / 2.0;
		vr_y += (_vr_h - vr_h) / 2.0;
		
		SetCurMaxIter( CUR_MAX_ITER + dS*ITER_DIFF );
	}
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {

	switch (message) {
		case WM_PAINT:
			{
				PAINTSTRUCT ps;
				HDC hDC = BeginPaint( hWnd, &ps );
				
				backbuff_flip( hDC );
				
				EndPaint( hWnd, &ps );
				break;
			}
			
		case WM_SIZE:
			{
				backbuff_done();
				HDC hDC = GetDC( 0 );
				backbuff_init( LOWORD(lParam), HIWORD(lParam), hDC );
				ReleaseDC( 0, hDC );
				
				paintFractal( hWnd );
				InvalidateRect(hWnd, NULL, false);
				
				break;
			}
		case WM_MOUSEWHEEL:
			{
				switch ( (int(wParam) >> 16) > 0 ) {
					case true:
						{
							SendMessage( hWnd, WM_LBUTTONDOWN, wParam, lParam);
							break;
						}
					case false:
						{
							SendMessage( hWnd, WM_RBUTTONDOWN, wParam, lParam);
							break;
						}
				}
				break;
			}
		case WM_LBUTTONDOWN:
			{
				RECT rcSize;
				GetClientRect( hWnd, &rcSize );
				int x = LOWORD(lParam);
				int y = HIWORD(lParam);
				
				ScaleToXY( x, y, rcSize.right, rcSize.bottom, true, (wParam & MK_CONTROL) );
				
				paintFractal( hWnd );
				InvalidateRect(hWnd, NULL, false);
				break;
			}
			
		case WM_RBUTTONDOWN:
			{
				ScaleToXY( 0, 0, 0, 0, false, false );
				
				paintFractal( hWnd );
				InvalidateRect(hWnd, NULL, false);
				break;
			}

		case WM_KEYDOWN:
			{
				switch (wParam) {
					case VK_ADD:
						{
							RECT rcSize;
							GetClientRect( hWnd, &rcSize );
							int x = rcSize.right  / 2;
							int y = rcSize.bottom / 2;
							
							ScaleToXY( x, y, rcSize.right, rcSize.bottom, true, false );
							
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_SUBTRACT:
						{
							ScaleToXY( 0, 0, 0, 0, false, false );

							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_LEFT:
						{
							vr_x -= SHIFT_DIFF * (vr_w / 4);
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_RIGHT:
						{
							vr_x += SHIFT_DIFF * (vr_w / 4);
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_UP:
						{
							vr_y += SHIFT_DIFF * (vr_h / 4);
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_DOWN:
						{
							vr_y -= SHIFT_DIFF * (vr_h / 4);
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_BACK:
						{
							vr_x = VR_X, vr_y = VR_Y,
							vr_w = VR_W, vr_h = VR_H;
							SetCurMaxIter( STD_ITER );
							
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_F8:
						{
							GenerateDefaultPalette();
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_F11:
					case VK_F12:
						{
							GenerateRandomPalette( (wParam == VK_F11 ), 0 );
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_ESCAPE:
						{
							SendMessage( hWnd, WM_CLOSE, 0, 0 );
							break;
						}
/*
					case 'S':
						{
							char buff[20];
							wsprintf( buff, "%X%X.pal", EAXrand(), EDXrand() );
							SavePalette( buff );
							break;
						}
*/
					case VK_PRIOR: // PageUp
						{
							SetCurMaxIter( CUR_MAX_ITER + ITER_DIFF );
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_NEXT: // PageDown
						{
							SetCurMaxIter( CUR_MAX_ITER - ITER_DIFF );
							paintFractal( hWnd );
							InvalidateRect(hWnd, NULL, false);
							break;
						}
					case VK_F1:
						{
							char *help =	"Левая кнопка мыши - увеличение в точку под курсором (при нажатом Control - сдвиг без масштаба),\n"\
											"Правая кнопка мыши - уменьшение от центра,\n"\
											"Колесо мыши вверх - аналог левой кнопки мыши,\n"\
											"Колесо мыши вниз - аналог правой кнопки мыши,\n"\
											"Кнопка [+] - увеличение в центр,\n"\
											"Кнопка [-] - уменьшение от центра,\n"\
											"Стрелки - перемещение без масштабирования,\n"\
											"Кнопка [Backspace] - сброс настроек в начальное состояние (кроме палитры),\n"\
											"Кнопки [PageUp] и [PageDown] - увеличение и уменьшение текущего числа итераций,\n"\
											"Кнопка [F8] - инициализация палитры стандартными значениями,\n"\
											"Кнопка [F11] - случайное заполнение палитры единым RGB значением,\n"\
											"Кнопка [F12] - случайное заполнение палитры отдельными компонентами,\n"\
											"Кнопка [F1] - вызов данного сообщения.";
							
							MessageBox( hWnd, help, "Интерфейс управления", MB_ICONINFORMATION );
							break;
						}
				} // switch (wParam)
				break;
			}
			
		case WM_CREATE:
			{
				vr_x = VR_X, vr_y = VR_Y,
				vr_w = VR_W, vr_h = VR_H;
				
				break;
			}

		case WM_DESTROY:
			{
				backbuff_done();
				PostQuitMessage(0);
				break;
			}
		default:
			{
				return DefWindowProc(hWnd, message, wParam, lParam);
			}
	}
	return 0;
}

void paintMandelbrot_Slow(HDC hDC, int w, int h, double dx, double dy, double left, double top) {
	DWORD i;
	double x2, y2, px, py, zx, zy;
	py = top;
	
	dx = dx / w;
	dy = dy / h;
	
	COLORREF *pix = (COLORREF*)lpColor;
	
	for (int y = 0; y < h; ++y) {
	
		px = left;
		
		for (int x = 0; x < w; ++x) {
		
			zx = px, zy = py; // экономлю итерацию на 0*0
			
			for (i = 0; i < CUR_MAX_ITER; ++i) {
			
				x2 = zx * zx;
				y2 = zy * zy;
				
				if (x2 + y2 > 4.0) { break; }
				
				zy = zx * zy * 2 + py;
				zx = x2 - y2 + px;
			}
			
			*pix++ = (i == CUR_MAX_ITER) ? 0 : palette[ i ];
			
			px += dx;
			
		} // for x
		
		py += dy;
		
	} // for y
}

// double to string
void d2a(double x, char* a) {
	char *p = a;
	
	// обработка знака
	if (x < 0.0) *p++ = '-', x = -x;

	// вычисление порядка
	DWORD x_l = *(DWORD*)((DWORD*)&x + 1);
	DWORD q = (x_l >> 20) & (~(1 << 11));
	q -= 1023;
	
	// формирование целой части
	DWORD	i;
	__asm {
		fld		x
		fistp	i
	}
	p += q; // не совпадает с числом десятичных знаков (!!!)
	while (i) {
		*p-- = ((i % 10) + '0');
		i /= 10;
	}
	
//	while (x >= 1.0) {
//		BYTE b = 
//	}
	
}