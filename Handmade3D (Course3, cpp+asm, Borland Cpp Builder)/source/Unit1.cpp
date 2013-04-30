#include <windows.h>
#include "F3DL.h"
#include "F3DProj.h"
#include "F3DObject.h"
#include "CubeObj.h"

//#define USE_CHANGE_PRIORITY
#define WINDOWED

#ifdef USE_CHANGE_PRIORITY
	#define DEMO_APP_PRIORITY HIGH_PRIORITY_CLASS
#endif

#define		Wnd_WndStyle	WS_OVERLAPPEDWINDOW // && (!WS_BORDER) || WS_SYSMENU	
#define		FS_WndStyle		WS_POPUP

#ifdef WINDOWED
	#define WndStyle Wnd_WndStyle
#else
	#define WndStyle FS_WndStyle
#endif

#define STDCALL long __stdcall

// consts
char *AppName = "AppName";
char *AppClass = "AppClass";

const	DWORD	TimerID = 0x100;
const	DWORD	TimerInterval = 500;

// vars
	HWND		hWnd;

	DWORD		Width = 600;
	DWORD		Height = 600;

	bool		ActiveDraw=false;

	bool		Pause = false;

	DWORD		LastTick;
	DWORD		FPS;
	DWORD		FrameCount;
	char		str[100];

    double		tmp_pos, tmp_dir;

	CCube		cube, cube2;

void DefaultProcedure()
{
	f3d_ZbufClear();
	f3d_ClearByte(0);

	if (!Pause)
	{
		double delta;

		if (!FPS)	delta = 0; // простенькая независимость скорости вращения от FPS
		else		delta = 96.0/FPS;

//        cube

		cube.Rotate(-0.5*delta, 0.1*delta, 0.7*delta);
		cube2.Rotate(-0.3*delta, -0.2*delta, -0.1*delta);
	}
	cube.Draw();
	cube2.Draw();

	f3d_Flip();
}

STDCALL WindowProc(HWND hwnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
    {
    	case WM_CREATE:
		{
			HDC hDC = GetDC(hwnd);
//			ReleaseDC(hwnd, hDC);
			if (f3d_Init(hDC, Width, Height) != F3D_NO_ERROR)
			{
				MessageBox(hwnd, "[Init] error", NULL, MB_OK | MB_ICONERROR);
				PostQuitMessage(0);
				break;
			}

			if (f3d_ZbufInit() != F3D_NO_ERROR)
			{
				MessageBox(hwnd, "[ZbufInit] error", NULL, MB_OK | MB_ICONERROR);
				f3d_Done();
				PostQuitMessage(0);
				break;
			}

			const double delta = 0.4;

			// размеры и положение
			cube.Translate(-delta, 0, 3);
			cube.Scale(0.6, 0.6, 0.6);
			cube.Rotate(-45, -45, 0);

			// использовать освещение при выводе
			cube.SetColor(0x3388EE);
			cube.SetFlags(cube.GetFlags() | F3DOF_USE_LIGHT | F3DOF_USE_LIGHT_DIR);
			cube.SetDirLight(Triple(0, -1, 0));

			// размеры и положение
			cube2.Translate(delta, 0, 3);
			cube2.Scale(0.6, 0.6, 0.6);
			cube2.Rotate(45, 45, 0);

			// использовать освещение при выводе
			cube2.SetColor(0xEEEE00);
			cube2.SetFlags(cube.GetFlags() | F3DOF_USE_LIGHT | F3DOF_USE_COLOR_LIGHT);
			cube2.SetLightColor(0xFF00FF);

			LastTick = GetTickCount();
			FrameCount = 0;
			SetTimer(hwnd, TimerID, TimerInterval, NULL);

			randomize();

			SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, NULL, 0);
		}

		case WM_TIMER:
		{
			LastTick = GetTickCount() - LastTick;

			if (LastTick)
				FPS = 1000 * FrameCount / LastTick;
			else
                FPS = 0;

			LastTick = GetTickCount();
			FrameCount = 0;

			itoa(FPS, str, 10);
			SetWindowText(hwnd, str);

        	break;
		}

		case WM_SIZE:
		{
			// если восстанавливаюсь
			if ((wParam == SIZE_RESTORED) && (LOWORD(lParam) == (S32)Width) && (HIWORD(lParam) == (S32)Height)) return 1;

			// если свернут
			if (!LOWORD(lParam) || !HIWORD(lParam)) return 1;

			f3d_ZbufDone();

			f3d_Done();

			Width = LOWORD(lParam);
			Height = HIWORD(lParam);

			HDC hDC = GetDC(hwnd);
			if (f3d_Init(hDC, Width, Height) != F3D_NO_ERROR)
			{
				MessageBox(hwnd, "Re[Init] error", NULL, MB_OK | MB_ICONERROR);
				PostQuitMessage(0);
				break;
			}
			ReleaseDC(hwnd, hDC);

			if (f3d_ZbufInit() != F3D_NO_ERROR)
			{
				MessageBox(hwnd, "Re[ZbufInit] error", NULL, MB_OK | MB_ICONERROR);
				f3d_Done();
				PostQuitMessage(0);
				break;
			}

			return 0;
		}

		case WM_SETFOCUS:
        {
        	ActiveDraw = true;
            SendMessage(hwnd, WM_SYSCOMMAND, SC_RESTORE, 0);
            break;
        }

		case WM_KILLFOCUS:
        {
        	ActiveDraw = false;
            SendMessage(hwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0);
            break;
        }

        case WM_KEYDOWN:
		{
			if (wParam == VK_SPACE) Pause = !Pause;
			if (wParam == '1') cube.SetFlags(cube.GetFlags() ^ F3DOF_USE_LIGHT_DIR);
			if (wParam == '2') cube2.SetFlags(cube2.GetFlags() ^ F3DOF_USE_COLOR_LIGHT);
			if (wParam == VK_ADD) cube2.SetLumLight(cube2.GetLumLight()+0.1);
			if (wParam == VK_SUBTRACT) cube2.SetLumLight(cube2.GetLumLight()-0.1);
			if (wParam == VK_MULTIPLY) cube2.SetLightColor(random(0x1000000));


			if (wParam != VK_ESCAPE) break;
		}

        case WM_CLOSE:
		case WM_DESTROY:
		{
			f3d_Done();

            KillTimer(hwnd, TimerID);

			SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, NULL, 0);

            ExitProcess(0);
        }
    }
    return DefWindowProc(hwnd, Message, wParam, lParam);
}

WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
#ifdef USE_CHANGE_PRIORITY
	SetPriorityClass(GetCurrentProcess(), DEMO_APP_PRIORITY); // получаю побольше ресурсов
	// void FAR PASCAL RepaintScreen(); Адрес: USER.275
#endif

	WNDCLASS	wcl;

	ZeroMemory(&wcl, sizeof(wcl));
	wcl.style = CS_OWNDC;
    wcl.lpfnWndProc = WindowProc;
	wcl.hInstance = hInstance;
	wcl.hIcon = LoadIcon(0, IDI_APPLICATION);
	wcl.hCursor = LoadCursor(0, IDC_ARROW);
	wcl.hbrBackground = GetStockObject(BLACK_BRUSH); // конечно нехорошо, но при WM_SIZE выглядит лучше
	wcl.lpszClassName = AppClass;

	RegisterClass(&wcl);

	hWnd = CreateWindowEx( 0,
							 AppClass,
							 AppName,
							 WndStyle,
							 CW_USEDEFAULT,
							 CW_USEDEFAULT,
							 CW_USEDEFAULT,
							 CW_USEDEFAULT,
							 0,
							 0,
							 hInstance,
							 NULL );

	f3d_SetClientRect(hWnd, Width, Height, true);

	ShowWindow(hWnd, SW_NORMAL);
	UpdateWindow(hWnd);

	MSG		message;

	while (1)
    {
		if (PeekMessage(&message, 0, 0, 0, PM_NOREMOVE))
		{
			if (!GetMessage(&message, 0, 0, 0)) break;
			else
			{
				TranslateMessage(&message);
				DispatchMessage(&message);
            }
        } else
        {
			if (ActiveDraw)
			{
				FrameCount++;
				DefaultProcedure();
			}
			else
				WaitMessage();
        }
	}

	return 0;

}
 