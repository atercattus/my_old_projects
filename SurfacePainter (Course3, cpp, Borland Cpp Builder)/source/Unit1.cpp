#include <vcl.h>
//#######################################################
#include "Graph3D.h"
#include "Parser.h"
//#######################################################

#include "Unit1.h"
#pragma resource "*.dfm"
TMainForm *MainForm;
__fastcall TMainForm::TMainForm(TComponent* Owner) : TForm(Owner) {}
//---------------------------------------------------------------------------

//#######################################################
CGraph3D	*graph=NULL;
CParser		parser;
//#######################################################

//---------------------------------------------------------------------------
void __fastcall TMainForm::FormPaint(TObject *Sender)
{
	graph->DrawSurface();
}

//---------------------------------------------------------------------------
float plain(float x, float z)
{
	float y;
	if (parser.Eval(x, z, y))	return y;
	else						return 0.0;
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::FormCreate(TObject *Sender)
{
	DWORD tr_count = (SurfX-1)*(SurfZ-1)*2;
	Caption = Caption + "  (Треугольников: " + IntToStr(tr_count) + ")";

	if (!parser.Parse("formula.txt"))
	{
		MessageBox(Handle, "Ошибка считывания поверхности!", "Ошибка", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
		return;
	}

	MessageBox(NULL, "Поверхность считана", "Сообщение", MB_OK | MB_ICONINFORMATION);

	randomize(); // ?

	graph = new CGraph3D(Canvas);
	if (!graph) PostQuitMessage(0);

	// формирую данные поверхности
	graph->InitSurface(plain);
	// устанавливаю матрицу вида в единичную
	graph->ReInit();

	// делаю красивый вид
	graph->Translate(0, 0, -1.8);
	graph->Rotate(-0.4, -0.4, 0);

	// немного красивей
	DoubleBuffered = true;
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::FormKeyDown(TObject *Sender, WORD &Key,
	  TShiftState Shift)
{
	// выход
	if (Key == VK_ESCAPE) Close();

	// сброс матрицы вида в начальное состояние
	if (Key == VK_F2)
	{
		graph->ReInit();
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_UP)
	{
		graph->Translate(0, 0, 0.1);
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_DOWN)
	{
		graph->Translate(0, 0, -0.1);
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_NUMPAD6)
	{
		graph->Rotate(0, -0.05, 0);
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_NUMPAD4)
	{
		graph->Rotate(0, 0.05, 0);
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_LEFT)
	{
		graph->Rotate(0, 0.001, 0);
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_RIGHT)
	{
		graph->Rotate(0, +0.001, 0);
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_NUMPAD8)
	{
		graph->Rotate(-0.01, 0, 0);	
		InvalidateRect(Handle, NULL, false);
	}

	if (Key == VK_NUMPAD2)
	{
		graph->Rotate(0.01, 0, 0);	
		InvalidateRect(Handle, NULL, false);
	}
}

//---------------------------------------------------------------------------
void __fastcall TMainForm::FormResize(TObject *Sender)
{
	// вычисление размеров половин области вывода
	TRect rect = GetClientRect();
	graph->ChangeBounds(rect.Width(), rect.Height());

	InvalidateRect(Handle, NULL, false);
}

