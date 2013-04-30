#include <vcl.h>

#include "AboutFormU.h"
#include "ParamFormU.h"

Graphics::TBitmap *b;

#pragma resource "*.dfm"
TAboutForm *AboutForm;
__fastcall TAboutForm::TAboutForm(TComponent* Owner) : TForm(Owner) {}
//---------------------------------------------------------------------------

bool ClickBool;
int Xs, Ys;
int Index, Yn, TH;

TCanvas *CANVAS;

const DELTA = 8;

void __fastcall TAboutForm::Label1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
	AboutForm->OnMouseDown(AboutForm, Button, Shift, X, Y);
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::Label1MouseUp(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
	AboutForm->OnMouseUp(AboutForm, Button, Shift, X, Y);
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::Label1MouseMove(TObject *Sender,
      TShiftState Shift, int X, int Y)
{
	AboutForm->OnMouseMove(AboutForm, Shift, X, Y);	
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::FormMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
	ClickBool = true;
	Xs = X;
    Ys = Y;
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::FormMouseMove(TObject *Sender,
      TShiftState Shift, int X, int Y)
{
	if (ClickBool)
    {
		X = Left + (X-Xs);
		Y = Top + (Y-Ys);
		MoveWindow(Handle, X, Y, Width, Height, true);
    }

}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::FormMouseUp(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
	ClickBool = false;
}
//---------------------------------------------------------------------------

void __fastcall TAboutForm::FormCreate(TObject *Sender)
{
	HRGN rgn = CreateRoundRectRgn(0, 0, Width, Height, 20, 20);
    SetWindowRgn(Handle, rgn, true);

//    DoubleBuffered = true;

    char str[256-32];
    for (int i=32; i<256; i++)
    	str[i-32] = i;
//    str[256-32] = 0; // -?

    b = new Graphics::TBitmap;
    b->Width = PB->Width;
    b->Height = PB->Height;
    b->PixelFormat = pf24bit;

	CANVAS = b->Canvas;
    CANVAS->Brush->Color = Color;
    CANVAS->Brush->Style = bsSolid;
    CANVAS->Font = PB->Font;

    TH = CANVAS->TextHeight(str)+DELTA;

    Yn = PB->Height;
    Index = 0;

    CANVAS->Brush->Style = bsClear;
    CANVAS->Pen->Color = Color;
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::Label3Click(TObject *Sender)
{
    Index = 0;
    Yn = PB->Height;
	ModalResult = mrOk;
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::Label3MouseEnter(TObject *Sender)
{
	Label3->Font->Style = TFontStyles() << fsUnderline;
}
//---------------------------------------------------------------------------

void TAboutForm::PaintText(bool dY)
{
	int Count;
    char str[100];

	if (dY) Yn--;

	if (Yn < -TH-2)
    {
		Yn = -3;
		Index++;
    }

	if (Index > Memo->Lines->Count-1)
    {
		Index = 0;
		Yn = PB->Height;
	}

	Count = (PB->Height - Yn)/TH;
    if (Count*TH != PB->Height) Count++;

    CANVAS->Brush->Color = Color;
    CANVAS->Brush->Style = bsSolid;
    CANVAS->FillRect(CANVAS->ClipRect);
    CANVAS->Brush->Style = bsClear;

    for (int i=Index; i<Index+Count; i++)
    {
    	strcpy(str, Memo->Lines->operator[](i).c_str());
        
        CANVAS->Font->Style = TFontStyles();

        if (str[0] == '\\')
        {
	        switch (str[1])
    	    {
	        	case 'b':
    	        {
                	CANVAS->Font->Style = TFontStyles() << fsBold;
                    strcpy(str, &str[2]);
                    break;
        	    }

	            case 'i':
    	        {
                	CANVAS->Font->Style = TFontStyles() << fsItalic;
                    strcpy(str, &str[2]);
                    break;
        	    }

                default:
		            strcpy(str, &str[1]);
	        } // switch str[2]
        }

        int x = (CANVAS->ClipRect.Width() - CANVAS->TextWidth(str)) / 2, 
            y = Yn+(i-Index)*TH;

		CANVAS->TextOutA(x, y, str);

	} // for i

    BitBlt(PB->Canvas->Handle, 0, 0, CANVAS->ClipRect.Width(), CANVAS->ClipRect.Height(), CANVAS->Handle, 0, 0, SRCCOPY);
}

void __fastcall TAboutForm::Label3MouseLeave(TObject *Sender)
{
	Label3->Font->Style = TFontStyles();	
}
//---------------------------------------------------------------------------
void __fastcall TAboutForm::TimerTimer(TObject *Sender)
{
	PaintText(true);
}
void __fastcall TAboutForm::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
	if (Key == 27)
    {
    	Index = 0;
        Yn = PB->Height;
    	ModalResult = mrOk;
    }
}
//---------------------------------------------------------------------------

void __fastcall TAboutForm::FormDestroy(TObject *Sender)
{
	delete b;	
}
//---------------------------------------------------------------------------



