#include <vcl.h>
#include "ParamFormU.h"
#include "AddItemFormU.h"

#pragma warn -lvc

#pragma resource "*.dfm"
TParamForm *ParamForm;
__fastcall TParamForm::TParamForm(TComponent* Owner) : TForm(Owner){}
//---------------------------------------------------------------------------

void LinearFill(TForm *frm)
{
	const TopColor    = 0xFDDFC4;
	const BottomColor = 0xFBB16F;

    float R = GetRValue(TopColor);         
    float G = GetGValue(TopColor);
    float B = GetBValue(TopColor);

    int H = frm->ClientHeight,
        W = frm->ClientWidth;

    float dR = (float)(GetRValue(BottomColor)-R)/H;
    float dG = (float)(GetGValue(BottomColor)-G)/H;
    float dB = (float)(GetBValue(BottomColor)-B)/H;

    for (int i=0; i<H; i++)
    {
        frm->Canvas->MoveTo(0, i);
        frm->Canvas->Pen->Color = (TColor)RGB(R, G, B);
        frm->Canvas->LineTo(W, i);

        R += dR;
        G += dG;
        B += dB;
    }
}

void __fastcall TParamForm::FormCreate(TObject *Sender)
{
	DoubleBuffered = true;
}
//---------------------------------------------------------------------------
void __fastcall TParamForm::FormPaint(TObject *Sender)
{
	LinearFill(ParamForm);
}
//---------------------------------------------------------------------------
void __fastcall TParamForm::FormShow(TObject *Sender)
{
	Edit1->SetFocus(); 	
}
//---------------------------------------------------------------------------

void __fastcall TParamForm::BitBtn1Click(TObject *Sender)
{
    int tmp;

    for (int i=1; i<=8; i++)
    {
    	TEdit *e = (TEdit*)ParamForm->FindChildControl("Edit"+IntToStr(i));
        if (!e) continue; // ???

        if (!AnsiString2Int(e->Text, &tmp))
        {
        	e->SetFocus(); 
            return;
        }

        e->Text = IntToStr(tmp);
         
    }

    ModalResult = mrOk;

}
//---------------------------------------------------------------------------

