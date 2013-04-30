#include <vcl.h>
#include "SelectAlgoFormU.h"
#pragma resource "*.dfm"
TSelectAlgoForm *SelectAlgoForm;
__fastcall TSelectAlgoForm::TSelectAlgoForm(TComponent* Owner) : TForm(Owner){}
//---------------------------------------------------------------------------

#include "ParamFormU.h"

void __fastcall TSelectAlgoForm::SpeedButton3Click(TObject *Sender)
{
	ModalResult = cBrunchAndBound;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton1Click(TObject *Sender)
{
	ModalResult = cHillClimbing;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton2Click(TObject *Sender)
{
	ModalResult = cSimulAnneal;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton4Click(TObject *Sender)
{
	ModalResult = cBalancedProfit;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton5Click(TObject *Sender)
{
	ModalResult = cRandomSearch;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton6Click(TObject *Sender)
{
	ModalResult = cIncImp;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton7Click(TObject *Sender)
{
	ModalResult = cStopMoment;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::SpeedButton8Click(TObject *Sender)
{
	ModalResult = cLocalOptimum;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::FormPaint(TObject *Sender)
{
	LinearFill(SelectAlgoForm);
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
	if (Key == 27) ModalResult = mrCancel;	
}
//---------------------------------------------------------------------------

void __fastcall TSelectAlgoForm::FormCreate(TObject *Sender)
{
/*
	if (Screen->Height*0.8 < Height)
    	Height = Screen->Height*0.8;
*/
}
//---------------------------------------------------------------------------



