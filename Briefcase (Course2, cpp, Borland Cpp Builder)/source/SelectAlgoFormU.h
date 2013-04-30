#ifndef SelectAlgoFormUH
#define SelectAlgoFormUH

#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TSelectAlgoForm : public TForm
{
__published:	// IDE-managed Components
	TSpeedButton *SpeedButton1;
	TSpeedButton *SpeedButton2;
	TLabel *Label1;
	TLabel *Label2;
	TLabel *Label3;
	TLabel *Label4;
	TSpeedButton *SpeedButton3;
	TLabel *Label5;
	TLabel *Label6;
	TBevel *Bevel1;
	TSpeedButton *SpeedButton4;
	TLabel *Label7;
	TLabel *Label8;
	TSpeedButton *SpeedButton5;
	TLabel *Label9;
	TLabel *Label10;
	TSpeedButton *SpeedButton6;
	TLabel *Label11;
	TLabel *Label12;
	TSpeedButton *SpeedButton7;
	TLabel *Label13;
	TLabel *Label14;
	TSpeedButton *SpeedButton8;
	TLabel *Label15;
	TLabel *Label16;
	void __fastcall SpeedButton3Click(TObject *Sender);
	void __fastcall SpeedButton1Click(TObject *Sender);
	void __fastcall SpeedButton2Click(TObject *Sender);
	void __fastcall SpeedButton4Click(TObject *Sender);
	void __fastcall SpeedButton5Click(TObject *Sender);
	void __fastcall SpeedButton6Click(TObject *Sender);
	void __fastcall SpeedButton7Click(TObject *Sender);
	void __fastcall SpeedButton8Click(TObject *Sender);
	void __fastcall FormPaint(TObject *Sender);
	void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
	void __fastcall FormCreate(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TSelectAlgoForm(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TSelectAlgoForm *SelectAlgoForm;
//---------------------------------------------------------------------------

const cBrunchAndBound	= 0x1000;
const cHillClimbing		= 0x1001;
const cSimulAnneal		= 0x1002;
const cBalancedProfit	= 0x1003;
const cRandomSearch		= 0x1004;
const cIncImp			= 0x1005;
const cStopMoment		= 0x1006;
const cLocalOptimum		= 0x1007;

#endif
