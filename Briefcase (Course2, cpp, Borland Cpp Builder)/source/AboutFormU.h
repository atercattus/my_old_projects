//---------------------------------------------------------------------------

#ifndef AboutFormUH
#define AboutFormUH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <Buttons.hpp>
#include <AppEvnts.hpp>
//---------------------------------------------------------------------------
class TAboutForm : public TForm
{
__published:	// IDE-managed Components
	TImage *Image1;
	TLabel *Label1;
	TLabel *Label2;
	TMemo *Memo;
	TLabel *Label3;
	TTimer *Timer;
	TPaintBox *PB;
	TApplicationEvents *ApplicationEvents1;
	void __fastcall Label1MouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
	void __fastcall Label1MouseUp(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
	void __fastcall Label1MouseMove(TObject *Sender, TShiftState Shift, int X,
          int Y);
	void __fastcall FormMouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
	void __fastcall FormMouseMove(TObject *Sender, TShiftState Shift, int X,
          int Y);
	void __fastcall FormMouseUp(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall Label3Click(TObject *Sender);
	void __fastcall Label3MouseEnter(TObject *Sender);
	void __fastcall Label3MouseLeave(TObject *Sender);
	void __fastcall TimerTimer(TObject *Sender);
	void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
	void __fastcall FormDestroy(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TAboutForm(TComponent* Owner);
    void PaintText(bool dY);
};
//---------------------------------------------------------------------------
extern PACKAGE TAboutForm *AboutForm;
//---------------------------------------------------------------------------
#endif
