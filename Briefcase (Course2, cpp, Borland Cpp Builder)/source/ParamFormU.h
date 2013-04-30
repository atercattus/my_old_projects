//---------------------------------------------------------------------------

#ifndef ParamFormUH
#define ParamFormUH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <Buttons.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TParamForm : public TForm
{
__published:	// IDE-managed Components
	TBitBtn *BitBtn1;
	TBitBtn *BitBtn2;
	TBevel *Bevel1;
	TLabel *Label9;
	TLabel *Label2;
	TEdit *Edit1;
	TEdit *Edit2;
	TBevel *Bevel2;
	TLabel *Label12;
	TLabel *Label3;
	TLabel *Label4;
	TLabel *Label5;
	TEdit *Edit3;
	TEdit *Edit4;
	TEdit *Edit5;
	TBevel *Bevel3;
	TLabel *Label16;
	TLabel *Label6;
	TLabel *Label7;
	TLabel *Label8;
	TEdit *Edit6;
	TEdit *Edit7;
	TEdit *Edit8;
	TLabel *Label1;
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall FormPaint(TObject *Sender);
	void __fastcall FormShow(TObject *Sender);
	void __fastcall BitBtn1Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TParamForm(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TParamForm *ParamForm;
//---------------------------------------------------------------------------

void LinearFill(TForm *frm);

#endif
