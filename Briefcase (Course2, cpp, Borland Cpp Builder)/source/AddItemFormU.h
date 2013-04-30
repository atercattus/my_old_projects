#ifndef AddItemFormU_H
#define AddItemFormU_H

#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>

class TAddItemForm : public TForm
{
__published:	// IDE-managed Components
	TLabel *Label1;
	TEdit *NameEdit;
	TLabel *Label2;
	TEdit *CostEdit;
	TLabel *Label3;
	TEdit *ProfitEdit;
	TLabel *Label4;
	TBitBtn *BitBtn1;
	TBitBtn *BitBtn2;
	void __fastcall FormShow(TObject *Sender);
	void __fastcall BitBtn2Click(TObject *Sender);
	void __fastcall FormPaint(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TAddItemForm(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TAddItemForm *AddItemForm;
//---------------------------------------------------------------------------

bool AnsiString2Int(AnsiString &str, int *val);

#endif
