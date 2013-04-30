#ifndef MainFormU_H
#define MainFormU_H

#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Chart.hpp>
#include <ExtCtrls.hpp>
#include <Menus.hpp>
#include <Buttons.hpp>
#include <ComCtrls.hpp>
#include <Dialogs.hpp>
#include <Graphics.hpp>
//---------------------------------------------------------------------------
class TMainForm : public TForm
{
__published:	// IDE-managed Components
    TMainMenu *MainMenu1;
    TMenuItem *N1;
    TMenuItem *N2;
    TMenuItem *N3;
    TMenuItem *N4;
    TMenuItem *N5;
    TMenuItem *N6;
    TMenuItem *N7;
    TStatusBar *StatusBar1;
    TPanel *Panel1;
    TLabel *Label1;
    TListView *InvestList;
    TGroupBox *GroupBox2;
    TSpeedButton *SpeedButton1;
    TSpeedButton *SpeedButton2;
    TSpeedButton *SpeedButton3;
    TSpeedButton *SpeedButton4;
    TSpeedButton *SpeedButton5;
    TLabel *HintLabel;
    TPanel *Panel2;
    TSplitter *Splitter1;
    TLabel *Label2;
	TListView *BestList;
    TGroupBox *GroupBox1;
    TGroupBox *GroupBox3;
    TLabel *Label3;
    TLabel *Label4;
    TLabel *SummCostAll;
    TLabel *SummProfitAll;
	TMenuItem *N8;
	TMenuItem *N9;
	TMenuItem *N10;
	TSaveDialog *SD;
	TOpenDialog *OD;
        TLabel *Label5;
        TLabel *Label6;
        TLabel *SummCostSel;
        TLabel *SummProfitSel;
        TLabel *Label7;
        TLabel *Label8;
        TLabel *Label11;
        TLabel *Label9;
        TLabel *Label10;
        TLabel *Label12;
        TBevel *Bevel1;
    void __fastcall N3Click(TObject *Sender);
    void __fastcall SpeedButton1MouseMove(TObject *Sender,
          TShiftState Shift, int X, int Y);
	void __fastcall SpeedButton4Click(TObject *Sender);
	void __fastcall SpeedButton1Click(TObject *Sender);
	void __fastcall SpeedButton2Click(TObject *Sender);
	void __fastcall N10Click(TObject *Sender);
	void __fastcall N9Click(TObject *Sender);
	void __fastcall SpeedButton3Click(TObject *Sender);
	void __fastcall SpeedButton5Click(TObject *Sender);
	void __fastcall N5Click(TObject *Sender);
	void __fastcall FormShow(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
    __fastcall TMainForm(TComponent* Owner);
};

//---------------------------------------------------------------------------
extern PACKAGE TMainForm *MainForm;
//---------------------------------------------------------------------------

#endif // MainFormU_H
