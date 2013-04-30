#include <vcl.h>
#include <fstream.h>

#pragma warn -lvc

#include "MainFormU.h"
#include "SelectAlgoFormU.h"
#include "AddItemFormU.h"
#include "ParamFormU.h"
#include "AboutFormU.h"

#pragma resource "*.dfm"
#pragma resource "XPMan.res"
TMainForm *MainForm;
__fastcall TMainForm::TMainForm(TComponent* Owner) : TForm(Owner) {}
//---------------------------------------------------------------------------

#include "CBriefcase.cpp"

int params[8] = {STD_MAX_COST, 20,  1, 1, 1,  5, 5, 3};

CBrunchAndBound briefcase(0, params[0]);

void __fastcall TMainForm::N3Click(TObject *Sender)
{
    Close();                     
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::SpeedButton1MouseMove(TObject *Sender,
      TShiftState Shift, int X, int Y)
{
    HintLabel->Caption = ((TSpeedButton*)Sender)->Hint;
}
//---------------------------------------------------------------------------

void outBriefcase();
void outbestBriefcase(const CBriefcase &v);

template<class T> void RUN()
{
	T tmp(*(T*)&briefcase);
    tmp.calcBestVariant();
    outbestBriefcase(tmp);
}

void __fastcall TMainForm::SpeedButton4Click(TObject *Sender)
{
	if (!InvestList->Items->Count)
    {
    	MessageBox(Handle, "Портфель пуст!", "Ошибка!", MB_OK | MB_ICONERROR);
        return;
    }

	switch (SelectAlgoForm->ShowModal())
    {
    	case cBrunchAndBound:
        {
        	briefcase.calcBestVariant();
            outbestBriefcase(briefcase);
        	break;
        }

    	case cHillClimbing:
        {
        	RUN<CHillClimbing>();
        	break;
        }

    	case cSimulAnneal:
        {
			CSimulatedAnnealing tmp(*(CSimulatedAnnealing*)&briefcase);
            tmp.setMaxUnchanged(params[2]);
            tmp.setMaxSaved(params[3]);
            tmp.setK_changed(params[4]);
		    tmp.calcBestVariant();
		    outbestBriefcase(tmp);
        	break;
        }

    	case cBalancedProfit:
        {
        	RUN<CBalancedProfit>();
        	break;
        }

    	case cRandomSearch:
        {
        	RUN<CRandomSearchMod>();
        	break;
        }

    	case cIncImp:
        {
        	RUN<CIncrementalImprovement>();
        	break;
        }

    	case cStopMoment:
        {
        	RUN<CStopMoment>();
        	break;
        }

    	case cLocalOptimum:
        {
			CLocalOptimum tmp(*(CLocalOptimum*)&briefcase);
            tmp.setNum(params[7]);
            tmp.setMaxBadTests(params[5]);
            tmp.setMaxBadChanges(params[6]);
		    tmp.calcBestVariant();
		    outbestBriefcase(tmp);
        	break;
        }

        default: return;        
    }

    Label9->Caption = IntToStr(DeltaTime);
    Label10->Caption = IntToStr(ITER_COUNT);
    Label12->Caption = FloatToStrF((float)DeltaTime/ITER_COUNT, ffGeneral, 4, 3);
}
//---------------------------------------------------------------------------

void outBriefcase()
{
	int SumCost=0, SumProfit=0;

	MainForm->InvestList->Items->Clear();
    for (int i=0; i<briefcase.NumItem; i++)
    {
    	TListItem *l = MainForm->InvestList->Items->Add();
        l->Caption = briefcase.invest[i].name;
        l->SubItems->Add(IntToStr(briefcase.invest[i].cost));
        l->SubItems->Add(IntToStr(briefcase.invest[i].profit));

        SumCost += briefcase.invest[i].cost;
        SumProfit += briefcase.invest[i].profit;
    }

    MainForm->SummCostAll->Caption = IntToStr(SumCost);
    MainForm->SummProfitAll->Caption = IntToStr(SumProfit);
}

void outbestBriefcase(const CBriefcase &v)
{
	int SumCost=0, SumProfit=0;

	MainForm->BestList->Items->Clear();
    for (int i=0; i<v.NumItem; i++)
    {
    	if (!v.best[i]) continue;
        
    	TListItem *l = MainForm->BestList->Items->Add();
        l->Caption = v.invest[i].name;
        l->SubItems->Add(IntToStr(v.invest[i].cost));
        l->SubItems->Add(IntToStr(v.invest[i].profit));

        SumCost += v.invest[i].cost;
        SumProfit += v.invest[i].profit;
    }

    MainForm->SummCostSel->Caption = IntToStr(SumCost);
    MainForm->SummProfitSel->Caption = IntToStr(SumProfit);
}

void __fastcall TMainForm::SpeedButton1Click(TObject *Sender)
{
	if (AddItemForm->ShowModal() == mrOk)
    {
    	if (briefcase.exists(AddItemForm->NameEdit->Text.c_str()))
        {
        	MessageBox(0, "Инвестиция с данным названием уже имеется в портфеле!", "Ошибка!", MB_OK | MB_ICONERROR);
            return;
        }

    	int cost, profit;
        AnsiString2Int(AddItemForm->CostEdit->Text, &cost);
        AnsiString2Int(AddItemForm->ProfitEdit->Text, &profit);

        briefcase.add(AddItemForm->NameEdit->Text.c_str(), cost, profit);

        outBriefcase();
    }	
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::SpeedButton2Click(TObject *Sender)
{
	if (InvestList->ItemIndex > -1)
    {
    	if (MessageBox(Handle, "Удалить выбранный элемент?", "Подтверждение", MB_OKCANCEL | MB_ICONQUESTION) == mrOk)
        {
        	briefcase.del(InvestList->ItemIndex);
            outBriefcase();
        }
    }
}
//---------------------------------------------------------------------------

void saveBriefcase(char *path)
{
	ofstream f(path, ios::out);
    if (!f)
    {
    	MessageBox(0, "Не могу создать файл", "Ошибка", MB_OK | MB_ICONERROR);
        return;
    }

    f << briefcase.NumItem << endl;

    for (int i=0; i<briefcase.NumItem; i++)
    	f << briefcase.invest[i].cost << ' ' << briefcase.invest[i].profit << ' ' << briefcase.invest[i].name << endl;

    MessageBox(0, "Портфель был сохранен", "Информация", MB_OK | MB_ICONINFORMATION);
}

void loadBriefcase(char *path)
{
	ifstream f(path, ios::in);
    if (!f)
    {
    	MessageBox(0, "Не могу открыть файл", "Ошибка", MB_OK | MB_ICONERROR);
        return;
    }

    int ItemCount, cost, profit;
    char tmp[200];

    f >> ItemCount;

    briefcase.clear(); 

    for (int i=0; i<ItemCount; i++)
    {
    	f >> cost >> profit;
        f.getline(tmp, 200, '\n');

        briefcase.add(&tmp[1], cost, profit);             	
    }

    outBriefcase();

    MessageBox(0, "Портфель был загружен", "Информация", MB_OK | MB_ICONINFORMATION);
}

void __fastcall TMainForm::N10Click(TObject *Sender)
{
	if (SD->Execute())
    	saveBriefcase(SD->FileName.c_str());
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::N9Click(TObject *Sender)
{
	if (OD->Execute())
    	loadBriefcase(OD->FileName.c_str());	
}
//---------------------------------------------------------------------------


void __fastcall TMainForm::SpeedButton3Click(TObject *Sender)
{
	if (MessageBox(Handle, "Очистить список инвестиций?", "Подтверждение", MB_OKCANCEL | MB_ICONQUESTION) == mrOk)
    {
    	briefcase.clear();
        outBriefcase();
    }
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::SpeedButton5Click(TObject *Sender)
{
	for (int i=0; i<8; i++)
    {
    	TEdit *e = (TEdit*)ParamForm->FindChildControl("Edit"+IntToStr(i+1));
        if (!e) continue; // ???

        e->Text = IntToStr(params[i]);
    }

	if (ParamForm->ShowModal() == mrOk)
    {
        for (int i=0; i<8; i++)
        {
	    	TEdit *e = (TEdit*)ParamForm->FindChildControl("Edit"+IntToStr(i+1));
    	    if (!e) continue; // ???

            AnsiString2Int(e->Text, &params[i]);

        }

        briefcase.setMaxCost(params[0]);

        if (params[1]<1)
        {
        	MessageBox(Handle, "Неправильное значение количества итераций!", "Ошибка", MB_OK | MB_ICONERROR);
            params[1] = 1;
        }

        if (params[1]>20)
        {
        	if (MessageBox( Handle, "Введено значительное число итераций. Уменьшить до рекомендуемого значения?",
                            "Предупреждение", MB_OKCANCEL | MB_ICONQUESTION) == IDOK)
            	params[1] = 20;
        }
        
        ITER_COUNT = params[1];

    }
}
//---------------------------------------------------------------------------


void __fastcall TMainForm::N5Click(TObject *Sender)
{
   	AboutForm->Timer->Enabled = true;
	AboutForm->ShowModal();
	AboutForm->Timer->Enabled = false;
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::FormShow(TObject *Sender)
{
    for (int i=1; i<=8; i++)
    {
    	TEdit *e = (TEdit*)ParamForm->FindChildControl("Edit"+IntToStr(i));
        if (!e) continue; // ???
        e->Text = IntToStr(params[i-1]);
    }
}
//---------------------------------------------------------------------------

void __fastcall TMainForm::N8Click(TObject *Sender)
{
	briefcase.clear();
    outBriefcase();
    outbestBriefcase(briefcase);

    Label9->Caption = "0";
    Label10->Caption = "0";
    Label12->Caption = "0";
}
//---------------------------------------------------------------------------

