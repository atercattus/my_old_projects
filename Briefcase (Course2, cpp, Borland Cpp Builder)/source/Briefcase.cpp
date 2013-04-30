//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("MainFormU.cpp", MainForm);
USEFORM("SelectAlgoFormU.cpp", SelectAlgoForm);
USEFORM("AddItemFormU.cpp", AddItemForm);
USEFORM("ParamFormU.cpp", ParamForm);
USEFORM("AboutFormU.cpp", AboutForm);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
    try
    {
         Application->Initialize();
         Application->Title = "Briecase";
		Application->CreateForm(__classid(TMainForm), &MainForm);
		Application->CreateForm(__classid(TSelectAlgoForm), &SelectAlgoForm);
		Application->CreateForm(__classid(TAddItemForm), &AddItemForm);
		Application->CreateForm(__classid(TParamForm), &ParamForm);
		Application->CreateForm(__classid(TAboutForm), &AboutForm);
		Application->Run();
    }
    catch (Exception &exception)
    {
         Application->ShowException(&exception);
    }
    catch (...)
    {
         try
         {
             throw Exception("");
         }
         catch (Exception &exception)
         {
             Application->ShowException(&exception);
         }
    }
    return 0;
}
//---------------------------------------------------------------------------
