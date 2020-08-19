//---------------------------------------------------------------------------

#ifndef changeityH
#define changeityH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
//---------------------------------------------------------------------------
class TForm2 : public TForm
{
__published:	// IDE-managed Components
	TRichEdit *cc;
	TButton *×ª»»;
	TLabel *Label1;
	TLabel *Label2;
	TLabel *Label3;
	TLabel *successTime;
	TButton *clear;
	TButton *help;
	TRichEdit *Column1;
	TEdit *Column2;
	void __fastcall ×ª»»Click(TObject *Sender);
	void __fastcall clearClick(TObject *Sender);
	void __fastcall ccChange(TObject *Sender);
	void __fastcall helpClick(TObject *Sender);
	String  __fastcall trim(String &mystr);
	void __fastcall Column1Change(TObject *Sender);
	void __fastcall Column2Change(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TForm2(TComponent* Owner);
	void Utf8SubStr(String string, int length);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm2 *Form2;
//---------------------------------------------------------------------------
#endif
