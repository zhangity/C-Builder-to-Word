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
	void __fastcall ×ª»»Click(TObject *Sender);

private:	// User declarations
public:		// User declarations
	__fastcall TForm2(TComponent* Owner);
	void Utf8SubStr(String string, int length);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm2 *Form2;
//---------------------------------------------------------------------------
#endif
