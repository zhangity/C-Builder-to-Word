//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include "comobj.hpp"

#include "changeity.h"
#include <boost/regex.hpp>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm2 *Form2;
using namespace boost;
//---------------------------------------------------------------------------
__fastcall TForm2::TForm2(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TForm2::ת��Click(TObject *Sender)
{
		String stext = cc->Text.Trim().c_str();
	   char str[5000] = "";
//	   strcpy(str,AnsiString(stext.c_str()).c_str());
		strcpy(str,stext.t_str());
	   char split[] = "\n";
	   char * p = strtok (str,split);

	   TStringList *text;
	   text = new TStringList();
	   int i = 0;
	   while(p!= NULL) {
		   String ssss = p;
		   if (ssss.IsEmpty() == false) {
			  AnsiString sstext =  AnsiString(p);

			  TReplaceFlags rf;
			  rf << rfReplaceAll;
			  AnsiString Result = StringReplace(sstext,"\r","",rf);

//			  Result = StringReplace(Result,".","",rf);
//			  ShowMessage(Result);
//			  Result = Result.SubString(3,Result.Length());
//              ShowMessage(Result);
//			  String pattern = " ";
//			  regex re(pattern);
//			  String fmt = "20%";
//			  String ret = regex_replace(str,re,fmt);
            // ������ʽȥ������ǰ�������
//			std::string str = "192.168.1.1";
//
//			boost::regex e1, e2;
//			e1.assign("\\d+");
//			regex expression("\\d+");
////			boost::basic_regex();
////			boost::smatch what;
//			cmatch what;
//
////			std::string::const_iterator start = str.begin();
////			std::string::const_iterator end = str.end();
////			while ( boost::regex_search(start, end, what, expression) )
////			{
////				ShowMessage(what[0].str().c_str());
////				start = what[0].second;
////			}
////			return boost::regex_split(std::back_inserter(l), s);
////			boost::regex_search(start, end, what, e1);
//			regex_match(str, what, expression);
			std::string test = "1234567890";
			boost::regex testPattern( "\\d" );
			boost::match_results<std::string::const_iterator> testMatches;
			std::string::const_iterator startPos = test.begin();
			std::string::const_iterator endPos = test.end();
			while( regex_search( startPos, endPos, testMatches, testPattern ) ) {
				// Do stuff: record match value, increment start position
			}

			  text->Append(Result);
		   }
		   p = strtok(NULL,split);
       }


Variant WordApp,WordDocs,WordDoc,vWordApp;
Variant word_activedoc;
Variant word_select;
Variant word_table;
Variant my_cell;

vWordApp = CreateOleObject("Word.Application");
// ��ʾWord����
vWordApp.OlePropertySet("Visible", true);
// �½�һ���ĵ�
vWordApp.OlePropertyGet("Documents").OleFunction("Add");
//
Variant vSelect = vWordApp.OlePropertyGet("Selection");

// ����һ������
//String str1 = "Ҫ����ĵ�һ������\r\n��һ������";
//vSelect.OleProcedure("TypeText", WideString(str1));
//vSelect.OleProcedure("TypeParagraph");

// ����һ�����
int nRowCount = text->Count + 1; // ��
int nColCount = 2; // ��
vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Tables")
.OleProcedure("Add", vSelect.OlePropertyGet("Range"),
nRowCount, nColCount,1,0);
 // DefaultTableBehavior:=wdWord9TableBehavior
 // AutoFitBehavior:=wdAutoFitFixed
   word_table = vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Tables").OleFunction("Item", 1);
   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)1);
   my_cell.OlePropertySet("Range", "���");

 for (int j = 1; j < text->Count + 1; j++) {
	my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)1);
	my_cell.OlePropertySet("Range", j);
	my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)2);

	my_cell.OlePropertySet("Range", text->Strings[j-1].t_str());
 }

// ������Ƶ��ĵ���β
vWordApp.OlePropertyGet("Selection").OleProcedure("EndKey", 6); //wdStory

// �ٲ���һ������
//String str2 = "Ҫ����ĵڶ�������";
//vSelect.OleProcedure("TypeText", WideString(str2));


//	//����Word����
//   WordApp=CreateOleObject("Word.Application");
//   WordDocs=WordApp.OlePropertyGet("Documents");
//   //��Word�ĵ���
//   WordDoc=WordDocs.OleFunction("Open","c:\\test.doc");
//   WordApp.OlePropertySet("Visible",true);
//   WordDoc = WordApp.OlePropertyGet("Documents");
//   word_activedoc = WordApp.OlePropertyGet("ActiveDocument");
//   //�޸ı��: Cell ��,��
//   word_table = word_activedoc.OlePropertyGet("Tables").OleFunction("Item", 1);
//   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)1);
//   my_cell.OlePropertySet("Range", "11)");
//   my_cell = word_table.OleFunction("Cell", (Variant)2, (Variant)1);
//   my_cell.OlePropertySet("Range", "21)");
//   my_cell = word_table.OleFunction("Cell", (Variant)2, (Variant)2);
//   my_cell.OlePropertySet("Range", "22)");
//   //�����޸�����:
//   word_activedoc.OleProcedure("Save");
//   //��ӡ�ĵ�:
//   WordApp.OlePropertyGet("ActiveDocument").OleFunction("PrintOut");
//   WordApp.OleProcedure("Quit");

}
//---------------------------------------------------------------------------

