//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include "comobj.hpp"
#include "changeity.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm2 *Form2;
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
	   strcpy(str,AnsiString(stext.c_str()).c_str());
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
			  Result = StringReplace(Result,".","",rf);
			  Result = Result.SubString(3,Result.Length());

//			  String pattern = " ";
//			  regex re(pattern);
//			  String fmt = "20%";
//			  String ret = regex_replace(str,re,fmt);

			  text->Append(Result);
		   }
		   p = strtok(NULL,split);
       }


Variant WordApp,WordDocs,WordDoc,vWordApp;
Variant word_activedoc;
Variant word_select;
Variant word_table;
Variant my_cell;


// ��ʾWord����
vWordApp.OlePropertySet("Visible", true);
// �½�һ���ĵ�
vWordApp.OlePropertyGet("Documents").OleFunction("Add");
//
Variant vSelect = vWordApp.OlePropertyGet("Selection");

// ����һ������
String str1 = "Ҫ����ĵ�һ������\r\n��һ������";
vSelect.OleProcedure("TypeText", WideString(str1));
vSelect.OleProcedure("TypeParagraph");

// ����һ�����
int nRowCount = text->Count + 1; // ��
int nColCount = 2; // ��
vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Tables")
.OleProcedure("Add", vSelect.OlePropertyGet("Range"),
nRowCount, nColCount,1,0);
 // DefaultTableBehavior:=wdWord9TableBehavior
 // AutoFitBehavior:=wdAutoFitFixed
word_table = word_activedoc.OlePropertyGet("Tables").OleFunction("Item", 1);
   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)1);
   my_cell.OlePropertySet("Range", "���");

 for (int j = 1; j < text->Count + 1; j++) {
	my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)1);
	my_cell.OlePropertySet("Range", j);
	my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)2);
	my_cell.OlePropertySet("Range", text->Strings[j-1].c_str());
 }

// ������Ƶ��ĵ���β
vWordApp.OlePropertyGet("Selection").OleProcedure("EndKey", 6); //wdStory

// �ٲ���һ������
String str2 = "Ҫ����ĵڶ�������";
vSelect.OleProcedure("TypeText", WideString(str2));


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
