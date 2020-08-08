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
void __fastcall TForm2::转换Click(TObject *Sender)
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


// 显示Word界面
vWordApp.OlePropertySet("Visible", true);
// 新建一个文档
vWordApp.OlePropertyGet("Documents").OleFunction("Add");
//
Variant vSelect = vWordApp.OlePropertyGet("Selection");

// 插入一段文字
String str1 = "要插入的第一段文字\r\n换一个行先";
vSelect.OleProcedure("TypeText", WideString(str1));
vSelect.OleProcedure("TypeParagraph");

// 插入一个表格
int nRowCount = text->Count + 1; // 行
int nColCount = 2; // 列
vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Tables")
.OleProcedure("Add", vSelect.OlePropertyGet("Range"),
nRowCount, nColCount,1,0);
 // DefaultTableBehavior:=wdWord9TableBehavior
 // AutoFitBehavior:=wdAutoFitFixed
word_table = word_activedoc.OlePropertyGet("Tables").OleFunction("Item", 1);
   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)1);
   my_cell.OlePropertySet("Range", "序号");

 for (int j = 1; j < text->Count + 1; j++) {
	my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)1);
	my_cell.OlePropertySet("Range", j);
	my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)2);
	my_cell.OlePropertySet("Range", text->Strings[j-1].c_str());
 }

// 将光标移到文档结尾
vWordApp.OlePropertyGet("Selection").OleProcedure("EndKey", 6); //wdStory

// 再插入一段文字
String str2 = "要插入的第二段文字";
vSelect.OleProcedure("TypeText", WideString(str2));


//	//创建Word对象：
//   WordApp=CreateOleObject("Word.Application");
//   WordDocs=WordApp.OlePropertyGet("Documents");
//   //打开Word文档：
//   WordDoc=WordDocs.OleFunction("Open","c:\\test.doc");
//   WordApp.OlePropertySet("Visible",true);
//   WordDoc = WordApp.OlePropertyGet("Documents");
//   word_activedoc = WordApp.OlePropertyGet("ActiveDocument");
//   //修改表格: Cell 行,列
//   word_table = word_activedoc.OlePropertyGet("Tables").OleFunction("Item", 1);
//   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)1);
//   my_cell.OlePropertySet("Range", "11)");
//   my_cell = word_table.OleFunction("Cell", (Variant)2, (Variant)1);
//   my_cell.OlePropertySet("Range", "21)");
//   my_cell = word_table.OleFunction("Cell", (Variant)2, (Variant)2);
//   my_cell.OlePropertySet("Range", "22)");
//   //保存修改内容:
//   word_activedoc.OleProcedure("Save");
//   //打印文档:
//   WordApp.OlePropertyGet("ActiveDocument").OleFunction("PrintOut");
//   WordApp.OleProcedure("Quit");

}
//---------------------------------------------------------------------------
