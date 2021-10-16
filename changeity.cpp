

#include <vcl.h>
#pragma hdrstop
#include "comobj.hpp"
#include <Clipbrd.hpp>
#include "changeity.h"
#include <regex>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm2 *Form2;
//using namespace boost;
using namespace std;

//---------------------------------------------------------------------------
__fastcall TForm2::TForm2(TComponent* Owner)
	: TForm(Owner)
{
Column1->Text = 37.7;
Column2->Text =  461;

}
//---------------------------------------------------------------------------
void __fastcall TForm2::转换Click(TObject *Sender)
{
		String stext = cc->Text.Trim().c_str();

		// 去掉空格
		TReplaceFlags rf1;
		rf1 << rfReplaceAll;
		stext =  StringReplace(AnsiString(stext.c_str()).c_str()," ","",rf1);

	   char str[5000] = "";
//	   strcpy(str,AnsiString(stext.c_str()).c_str());
		strcpy(str,stext.t_str());

	   char split[] = "\n";
	   char * p = strtok (str,split);

	   TStringList *text, *regexString;
	   text = new TStringList();
	   regexString = new TStringList();
	   // 补充正则表达式
	   regexString->Append("^\\w+\\."); // a. A. 1. _
	   regexString->Append("^\\w+、");  // a、A、1、_
	   regexString->Append("^\\([\u4E00-\u9FA5]{1,3}\\)");  // (一)  (二) (三)
	   regexString->Append("^\\（[\u4E00-\u9FA5]{1,3}\\）");   // 中文的括号
	   regexString->Append("^\\(\\w+\\)");   // 英文的括号
	   regexString->Append("^\\（\\w+\\）"); // (a) (b) (A) (B) (1) （2) 英文括号
	   regexString->Append("^\\w+\\）");   // 中文的括号
	   regexString->Append("^\\w+\\)"); // a) b) A) B) 1) 2) 英文括号
	   regexString->Append("^\\w+\\>");  // 1>  a>
	   regexString->Append("^[\u4E00-\u9FA5]{1,3}、");  // 一、 二、 十一、
		// 最大到20 特殊字符无法用范围的方式去写
//	   regexString->Append("^[\u2460\u2461\u2462\u2463\u2464\u2465\u2466\u2467\u2468\u2469\u2473]");  // ① ② ③
	   regexString->Append("^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]、");  // ① ② ③
	   regexString->Append("^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\\.");  // ① ② ③
	   regexString->Append("^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\\．");  // ① ② ③
	   regexString->Append("^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]");  // ① ② ③

	   regexString->Append("^[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇]"); //  特殊的 ⑴ ⑵ ⑶ ⑷

	   // 不带点的场景  需要放在最后否则会和其他1) 这种有冲突
	   regexString->Append("^\\w+");  // a A 1_

	   int i = 0;
	   while(p!= NULL) {
		   String ssss = p;
		   if (ssss.IsEmpty() == false) {
			  AnsiString sstext =  AnsiString(p);

			  TReplaceFlags rf;
			  rf << rfReplaceAll;
			  AnsiString Result = StringReplace(sstext,"\r","",rf);

              Result = Result.Trim();
			  // 去除无效的换行
			  if (Result.IsEmpty()) {
					p = strtok(NULL,split);
                  continue;
			  }

		   // 正则属性用法参考 https://blog.csdn.net/gufengaoyue/article/details/16342115

			 Variant    Axl= Variant::CreateObject("VBScript.RegExp"); //利用脚本里的正则
			  Axl.OlePropertySet("Global",false);//只匹配第一个

			  for (int kk = 0; kk < regexString->Count; kk++) {
					Axl.OlePropertySet("Pattern", regexString->Strings[kk].t_str());  //写入正则表达式
					if (Axl.OleFunction("Test",Result.c_str())) {
					   Result = Axl.OleFunction("Replace",Result.c_str(), "");
//					   ShowMessage(Result);
					   break;
					}

			  }

//			  Axl.OlePropertySet("Pattern", "\\d+\\.");  //写入正则表达式
//			  Variant    rtn=Axl.OleFunction("Execute",Result.c_str());//开始执行
//			  for (int k=0; k!= rtn.OlePropertyGet("Count"); k++) {
//				 ShowMessage((rtn.OlePropertyGet("Item",k).OlePropertyGet("Value")));//取出匹配的东东
//			  }
//
//			  Result = Axl.OleFunction("Replace",Result.c_str(), "");
			  Result = Result.Trim();  // 去掉前后的空格
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
	// 显示Word界面
	vWordApp.OlePropertySet("Visible", false);
	// 新建一个文档
	vWordApp.OlePropertyGet("Documents").OleFunction("Add");
	//
	Variant vSelect = vWordApp.OlePropertyGet("Selection");

	// 插入一段文字
	//String str1 = "要插入的第一段文字\r\n换一个行先";
	//vSelect.OleProcedure("TypeText", WideString(str1));
	//vSelect.OleProcedure("TypeParagraph");
	// RGB 颜色 bigdecimal 值
	long shenseColor = 15057564;
	long qianseColor = 16181982;

	// 插入一个表格
	int nRowCount = text->Count + 1; // 行
	int nColCount = 2; // 列
	// 添加表格的微软官方API https://docs.microsoft.com/zh-cn/office/vba/api/word.tables.add
	vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Tables")
	.OleProcedure("Add", vSelect.OlePropertyGet("Range"),nRowCount, nColCount,1,0);

	   word_table = vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Tables").OleFunction("Item", 1);

	   // 设置表格列宽为自动    1-自动 2-百分比 3-固定 磅为单位
	   Variant varColumnWidth1 = word_table.OlePropertyGet("Columns").OleFunction("Item", 1); // 第1列
	   Variant varColumnWidth2 = word_table.OlePropertyGet("Columns").OleFunction("Item", 2); // 第2列
	   varColumnWidth1.OlePropertySet("PreferredWidthType", 3);
	   varColumnWidth1.OlePropertySet("PreferredWidth", atoi(Column1->Text.t_str())); //37.7
	   varColumnWidth2.OlePropertySet("PreferredWidthType", 3);
	   varColumnWidth2.OlePropertySet("PreferredWidth", atoi(Column2->Text.t_str())); // 451.61

	   // 第一行第一列写入序号两个字 水平居中 垂直居中
	   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)1);
	   my_cell.OlePropertySet("Range", "序号");
	   my_cell.OlePropertyGet("Range").OlePropertyGet("ParagraphFormat").OlePropertySet("Alignment", 1);
	   my_cell.OlePropertySet("VerticalAlignment", 1);
	   // 表格添加背景色 2021年10月16日
	   if (CheckBackColor->Checked) {
		  my_cell.OlePropertyGet("Shading").OlePropertySet("BackgroundPatternColor", shenseColor);
	   }

	   // 第一行 第二列 水平居中  垂直居中
	   my_cell = word_table.OleFunction("Cell", (Variant)1, (Variant)2);
	   my_cell.OlePropertyGet("Range").OlePropertyGet("ParagraphFormat").OlePropertySet("Alignment", 1);
	   my_cell.OlePropertySet("VerticalAlignment", 1);
	   // 表格添加背景色 2021年10月16日
	   if (CheckBackColor->Checked){
		   my_cell.OlePropertyGet("Shading").OlePropertySet("BackgroundPatternColor", shenseColor);
	   }


	 for (int j = 1; j < text->Count + 1; j++) {
		// 第一列数据
		my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)1);
		my_cell.OlePropertySet("Range", j);
		// 水平居中
		my_cell.OlePropertyGet("Range").OlePropertyGet("ParagraphFormat").OlePropertySet("Alignment", 1);
		// 垂直居中
		my_cell.OlePropertySet("VerticalAlignment", 1);
		// 表格添加背景色 2021年10月16日
		if (CheckBackColor->Checked) {
			 my_cell.OlePropertyGet("Shading").OlePropertySet("BackgroundPatternColor", qianseColor);
		}

        // 第二列数据
		my_cell = word_table.OleFunction("Cell", (Variant)(j + 1), (Variant)2);
		my_cell.OlePropertySet("Range", text->Strings[j-1].t_str());

		// 给表格添加颜色 2021年10月16日
//	   long colorRGB =R +  G * 256 + B * 65536  ;   // RGB 转 bigdecimal

	 }

	// 将光标移到文档结尾
//	vWordApp.OlePropertyGet("Selection").OleProcedure("EndKey", 6); //wdStory
	// 设定选取范围 并复制
	vSelect.OleProcedure("SetRange", vSelect.OlePropertyGet("Start"),
				vWordApp.OlePropertyGet("ActiveDocument").OlePropertyGet("Content").OlePropertyGet("End"));
	vSelect.OleProcedure("Copy"); 
	// 再插入一段文字
	//String str2 = "要插入的第二段文字";
	//vSelect.OleProcedure("TypeText", WideString(str2));
//   //保存修改内容:
//   word_activedoc.OleProcedure("Save");
//   //打印文档:
//   WordApp.OlePropertyGet("ActiveDocument").OleFunction("PrintOut");
	// 不保存退出
   vWordApp.OleProcedure("Quit", 0);
   successTime->Caption = nRowCount - 1;

}


//---------------------------------------------------------------------------


void __fastcall TForm2::clearClick(TObject *Sender)
{
	cc->Text = "";   // 清空
	successTime->Caption = 0 ;// 转换次数改成0
	cc->SetFocus();   // 获取光标闪烁

}
//---------------------------------------------------------------------------

void __fastcall TForm2::ccChange(TObject *Sender)
{
	 successTime->Caption = 0 ;// 转换次数改成0
}
//---------------------------------------------------------------------------

void __fastcall TForm2::helpClick(TObject *Sender)
{
	ShowMessage("点击转换按钮后，表格已经复制到剪切板，直接粘贴即可！\r\n\r\n版本 v1.2 版权所有-吕天伊 2021年10月16日 ");
}

//---------------------------------------------------------------------------

void __fastcall TForm2::Column1Change(TObject *Sender)
{
       successTime->Caption = 0 ;// 转换次数改成0
}
//---------------------------------------------------------------------------

void __fastcall TForm2::Column2Change(TObject *Sender)
{
	successTime->Caption = 0 ;// 转换次数改成0
}
//---------------------------------------------------------------------------

