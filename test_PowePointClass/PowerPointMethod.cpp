#include "powerpointClass.h"
#using <system.dll>
//
//using namespace Microsoft::Office::Core;
//using namespace Microsoft::Office::Interop::PowerPoint;
//using namespace System::Diagnostics;
//using namespace System::Collections::Generic;
//using namespace System::Text::RegularExpressions;

//
//void powerpointClass::openPPT(System::String^ path) {
//	powerpointClass::app = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
//	powerpointClass::presens = powerpointClass::app->Presentations;
//	powerpointClass::presen = powerpointClass::presens->Open(
//		path,
//		MsoTriState::msoFalse,
//		MsoTriState::msoFalse,
//		MsoTriState::msoFalse
//	);
//}
//
////スライド内図形の取得
//void powerpointClass::getShapes() {
//	//取得するスライドの番号(既存では1頁目を取得)
//	int slideIndex = 1;
//	//表のデータを設定し、入力
//	inputData();
//
//}
//
//void powerpointClass::inputData() {
//
//	int slideIndex = 1;
//
//	//行列の0~19までの乱数を用意
//	//ヘッダーの分を加味して、+1
//	Random^ rnd = gcnew Random();
//	int ColumnsValue = rnd->Next(1, 4)+1;
//	Random^ rnd2 = gcnew Random();
//	int RowsValue = rnd2->Next(1, 8)+1;
//
//	//デバッグ用の出力
//	Debug::WriteLine(ColumnsValue);
//	Debug::WriteLine(RowsValue);
//
//	//指定した名前のテーブルを取得
//	Microsoft::Office::Interop::PowerPoint::Shape^ shape=getTable(slideIndex,"special1");
//	//取得できない場合の処理
//	if (shape == nullptr) {
//		Debug::WriteLine("noTable");
//	}
//
//	//乱数がテンプレートより多い場合と少ない場合の操作
//	//テンプレートより列数が少ない場合
//	if (ColumnsValue < shape->Table->Columns->Count) {
//		//削除する列数を定義
//		int deleteColumn = shape->Table->Columns->Count - ColumnsValue;
//		for (int i = 0; i < deleteColumn; i++) {
//			//末尾から削除処理
//			shape->Table->Columns[shape->Table->Columns->Count - i]->Delete();
//		}
//	//テンプレートより列数が多い場合
//	}else if (ColumnsValue > shape->Table->Columns->Count) {
//		//追加する列数を定義
//		int addColumn = ColumnsValue - shape->Table->Columns->Count;
//		for (int i = 0; i < addColumn; i++) {
//			//末尾に列の追加処理
//			shape->Table->Columns->Add(-1);
//			//空欄のテキストを入力
//			shape->Table->Columns[shape->Table->Columns->Count]->Cells[1]->Shape->TextFrame2->TextRange->Text = "header";
//		}
//	}
//	//テンプレートより行数が少ない場合
//	if (RowsValue < shape->Table->Rows->Count) {
//		//削除する行数を定義
//		int deleteRow = shape->Table->Rows->Count - RowsValue;
//		for (int i = 0; i < deleteRow; i++) {
//			//末尾から削除処理
//			shape->Table->Rows[shape->Table->Rows->Count - i]->Delete();
//		}
//	}
//	//テンプレートより行数が多い場合
//	else if (RowsValue > shape->Table->Rows->Count) {
//		//追加する行数を定義
//		int addRow = RowsValue-shape->Table->Rows->Count;
//		for (int i = 0; i < addRow; i++) {
//			//末尾に行を追加
//			shape->Table->Rows->Add(-1);
//			//前の行の項目タイトルを取得
//			String^ itemText = shape->Table->Columns[1]->Cells[shape->Table->Rows->Count-1]->Shape->TextFrame2->TextRange->Text;
//			//正規表現で数字を抽出
//			String^ num = Regex::Replace(itemText, "[^0-9]", "");
//			//変換後の数字を定義
//			int result;
//			//intに変換可能かどうかを判定
//			bool check = int::TryParse(num, result);
//			//変換可能であれば処理をする
//			if (check) {
//				//項目タイトルの末尾の数字を更新
//				shape->Table->Columns[1]->Cells[shape->Table->Rows->Count]->Shape->TextFrame2->TextRange->Text ="item"+(result+1);
//			}
//		}
//	}
//
//	//データ入力
//	//ヘッダーを飛ばす
//	for (int i = 1; i < shape->Table->Rows->Count; i++) {
//		for (int j = 1; j < shape->Table->Columns->Count; j++) {
//			//各セルに入力処理
//			shape->Table->Columns[j + 1]->Cells[i + 1]->Shape->TextFrame2->TextRange->Text=j+"_"+i;
//		}
//	}
//}
//
//Microsoft::Office::Interop::PowerPoint::Shape ^ powerpointClass::getTable(int slideIndex,String ^ tableName)
//{
//	//戻り値の変数を定義
//	powerpointClass::tableshape = nullptr;
//	//操作中のスライドから図形を抽出する処理
//	for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in powerpointClass::presen->Slides[slideIndex]->Shapes)
//	{
//		//指定された名称のテーブルを取得
//		if (var->HasTable == MsoTriState::msoTrue&&var->Name == tableName) {
//			//条件一致すれば取得
//			tableshape = var;
//		}
//
//	}
//	//戻り値を定義
//	return tableshape;
//}
//
//void powerpointClass::savePPT(System::String ^ fileName)
//{
//	//指定したファイル名で保存
//	powerpointClass::presen->SaveAs(fileName, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);
//}
//
//void powerpointClass::closePPT() {
//	//リソースの開放
//	System::Runtime::InteropServices::Marshal::ReleaseComObject(powerpointClass::tableshape);
//	
//	powerpointClass::presen->Close();
//	System::Runtime::InteropServices::Marshal::ReleaseComObject(powerpointClass::presen);
//	System::Runtime::InteropServices::Marshal::ReleaseComObject(powerpointClass::presens);
//
//	powerpointClass::app->Quit();
//	System::Runtime::InteropServices::Marshal::ReleaseComObject(powerpointClass::app);
//}
//
//
