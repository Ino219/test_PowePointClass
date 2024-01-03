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
////�X���C�h���}�`�̎擾
//void powerpointClass::getShapes() {
//	//�擾����X���C�h�̔ԍ�(�����ł�1�Ŗڂ��擾)
//	int slideIndex = 1;
//	//�\�̃f�[�^��ݒ肵�A����
//	inputData();
//
//}
//
//void powerpointClass::inputData() {
//
//	int slideIndex = 1;
//
//	//�s���0~19�܂ł̗�����p��
//	//�w�b�_�[�̕����������āA+1
//	Random^ rnd = gcnew Random();
//	int ColumnsValue = rnd->Next(1, 4)+1;
//	Random^ rnd2 = gcnew Random();
//	int RowsValue = rnd2->Next(1, 8)+1;
//
//	//�f�o�b�O�p�̏o��
//	Debug::WriteLine(ColumnsValue);
//	Debug::WriteLine(RowsValue);
//
//	//�w�肵�����O�̃e�[�u�����擾
//	Microsoft::Office::Interop::PowerPoint::Shape^ shape=getTable(slideIndex,"special1");
//	//�擾�ł��Ȃ��ꍇ�̏���
//	if (shape == nullptr) {
//		Debug::WriteLine("noTable");
//	}
//
//	//�������e���v���[�g��葽���ꍇ�Ə��Ȃ��ꍇ�̑���
//	//�e���v���[�g���񐔂����Ȃ��ꍇ
//	if (ColumnsValue < shape->Table->Columns->Count) {
//		//�폜����񐔂��`
//		int deleteColumn = shape->Table->Columns->Count - ColumnsValue;
//		for (int i = 0; i < deleteColumn; i++) {
//			//��������폜����
//			shape->Table->Columns[shape->Table->Columns->Count - i]->Delete();
//		}
//	//�e���v���[�g���񐔂������ꍇ
//	}else if (ColumnsValue > shape->Table->Columns->Count) {
//		//�ǉ�����񐔂��`
//		int addColumn = ColumnsValue - shape->Table->Columns->Count;
//		for (int i = 0; i < addColumn; i++) {
//			//�����ɗ�̒ǉ�����
//			shape->Table->Columns->Add(-1);
//			//�󗓂̃e�L�X�g�����
//			shape->Table->Columns[shape->Table->Columns->Count]->Cells[1]->Shape->TextFrame2->TextRange->Text = "header";
//		}
//	}
//	//�e���v���[�g���s�������Ȃ��ꍇ
//	if (RowsValue < shape->Table->Rows->Count) {
//		//�폜����s�����`
//		int deleteRow = shape->Table->Rows->Count - RowsValue;
//		for (int i = 0; i < deleteRow; i++) {
//			//��������폜����
//			shape->Table->Rows[shape->Table->Rows->Count - i]->Delete();
//		}
//	}
//	//�e���v���[�g���s���������ꍇ
//	else if (RowsValue > shape->Table->Rows->Count) {
//		//�ǉ�����s�����`
//		int addRow = RowsValue-shape->Table->Rows->Count;
//		for (int i = 0; i < addRow; i++) {
//			//�����ɍs��ǉ�
//			shape->Table->Rows->Add(-1);
//			//�O�̍s�̍��ڃ^�C�g�����擾
//			String^ itemText = shape->Table->Columns[1]->Cells[shape->Table->Rows->Count-1]->Shape->TextFrame2->TextRange->Text;
//			//���K�\���Ő����𒊏o
//			String^ num = Regex::Replace(itemText, "[^0-9]", "");
//			//�ϊ���̐������`
//			int result;
//			//int�ɕϊ��\���ǂ����𔻒�
//			bool check = int::TryParse(num, result);
//			//�ϊ��\�ł���Ώ���������
//			if (check) {
//				//���ڃ^�C�g���̖����̐������X�V
//				shape->Table->Columns[1]->Cells[shape->Table->Rows->Count]->Shape->TextFrame2->TextRange->Text ="item"+(result+1);
//			}
//		}
//	}
//
//	//�f�[�^����
//	//�w�b�_�[���΂�
//	for (int i = 1; i < shape->Table->Rows->Count; i++) {
//		for (int j = 1; j < shape->Table->Columns->Count; j++) {
//			//�e�Z���ɓ��͏���
//			shape->Table->Columns[j + 1]->Cells[i + 1]->Shape->TextFrame2->TextRange->Text=j+"_"+i;
//		}
//	}
//}
//
//Microsoft::Office::Interop::PowerPoint::Shape ^ powerpointClass::getTable(int slideIndex,String ^ tableName)
//{
//	//�߂�l�̕ϐ����`
//	powerpointClass::tableshape = nullptr;
//	//���쒆�̃X���C�h����}�`�𒊏o���鏈��
//	for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in powerpointClass::presen->Slides[slideIndex]->Shapes)
//	{
//		//�w�肳�ꂽ���̂̃e�[�u�����擾
//		if (var->HasTable == MsoTriState::msoTrue&&var->Name == tableName) {
//			//������v����Ύ擾
//			tableshape = var;
//		}
//
//	}
//	//�߂�l���`
//	return tableshape;
//}
//
//void powerpointClass::savePPT(System::String ^ fileName)
//{
//	//�w�肵���t�@�C�����ŕۑ�
//	powerpointClass::presen->SaveAs(fileName, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);
//}
//
//void powerpointClass::closePPT() {
//	//���\�[�X�̊J��
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
