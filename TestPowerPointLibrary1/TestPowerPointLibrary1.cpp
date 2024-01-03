#include "pch.h"
#using <system.dll>

#include "TestPowerPointLibrary1.h"


void TestPowerPointLibrary1::Special::open(System::String ^ path)
{
	app1 = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
	presens1 = app1->Presentations;
	presen1 = presens1->Open(
		path,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse,
		MsoTriState::msoFalse
	);
}

void TestPowerPointLibrary1::Special::save(System::String ^ fileName)
{
	//指定したファイル名で保存
	presen1->SaveAs(fileName, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);

}

void TestPowerPointLibrary1::Special::close()
{
	//リソースの開放
	//System::Runtime::InteropServices::Marshal::ReleaseComObject(tableshape);

	presen1->Close();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen1);
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presens1);

	app1->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(app1);
}

