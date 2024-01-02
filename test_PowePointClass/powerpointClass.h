
#pragma once


using namespace System;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;

ref class powerpointClass
{
public:
	
	static Microsoft::Office::Interop::PowerPoint::Application^ app;
	static Microsoft::Office::Interop::PowerPoint::Presentations^ presens;
	static Microsoft::Office::Interop::PowerPoint::Presentation^ presen;
	static Microsoft::Office::Interop::PowerPoint::Shape^ tableshape;

	powerpointClass() {
	//コンストラクター
		//Debug::WriteLine("test");
	}

	static void openPPT(System::String^ path);
	static void getShapes();
	static void inputData();
	static Microsoft::Office::Interop::PowerPoint::Shape^ getTable(int slideIndex, String^ tableName);
	static void savePPT(System::String^ fileName);
	static void closePPT();

};

