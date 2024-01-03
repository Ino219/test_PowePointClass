#include "powerpointClass.h"
//#include "\Users\chach\source\repos\PowerPointDll_Test\PowerPointDll_Test\PowerpointLib.h"

//#include "\Users\chach\source\repos\test_PowePointClass\TestPowerPointLibrary1\TestPowerPointLibrary1.h"
#using <system.dll>

//using namespace Microsoft::Office::Core;
//using namespace Microsoft::Office::Interop::PowerPoint;

using namespace System::Diagnostics;

int main() {
	/*int answer=init(1,2);
	Debug::WriteLine(answer);*/
	/*ClassLibrary1::Class1::openPPT("C:\\Users\\chach\\Desktop\\test.pptx");
	ClassLibrary1::Class1::savePPT("C:\\Users\\chach\\Desktop\\test_2.pptx");
	ClassLibrary1::Class1::closePPT();*/

	//TestPowerPointLibrary1::Special::open("C:\\Users\\chach\\Desktop\\test.pptx");
	//TestPowerPointLibrary1::Special::save("C:\\Users\\chach\\Desktop\\test_3.pptx");
	//TestPowerPointLibrary1::Special::close();

	TestPowerPointLibrary1::Special::open("C:\\Users\\chach\\Desktop\\test.pptx");
	TestPowerPointLibrary1::Special::save("C:\\Users\\chach\\Desktop\\test_3.pptx");
	TestPowerPointLibrary1::Special::close();
	//powerpointClass::openPPT("C:\\Users\\chach\\Desktop\\test.pptx");

	//powerpointClass::getShapes();

	//powerpointClass::savePPT("C:\\Users\\chach\\Desktop\\test_.pptx");

	//powerpointClass::closePPT();
	
	return 0;
}

