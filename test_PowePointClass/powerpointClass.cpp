#include "powerpointClass.h"
#using <system.dll>

using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;

using namespace System::Diagnostics;

int main() {

	powerpointClass::openPPT("C:\\Users\\chach\\Desktop\\test.pptx");

	powerpointClass::getShapes();

	powerpointClass::savePPT("C:\\Users\\chach\\Desktop\\test_.pptx");

	powerpointClass::closePPT();

	return 0;
}

