#include "testprojectClass.h"


int main() {


	powerpointClass::openPPT("C:\\Users\\chach\\Desktop\\test.pptx");

	powerpointClass::getShapes();

	powerpointClass::savePPT("C:\\Users\\chach\\Desktop\\test_.pptx");

	powerpointClass::closePPT();

	return 0;
}