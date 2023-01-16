#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring fn = DATAPATH;
	wstring output = OUTPUTPATH;
	wstring inputFile = fn + L"Template_Xls_5.xlsx";
	wstring outputFile = output + L"ExtractTextFromATextbox.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first textbox.
	XlsTextBoxShape* shape = dynamic_cast<XlsTextBoxShape*>(sheet->GetTextBoxes()->Get(0));

	//Extract text from the text box.
	wstring* content = new wstring();
	content->append(L"The text extracted from the TextBox is: \n");
	content->append(shape->GetText());

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();
}