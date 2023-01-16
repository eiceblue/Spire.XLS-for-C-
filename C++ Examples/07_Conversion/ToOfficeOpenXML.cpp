#include "pch.h"
using namespace Spire::Xls;

int main() {
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToOfficeOpenXML.xml";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"A1")->SetText(L"Hello World");
	sheet->GetRange(L"B1")->GetStyle()->SetKnownColor(ExcelColors::Gray25Percent);
	sheet->GetRange(L"C1")->GetStyle()->SetKnownColor(ExcelColors::Gold);

	//Save to file.
	workbook->SaveAsXml(outputFile.c_str());
	workbook->Dispose();
}
