#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddScrollBarControl.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set a value for range B10
	sheet->GetRange(L"B10")->SetValue(to_wstring(1).c_str());
	sheet->GetRange(L"B10")->GetStyle()->GetFont()->SetIsBold(true);

	//Add scroll bar control
	IScrollBarShape* scrollBar = sheet->GetScrollBarShapes()->AddScrollBar(10, 3, 150, 20);
	scrollBar->SetLinkedCell(sheet->GetRange(L"B10"));
	scrollBar->SetMin(1);
	scrollBar->SetMax(150);
	scrollBar->SetIncrementalChange(1);
	scrollBar->SetDisplay3DShading(true);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

