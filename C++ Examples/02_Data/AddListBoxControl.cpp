#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddListBoxControl.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set text for cells 
	sheet->GetRange(L"A7")->SetText(L"Beijing");
	sheet->GetRange(L"A8")->SetText(L"New York");
	sheet->GetRange(L"A9")->SetText(L"ChengDu");
	sheet->GetRange(L"A10")->SetText(L"Paris");
	sheet->GetRange(L"A11")->SetText(L"Boston");
	sheet->GetRange(L"A12")->SetText(L"London");

	sheet->GetRange(L"A7")->SetText(L"City :");
	sheet->GetRange(L"C13")->GetStyle()->GetFont()->SetIsBold(true);

	//Add listbox control
	IListBox* listBox = sheet->GetListBoxes()->AddListBox(13, 4, 120, 100);
	listBox->SetSelectionType(SelectionType::Single);
	listBox->SetSelectedIndex(2);
	listBox->SetDisplay3DShading(true);
	listBox->SetListFillRange(sheet->GetRange(L"A7:A12"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

