#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertControls_result.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	sheet->GetRange(L"F1")->SetText(L"Monday");
	sheet->GetRange(L"F2")->SetText(L"Tuesday");
	sheet->GetRange(L"F3")->SetText(L"Wednesday");
	sheet->GetRange(L"F4")->SetText(L"Thursday");
	sheet->GetRange(L"F5")->SetText(L"Friday");
	sheet->GetRange(L"F6")->SetText(L"Saturday");
	sheet->GetRange(L"F7")->SetText(L"Sunday");

	//Add a textbox 
	ITextBoxShape* textbox = sheet->GetTextBoxes()->AddTextBox(9, 2, 25, 100);
	textbox->SetText(L"Hello World");
	//Add a checkbox 
	ICheckBox* cb = sheet->GetCheckBoxes()->AddCheckBox(11, 2, 15, 100);
	cb->SetCheckState(Spire::Xls::CheckState::Checked);
	cb->SetText(L"Check Box 1");
	//Add a RadioButton 
	IRadioButton* rb = sheet->GetRadioButtons()->Add(13, 2, 15, 100);
	rb->SetText(L"Option 1");

	//Add a combox
	IComboBoxShape* cbx = dynamic_cast<IComboBoxShape*>(sheet->GetComboBoxes()->AddComboBox(15, 2, 15, 100));
	cbx->SetListFillRange(sheet->GetRange(L"F1:F7"));

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}

