#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;

	wstring inputFile = input_path + L"AddTextBox.xlsx";
	wstring outputFile = output_path + L"AddTextBox.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the first chart
	Chart* chart = sheet->GetCharts()->Get(0);

	//Add a Textbox
	ITextBoxLinkShape* textbox = chart->GetShapes()->AddTextBox();
	textbox->SetWidth(1200);
	textbox->SetHeight(320);
	textbox->SetLeft(1000);
	textbox->SetTop(480);
	textbox->SetText(L"This is a textbox");

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();

	
}
