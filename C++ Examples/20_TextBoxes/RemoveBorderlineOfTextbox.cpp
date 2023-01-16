#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring output = OUTPUTPATH;
	wstring outputFile = output + L"RemoveBorderlineOfTextbox.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();
	workbook->SetVersion(ExcelVersion::Version2013);

	//Create a new worksheet named "Remove Borderline" and add a chart to the worksheet.
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);
	sheet->SetName(L"Remove Borderline");
	Spire::Xls::Chart* chart = sheet->GetCharts()->Add();

	//Create textbox1 in the chart and input text information.
	XlsTextBoxShape* textbox1 = dynamic_cast<XlsTextBoxShape*>(chart->GetTextBoxes()->AddTextBox(50, 50, 100, 600));
	textbox1->SetText(L"The solution with borderline");

	//Create textbox2 in the chart, input text information and remove borderline.
	XlsTextBoxShape* textbox2 = dynamic_cast<XlsTextBoxShape*>(chart->GetTextBoxes()->AddTextBox(1000, 50, 100, 600));
	textbox2->SetText(L"The solution without borderline");
	textbox2->GetLine()->SetWeight(0);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}