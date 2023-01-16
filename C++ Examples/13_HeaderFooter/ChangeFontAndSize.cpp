#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ChangeFontAndSizeForHeaderAndFooter.xlsx";
	wstring outputFile = output_path + L"ChangeFontAndSize.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the new font and size for the header and footer
	wstring text = sheet->GetPageSetup()->GetLeftHeader();
	//"Arial Unicode MS" is font name, L"18" is font size
	text = L"&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS ";
	sheet->GetPageSetup()->SetLeftHeader(text.c_str());
	sheet->GetPageSetup()->SetRightFooter(text.c_str());

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}