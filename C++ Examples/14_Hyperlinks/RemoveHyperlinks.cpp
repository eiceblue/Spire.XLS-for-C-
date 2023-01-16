#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"HyperlinksSample1.xlsx";
	wstring outputFile = output_path + L"RemoveHyperlinks.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Get the collection of all hyperlinks in the worksheet
	HyperLinksCollection* links = dynamic_cast<HyperLinksCollection*>(sheet->GetHyperLinks());

	//Remove all link content
	sheet->GetRange(L"B1")->ClearAll();
	sheet->GetRange(L"B2")->ClearAll();
	sheet->GetRange(L"B3")->ClearAll();

	//Remove hyperlink and keep link text
	sheet->GetHyperLinks()->RemoveAt(0);
	sheet->GetHyperLinks()->RemoveAt(0);
	sheet->GetHyperLinks()->RemoveAt(0);

	//Save to file.
	workbook->SaveToFile(outputFile.c_str(), ExcelVersion::Version2013);
	workbook->Dispose();
}