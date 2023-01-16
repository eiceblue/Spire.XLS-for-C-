#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToHtml.xlsx";
    	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToHtml.html";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	HTMLOptions* options = new HTMLOptions();
	options->SetImageEmbedded(true);

	//Save to file.
	sheet->SaveToHtml(outputFile.c_str());
	workbook->Dispose();
}
