#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ReadImages.xlsx";
   	wstring output_path = OUTPUTPATH;
    	wstring outputFile = output_path + L"ToHtmlStream.html";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	//Set the html options
	HTMLOptions* options = new HTMLOptions();
	options->SetImageEmbedded(true);
	//Save sheet to html stream
	Stream* stream = new Stream();

	//Save to file.
	sheet->SaveToHtml(stream, options);
	workbook->Dispose();

	stream->Save(outputFile.c_str());
}
