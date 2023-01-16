#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"AllNamedRanges.xlsx";
	wstring outputFile = output_path + L"GetAllNamedRange.txt";

	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());
	wstring* content = new wstring();

	//Get all named range
	INameRanges* ranges = workbook->GetNameRanges();
	for (int i = 0; i < ranges->GetCount(); i++)
	{
		INamedRange* nameRange = ranges->Get(i);
		content->append(nameRange->GetName());
		content->append(L"\r\n");
	}
	//Save to file.
	ofs.open(outputFile, ios::out);
	wstring result(content->begin(),content->end());
	ofs << result << endl;
	ofs.close();
	workbook->Dispose();

	ifstream f(outputFile.c_str());
}