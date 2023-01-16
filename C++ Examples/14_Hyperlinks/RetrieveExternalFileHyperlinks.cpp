#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"RetrieveExternalFileHyperlinks.xlsx";
	wstring outputFile = output_path + L"RetrieveExternalFileHyperlinks.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	wstring* content = new wstring();

	//Retrieve external file hyperlinks.
	for (int i = 0; i < sheet->GetHyperLinks()->GetCount(); i++)
	{
		HyperLink* item = dynamic_cast<HyperLink*>(sheet->GetHyperLinks()->Get(i));
		wstring address = item->GetAddress();
		wstring sheetName = dynamic_cast<XlsRange*>(item->GetRange())->GetWorksheetName();
		IXLSRange* range = item->GetRange();
		content->append(L"Cell[" + to_wstring(range->GetRow()) + L"," + to_wstring(range->GetColumn()) + L"]" "in sheet \"" + sheetName + L"\" contains File URL: " + address+L"\r\n");
	}


	//Save to file.
	ofs.open(outputFile, ios::out);
	wstring result(content->begin(),content->end());
	ofs << result<< endl;
	ofs.close();
	workbook->Dispose();
}