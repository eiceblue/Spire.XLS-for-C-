#include "pch.h"
using namespace Spire::Xls;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"templateAz.xlsx";
	wstring outputFile = output_path + L"GetColorArgbData.txt";
	wfstream ofs;

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Get the first worksheet
	Worksheet* sheet = workbook->GetWorksheets()->Get(0);

	wstring* content = new wstring();

	//Get font color
	Spire::Common::Color* color1 = sheet->GetRange(L"B2")->GetStyle()->GetFont()->GetColor();

	//Read ARGB data of Color
	string a, r, g, b;
	a = color1->GetA(), r = color1->GetR(), g = color1->GetG(), b = color1->GetB();
	content->append("The font color of B2: ARGB=(" + a + "," + r + "," + g + "," + b + ")"+L"\n");

	Spire::Common::Color* color2 = sheet->GetRange(L"B3")->GetStyle()->GetFont()->GetColor();
	a = color2->GetA(), r = color2->GetR(), g = color2->GetG(), b = color2->GetB();
	content->append("The font color of B3: ARGB=(" + a + "," + r + "," + g + "," + b + ")"+L"\n");

	Spire::Common::Color* color3 = sheet->GetRange(L"B4")->GetStyle()->GetFont()->GetColor();
	a = color3->GetA(), r = color3->GetR(), g = color3->GetG(), b = color3->GetB();
	content->append("The font color of B4: ARGB=(" + a + "," + r + "," + g + "," + b + ")");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *content << endl;
	ofs.close();
	workbook->Dispose();

}
