#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"SaveStream.xls";
	std::wstring outputFile = output_path + L"SaveStream.xlsx";

	//Create a workbook
	Workbook* workbook = new Workbook();

	//Load the Excel document from disk
	workbook->LoadFromFile(inputFile.c_str());

	//Save an excel workbook to stream
	ofstream outputf(outputFile.c_str(), ios::out | ios::binary);
	Spire::Common::Stream* stream = new Spire::Common::Stream();
	workbook->SaveToStream(stream);

	workbook->Dispose();
}