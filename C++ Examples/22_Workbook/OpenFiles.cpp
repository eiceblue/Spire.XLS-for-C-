#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"ExcelSample_N1.xlsx";
	std::wstring inputFile_97 = input_path + L"ExcelSample97_N.xls";
	std::wstring inputFile_xml = input_path + L"OfficeOpenXML_N.xml";
	std::wstring inputFile_csv = input_path + L"CSVSample_N.csv";
	std::wstring outputFile = output_path + L"OpenFiles.txt";
	wfstream ofs;

	//Create string builder
	wstring* builder = new wstring();

	//1. Load file by file path
	//Create a workbook
	Workbook* workbook1 = new Workbook();
	//Load the document from disk
	workbook1->LoadFromFile(inputFile.c_str());
	builder->append(L"Workbook opened using file path successfully!");

	//2. Load file by file stream
	ifstream inputf(inputFile.c_str(), ios::in | ios::binary);
	Stream* stream = new Stream(inputf);
	//Create a workbook
	Workbook* workbook2 = new Workbook();
	//Load the document from disk
	workbook2->LoadFromStream(stream);
	builder->append(L"Workbook opened using file stream successfully!");
	delete stream;

	//3. Open Microsoft Excel 97 - 2003 file
	Workbook* wbExcel97 = new Workbook();
	wbExcel97->LoadFromFile(inputFile_97.c_str(), ExcelVersion::Version97to2003);
	builder->append(L"Microsoft Excel 97 - 2003 workbook opened successfully!");

	//4. Open xml file
	Workbook* wbXML = new Workbook();
	wbXML->LoadFromXml(inputFile_xml.c_str());
	builder->append(L"XML file opened successfully!");

	//5. Open csv file
	Workbook* wbCSV = new Workbook();
	wbCSV->LoadFromFile(inputFile_csv.c_str(), L",", 1, 1);
	builder->append(L"CSV file opened successfully!");

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << *builder << endl;
	ofs.close();
}