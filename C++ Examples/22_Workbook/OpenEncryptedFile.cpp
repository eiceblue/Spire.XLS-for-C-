#include "pch.h"
using namespace Spire::Xls;

int main() {
	std::wstring input_path = DATAPATH;
	std::wstring output_path = OUTPUTPATH;
	std::wstring inputFile = input_path + L"EncryptedFile.xlsx";
	std::wstring outputFile = output_path + L"OpenEncryptedFile.txt";
	wfstream ofs;

	//Create string builder
	wstring* builder = new wstring();

	vector<wstring> passwords = { L"password1", L"password2", L"password3", L"1234" };
	for (int i = 0; i < passwords.size(); i++)
	{
		try
		{
			//Create a workbook
			Workbook* workbook = new Workbook();

			//Open password
			workbook->SetOpenPassword(passwords[i].c_str());

			//Load the document
			workbook->LoadFromFile(inputFile.c_str());

			builder->append(L"Password = L" + passwords[i] + L" is correct." + L" The encrypted Excel file opened successfully!");
		}
		catch (const std::exception& ex)
		{
			builder->append(L"Password = L" + passwords[i] + L"  is not correct");
		}
	}

	//Save to file.
	ofs.open(outputFile, ios::out);
	ofs << builder << endl;
	ofs.close();
}