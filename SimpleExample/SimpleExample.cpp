#include "../NumberDuck.hpp"

#include <stdio.h>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("Simple Example\n");
	printf("Create a spreadsheet!\n\n");

	Workbook* pWorkbook = new Workbook(Workbook::License::AGPL);
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);

	Cell* pCell = pWorksheet->GetCellByAddress("A1");
	pCell->SetString("Totally cool spreadsheet!");

	pWorksheet->GetCell(1,1)->SetFloat(3.1417f);

	pWorkbook->Save("SimpleExample.xls", Workbook::FileType::XLS);
	delete pWorkbook;

	Workbook* pWorkbookIn = new Workbook(Workbook::License::AGPL);
	if (pWorkbookIn->Load("SimpleExample.xls"))
	{
		Worksheet* pWorksheetIn = pWorkbookIn->GetWorksheetByIndex(0);
		Cell* pCellIn = pWorksheetIn->GetCell(0,0);
		printf("Cell Contents: %s\n", pCellIn->GetString());
	}

	delete pWorkbookIn;

	return 0;
}
