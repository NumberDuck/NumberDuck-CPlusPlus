#include "../NumberDuck.hpp"
#include <stdio.h>

using namespace NumberDuck;

int main(int argc, char **argv)
{
	printf("Formula Example\n");
	printf("Create a spreadsheet with formulas!\n\n");
	
	Workbook* pWorkbook = new Workbook(Workbook::License::AGPL);
	Worksheet* pWorksheet = pWorkbook->GetWorksheetByIndex(0);
		
	for (int i = 0; i < 5; i++)
	{
		Cell* pCell = pWorksheet->GetCell(0, i);
		pCell->SetFloat(i * 2.34f);
	}

	pWorksheet->GetCell(1,0)->SetString("=SUM(A1:A5)");
	pWorksheet->GetCell(1,1)->SetString("=AVERAGE(A1:A5)");

	pWorksheet->GetCell(2,0)->SetFormula("=SUM(A1:A5)");
	pWorksheet->GetCell(2,1)->SetFormula("=AVERAGE(A1:A5)");
		
	pWorkbook->Save("FormulaExample.xls", Workbook::FileType::XLS);

	delete pWorkbook;


	Workbook* pWorkbookIn = new Workbook(Workbook::License::AGPL);
	if (pWorkbookIn->Load("FormulaExample.xls"))
	{
		Worksheet* pWorksheetIn = pWorkbookIn->GetWorksheetByIndex(0);
		Cell* pCellIn = pWorksheetIn->GetCell(2,1);
		printf("Formula: %s\n", pCellIn->GetFormula());
	}
	delete pWorkbookIn;

	return 0;
}

