#include <iostream>

#include <Xlsx/Workbook.h>

using namespace SimpleXlsx;

int main()
{
    CWorkbook Book( "Incognito" );
    CWorksheet & Sheet = Book.AddSheet( "Sheet 1" );
    CWorksheet & SecondSheet = Book.AddSheet( "Sheet 2" );

    Book.AddDefinedName( "HalfRad", 0.5, "Half radian" ).AddDefinedName( "TestSin", "sin(HalfRad)" );
    Book.AddDefinedName( "SingleCell", Sheet, CellCoord( 1, 0 ) );
    Book.AddDefinedName( "RangeCells", Sheet, CellCoord( 1, 0 ), CellCoord( 2, 0 ) );
    Book.AddDefinedName( "TestScope", Sheet, CellCoord( 1, 0 ), "", & SecondSheet );

    Sheet.AddSimpleRow( "=HalfRad" ).AddSimpleRow( "=TestSin" );
    Sheet.AddSimpleRow( "=SingleCell" ).AddSimpleRow( "=sum(RangeCells)" );

    SecondSheet.AddSimpleRow( "=TestScope" );

    if( Book.Save( "DefinedNames.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;
    return 0;
}
