#include <iostream>

#include <Xlsx/Workbook.h>

using namespace SimpleXlsx;

int main( int argc, char * argv[] )
{
    ( void )argc; ( void )argv;

    setlocale( LC_ALL, "" );

    CWorkbook Book( "Incognito" );
    CWorksheet & Sheet = Book.AddSheet( "Images" );

    Sheet.BeginRow();
    Sheet.AddCell( "Supported image formats:" );
    Sheet.AddEmptyCells( 3 );
    Sheet.AddCell( "Anchored image (try to resize column or row):" );
    Sheet.AddEmptyCells( 5 );
    Sheet.AddCell( "Scaled image:" );
    Sheet.EndRow();

    Book.AddImage( Sheet, "Image.gif", DrawingPoint( 0, 1 ) );
    Book.AddImage( Sheet, "Image.jpg", DrawingPoint( 0, 4 ) );
    Book.AddImage( Sheet, "Image.png", DrawingPoint( 0, 7 ) );
    Book.AddImage( Sheet, "Image.tif", DrawingPoint( 0, 10 ) );

    Book.AddImage( Sheet, "Anchored.png", DrawingPoint( 4, 1 ), DrawingPoint( 5, 3 ) );

    Book.AddImage( Sheet, "Image.png", DrawingPoint( 10, 1 ), 200, 200 );

    if( Book.Save( "Images.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;

    return 0;
}
