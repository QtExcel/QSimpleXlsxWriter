#include <iostream>
#include <math.h>
#include <stdio.h>
#include <stdlib.h>
#include <time.h>

#include <Xlsx/Workbook.h>

using namespace SimpleXlsx;

int main( int argc, char * argv[] )
{
    ( void )argc; ( void )argv;

    setlocale( LC_ALL, "" );

    time_t CurrentTime = time( NULL );

    CWorkbook book( "Incognito" );

    std::vector<ColumnWidth> ColWidth;
    ColWidth.push_back( ColumnWidth( 0, 3, 20 ) );
    CWorksheet & Sheet = book.AddSheet( "Unicode", ColWidth );

    Style style;
    style.horizAlign = ALIGN_H_CENTER;
    style.font.attributes = FONT_BOLD;
    size_t CenterStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Common test of Unicode support", CenterStyleIndex );
    Sheet.MergeCells( CellCoord( 1, 0 ), CellCoord( 1, 3 ) );
    Sheet.EndRow();

    Font TmpFont = book.GetFonts().front();
    TmpFont.attributes = FONT_ITALIC;
    Comment Com;
    Com.x = 250;
    Com.y = 100;
    Com.width = 100;
    Com.height = 30;
    Com.cellRef = CellCoord( 8, 1 );
    Com.isHidden = false;
    Com.AddContent( TmpFont, "Comment with custom style" );
    Sheet.AddComment( Com );

    Sheet.BeginRow();
    Sheet.AddCell( "English language" );
    Sheet.AddCell( "English language" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "Russian language" );
    Sheet.AddCell( L"Русский язык" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "Chinese language" );
    Sheet.AddCell( L"中文" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "French language" );
    Sheet.AddCell( L"le français" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.AddCell( "Arabic language" );
    Sheet.AddCell( L"العَرَبِيَّة‎‎" );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.EndRow();

    style.fill.patternType = PATTERN_NONE;
    style.font.theme = true;
    style.horizAlign = ALIGN_H_RIGHT;
    style.vertAlign = ALIGN_V_CENTER;

    style.numFormat.numberStyle = NUMSTYLE_MONEY;

    size_t MoneyStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Money symbol" );
    Sheet.AddCell( 123.45, MoneyStyleIndex );
    Sheet.EndRow();

    style.numFormat.numberStyle = NUMSTYLE_DATETIME;
    size_t DateTimeStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Write date/time" );
    Sheet.AddCell( CurrentTime, DateTimeStyleIndex );
    Sheet.EndRow();

    style.numFormat.formatString = "hh:mm:ss";
    size_t CustomDateTimeStyleIndex = book.AddStyle( style );
    Sheet.BeginRow();
    Sheet.AddCell( "Custom date/time" );
    Sheet.AddCell( CurrentTime, CustomDateTimeStyleIndex );
    Sheet.EndRow();

    Sheet.BeginRow();
    Sheet.EndRow();

    Style stPanel;
    stPanel.border.top.style = BORDER_THIN;
    stPanel.border.bottom.color = "FF000000";
    stPanel.fill.patternType = PATTERN_SOLID;
    stPanel.fill.fgColor = "FFCCCCFF";
    size_t PanelStyleIndex = book.AddStyle( stPanel );
    Sheet.BeginRow();
    Sheet.AddCell( "Cells with border", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex );
    Sheet.EndRow();

    if( book.Save( "Simple.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;

    return 0;
}
