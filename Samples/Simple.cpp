#include <cstdio>
#include <cstdlib>
#include <ctime>
#include <iostream>
#include <vector>

#include <Xlsx/Workbook.h>

#ifdef _WIN32
#include <windows.h>
#endif

#ifdef QT_CORE_LIB
#include <QDateTime>
#endif

using namespace SimpleXlsx;

int main()
{
    setlocale( LC_ALL, "" );
    time_t CurTime = time( NULL );

    CWorkbook book( "Incognito" );
    std::vector<ColumnWidth> ColWidth;
    ColWidth.push_back( ColumnWidth( 0, 3, 25 ) );
    CWorksheet & Sheet = book.AddSheet( "Unicode", ColWidth );

    Style style;
    style.horizAlign = ALIGN_H_CENTER;
    style.font.attributes = FONT_BOLD;
    const size_t CenterStyleIndex = book.AddStyle( style );

    Sheet.BeginRow();
    Sheet.AddCell( "Common test of Unicode support", CenterStyleIndex );
    Sheet.MergeCells( CellCoord( 1, 0 ), CellCoord( 1, 3 ) );
    Sheet.EndRow();

    Font TmpFont = book.GetFonts().front();
    TmpFont.attributes = FONT_ITALIC;
    Comment Com;
    Com.x = 300;
    Com.y = 100;
    Com.width = 100;
    Com.height = 30;
    Com.cellRef = CellCoord( 8, 1 );
    Com.isHidden = false;
    Com.AddContent( TmpFont, "Comment with custom style" );
    Sheet.AddComment( Com );

    Sheet.BeginRow().AddCell( "English language" ).AddCell( "English language" ).EndRow();
    Sheet.BeginRow().AddCell( "Russian language" ).AddCell( L"Русский язык" ).EndRow();
    Sheet.BeginRow().AddCell( "Chinese language" ).AddCell( L"中文" ).EndRow();
    Sheet.BeginRow().AddCell( "French language" ).AddCell( L"le français" ).EndRow();
    Sheet.BeginRow().AddCell( "Arabic language" ).AddCell( L"العَرَبِيَّة‎‎" ).EndRow();

    Sheet.AddEmptyRow();

    style.fill.patternType = PATTERN_NONE;
    style.font.theme = true;
    style.horizAlign = ALIGN_H_RIGHT;
    style.vertAlign = ALIGN_V_CENTER;

    style.numFormat.numberStyle = NUMSTYLE_MONEY;
    const size_t MoneyStyleIndex = book.AddStyle( style );
    Sheet.BeginRow().AddCell( "Money symbol" ).AddCell( 123.45, MoneyStyleIndex ).EndRow();

    Style stPanel;
    stPanel.border.top.style = BORDER_THIN;
    stPanel.border.bottom.color = "FF000000";
    stPanel.fill.patternType = PATTERN_SOLID;
    stPanel.fill.fgColor = "FFCCCCFF";
    const size_t PanelStyleIndex = book.AddStyle( stPanel );
    Sheet.AddEmptyRow().BeginRow();
    Sheet.AddCell( "Cells with border", PanelStyleIndex );
    Sheet.AddCell( "", PanelStyleIndex ).AddCell( "", PanelStyleIndex ).AddCell( "", PanelStyleIndex );
    Sheet.EndRow();

    style.numFormat.numberStyle = NUMSTYLE_DATETIME;
    style.font.attributes = FONT_NORMAL;
    style.horizAlign = ALIGN_H_LEFT;
    const size_t DateTimeStyleIndex = book.AddStyle( style );
    Sheet.AddEmptyRow().AddSimpleRow( "time_t", CenterStyleIndex );
    Sheet.AddSimpleRow( CellDataTime( CurTime, DateTimeStyleIndex ) );

    Style stRotated;
    stRotated.horizAlign = EAlignHoriz::ALIGN_H_CENTER;
    stRotated.vertAlign = EAlignVert::ALIGN_V_CENTER;
    stRotated.textRotation = 45;
    const size_t RotatedStyleIndex = book.AddStyle( stRotated );
    Sheet.AddSimpleRow( "Rotated text", RotatedStyleIndex, 3, 20 );
    Sheet.MergeCells( CellCoord( 14, 3 ), CellCoord( 19, 3 ) );

    /* Be careful with the style of date and time.
     * If milliseconds are specified, then the style used should take them into account.
     * Otherwise, Excel will round milliseconds and may change seconds.
     * See example below. */
    style.numFormat.formatString = "yyyy.mm.dd hh:mm:ss.000";
    const size_t CustomDateTimeStyleIndex = book.AddStyle( style );
    Sheet.AddSimpleRow( "Direct date and time", CenterStyleIndex );
    Sheet.AddSimpleRow( CellDataTime( 2020, 1, 1, 0, 0, 0, 500, CustomDateTimeStyleIndex ) );
    Sheet.AddSimpleRow( CellDataTime( 2020, 1, 1, 0, 0, 0, 500, DateTimeStyleIndex ) );
    Sheet.AddSimpleRow( CellDataTime( 2020, 1, 1, 0, 0, 0, 499, CustomDateTimeStyleIndex ) );
    Sheet.AddSimpleRow( CellDataTime( 2020, 1, 1, 0, 0, 0, 499, DateTimeStyleIndex ) );

#ifdef _WIN32
    Sheet.AddEmptyRow().AddSimpleRow( "Windows SYSTEMTIME", CenterStyleIndex );
    SYSTEMTIME lt;
    GetLocalTime( & lt );
    Sheet.AddSimpleRow( CellDataTime( lt, CustomDateTimeStyleIndex ) );
#endif

#if defined( QT_VERSION ) && ( QT_VERSION >= 0x040000 )
    Sheet.AddEmptyRow().AddSimpleRow( "Qt QDateTime", CenterStyleIndex );
    const QDateTime CurDT = QDateTime::currentDateTime();
    Sheet.AddSimpleRow( CellDataTime( CurDT, CustomDateTimeStyleIndex ) );
#endif

    if( book.Save( "Simple.xlsx" ) ) std::cout << "The book has been saved successfully" << std::endl;
    else std::cout << "The book saving has been failed" << std::endl;
    return 0;
}
