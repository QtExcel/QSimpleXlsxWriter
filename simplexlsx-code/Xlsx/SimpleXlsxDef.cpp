/*
  SimpleXlsxWriter
  Copyright (C) 2012-2020 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

  This software is provided 'as-is', without any express or implied
  warranty. In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.
*/

#include "SimpleXlsxDef.h"

void SimpleXlsx::Font::Clear()
{
    size = 11;
    name = "Calibri";
    theme = true;
    color = "";
    attributes = FONT_NORMAL;
}

bool SimpleXlsx::Font::operator==( const SimpleXlsx::Font & _font ) const
{
    return ( ( size == _font.size ) && ( name == _font.name ) && ( theme == _font.theme ) &&
             ( color == _font.color ) && ( attributes == _font.attributes ) );
}

void SimpleXlsx::Fill::Clear()
{
    patternType = PATTERN_NONE;
    fgColor = "";
    bgColor = "";
}

bool SimpleXlsx::Fill::operator==( const SimpleXlsx::Fill & _fill ) const
{
    return ( ( patternType == _fill.patternType ) && ( fgColor == _fill.fgColor ) && ( bgColor == _fill.bgColor ) );
}

void SimpleXlsx::Border::Clear()
{
    isDiagonalUp = false;
    isDiagonalDown = false;

    left.Clear();
    right.Clear();
    bottom.Clear();
    top.Clear();
}

bool SimpleXlsx::Border::operator==( const SimpleXlsx::Border & _border ) const
{
    return ( ( isDiagonalUp == _border.isDiagonalUp ) && ( isDiagonalDown == _border.isDiagonalDown ) &&
             ( left == _border.left ) && ( right == _border.right ) &&
             ( bottom == _border.bottom ) && ( top == _border.top ) );
}

void SimpleXlsx::NumFormat::Clear()
{
    id = 164;
    formatString = "";

    numberStyle = NUMSTYLE_GENERAL;
    positiveColor = NUMSTYLE_COLOR_DEFAULT;
    negativeColor = NUMSTYLE_COLOR_DEFAULT;
    zeroColor = NUMSTYLE_COLOR_DEFAULT;
    showThousandsSeparator = false;
    numberOfDigitsAfterPoint = 2;
}

bool SimpleXlsx::NumFormat::operator==( const SimpleXlsx::NumFormat & _num ) const
{
    return ( ( formatString == _num.formatString ) &&
             ( numberStyle == _num.numberStyle ) && ( numberOfDigitsAfterPoint == _num.numberOfDigitsAfterPoint ) &&
             ( positiveColor == _num.positiveColor ) && ( negativeColor == _num.negativeColor ) &&
             ( zeroColor == _num.zeroColor ) && ( showThousandsSeparator == _num.showThousandsSeparator ) );
}

void SimpleXlsx::Style::Clear()
{
    font.Clear();
    fill.Clear();
    border.Clear();
    horizAlign = ALIGN_H_NONE;
    vertAlign = ALIGN_V_NONE;
    wrapText = false;
    textRotation = 0;
}

std::string SimpleXlsx::CellCoord::ToString() const
{
    CellCoord::TConvBuf Buffer;
    return std::string( ToString< false >( Buffer ) );
}

char * SimpleXlsx::CellCoord::ToString( SimpleXlsx::CellCoord::TConvBuf & Buffer ) const
{
    return ToString< false >( Buffer );
}

std::string SimpleXlsx::CellCoord::ToStringAbs() const
{
    CellCoord::TConvBuf Buffer;
    return std::string( ToString< true >( Buffer ) );
}

char * SimpleXlsx::CellCoord::ToStringAbs( SimpleXlsx::CellCoord::TConvBuf & Buffer ) const
{
    return ToString< true >( Buffer );
}



double SimpleXlsx::CellDataTime::From_time_t( time_t val )
{
    const int64_t secondsFrom1900to1970 = 2208988800u;
    const double excelOneSecond = 0.0000115740740740741;
    struct tm * t = localtime( & val );

    time_t timeSinceEpoch = t->tm_sec + t->tm_min * 60 + t->tm_hour * 3600 + t->tm_yday * 86400 +
                            ( t->tm_year - 70 ) * 31536000 + ( ( t->tm_year - 69 ) / 4 ) * 86400 -
                            ( ( t->tm_year - 1 ) / 100 ) * 86400 + ( ( t->tm_year + 299 ) / 400 ) * 86400;

    return excelOneSecond * ( secondsFrom1900to1970 + timeSinceEpoch ) + 2;
}

double SimpleXlsx::CellDataTime::FromGregorian( uint16_t year, uint16_t month, uint16_t day, uint16_t hour, uint16_t minute, uint16_t second, uint16_t millisecond )
{
    const bool IsLeapYear = ( ( year & 0x0003 ) == 0 ) && ( year % 400 != 0 );
    const uint8_t DM[ 12 ] = { 31, uint8_t( IsLeapYear ? 29 : 28 ), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31  };
    const double excelOneSecond = 0.0000115740740740741;
    assert( ( year >= 1900 ) && ( year < 10000 ) ); // Excel restrictions
    assert( ( month >= 1 ) && ( month <= 12 ) && ( day >= 1 ) && ( day <= DM[ month - 1 ] ) );
    assert( ( hour <= 23 ) && ( minute <= 59 ) && ( second <= 59 ) && ( millisecond <= 999 ) );
    int PrevMonthDays = 0;
    for( int i = 0; i < month - 1; ++i )
        PrevMonthDays += DM[ i ];
    const int64_t YearsSince1900 = year - 1900;
    const int64_t LeapYears = ( YearsSince1900 - 1 ) / 4 + ( YearsSince1900 != 0 ? 1 : 0 );
    double TotalSeconds = second + minute * 60 + hour * 3600 + YearsSince1900 * 31536000 + ( PrevMonthDays + day - 1 + LeapYears ) * 86400;
    return excelOneSecond * ( TotalSeconds + millisecond / 1000.0 ) + 1.0;  // + 1.0 for 1900.01.01
}

void SimpleXlsx::Comment::Clear()
{
    contents.clear();
    cellRef.Clear();
    fillColor = "#FFEFD5";	// papaya whip
    isHidden = true;
    x = y = 50;
    width = height = 100;
}


SimpleXlsx::StyleList::StyleList()
{
    m_fmtLastId = BUILT_IN_STYLES_NUMBER;

    m_borders.clear();
    m_fonts.clear();
    m_fills.clear();
    m_nums.clear();
    m_styleIndexes.clear();
    m_stylePos.clear();
}

size_t SimpleXlsx::StyleList::Add( const SimpleXlsx::Style & style )
{
    std::vector<size_t> styleLinks( STYLE_LINK_NUMBER );

    // Check border existance
    bool addItem = true;
    for( size_t i = 0; i < m_borders.size(); i++ )
    {
        if( m_borders[i] == style.border )
        {
            addItem = false;
            styleLinks[STYLE_LINK_BORDER] = i;
            break;
        }
    }

    // Add border if it is not in collection yet
    if( addItem )
    {
        m_borders.push_back( style.border );
        styleLinks[STYLE_LINK_BORDER] = m_borders.size() - 1;
    }

    // Check font existance
    addItem = true;
    for( size_t i = 0; i < m_fonts.size(); i++ )
    {
        if( m_fonts[i] == style.font )
        {
            addItem = false;
            styleLinks[STYLE_LINK_FONT] = i;
            break;
        }
    }

    // Add font if it is not in collection yet
    if( addItem )
    {
        m_fonts.push_back( style.font );
        styleLinks[STYLE_LINK_FONT] = m_fonts.size() - 1;
    }

    // Check fill existance
    addItem = true;
    for( size_t i = 0; i < m_fills.size(); i++ )
    {
        if( m_fills[i] == style.fill )
        {
            addItem = false;
            styleLinks[STYLE_LINK_FILL] = i;
            break;
        }
    }

    // Add fill if it is not in collection yet
    if( addItem )
    {
        m_fills.push_back( style.fill );
        styleLinks[STYLE_LINK_FILL] = m_fills.size() - 1;
    }

    // Check number format existance
    addItem = true;
    for( size_t i = 0; i < m_nums.size(); i++ )
    {
        if( m_nums[i] == style.numFormat )
        {
            addItem = false;
            styleLinks[STYLE_LINK_NUM_FORMAT] = m_nums[ i ].id;
            break;
        }
    }

    // Add number format if it is not in collection yet
    if( addItem )
    {
        if( style.numFormat.id >= BUILT_IN_STYLES_NUMBER )
        {
            styleLinks[STYLE_LINK_NUM_FORMAT] = m_fmtLastId;
            style.numFormat.id = m_fmtLastId++;
        }
        else
        {
            styleLinks[STYLE_LINK_NUM_FORMAT] = m_nums.size();
        }

        m_nums.push_back( style.numFormat );
    }

    // Check style combination existance
    for( size_t i = 0; i < m_styleIndexes.size(); i++ )
    {
        if( ( m_styleIndexes[i] == styleLinks ) &&
                ( m_stylePos[i].horizAlign == style.horizAlign ) &&
                ( m_stylePos[i].vertAlign == style.vertAlign ) &&
                ( m_stylePos[i].wrapText == style.wrapText ) &&
                ( m_stylePos[i].textRotation == style.textRotation ) )
            return i;
    }

    StylePosInfo pos;
    pos.horizAlign = style.horizAlign;
    pos.vertAlign = style.vertAlign;
    pos.wrapText = style.wrapText;
    pos.textRotation = style.textRotation;
    m_stylePos.push_back( pos );

    m_styleIndexes.push_back( styleLinks );
    return m_styleIndexes.size() - 1;
}


SimpleXlsx::CSheet::~CSheet() {}

