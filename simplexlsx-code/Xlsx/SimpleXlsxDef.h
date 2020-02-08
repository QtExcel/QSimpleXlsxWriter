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

#ifndef XLSX_SIMPLE_XLSX_DEF_H
#define XLSX_SIMPLE_XLSX_DEF_H

#include <stdint.h>
#include <cassert>
#include <ctime>
#include <fstream>
#include <list>
#include <map>
#include <sstream>
#include <string>
#include <vector>
#include <utility>

#ifdef _WIN32
#include <windows.h>
#endif

#ifdef QT_CORE_LIB
#include <QDateTime>
#endif

#include "../UTF8Encoder.hpp"

#define SIMPLE_XLSX_VERSION	"0.34"

namespace SimpleXlsx
{
// Helper class for simultaneous work with std::string and std::wstring
class UniString
{
    public:
        UniString() {}

        UniString( const char * Str ) : m_string( Str ), m_wstring( m_string.begin(), m_string.end() ) {}
        UniString( const std::string & Str ) : m_string( Str ), m_wstring( m_string.begin(), m_string.end() ) {}

        UniString( const wchar_t * Str ) : m_wstring( Str )
        {
            m_string = UTF8Encoder::From_wstring( m_wstring );
        }
        UniString( const std::wstring & Str ) : m_wstring( Str )
        {
            m_string = UTF8Encoder::From_wstring( m_wstring );
        }

        // *INDENT-OFF*   For AStyle tool
        inline bool empty() const   {   return m_string.empty();    }

        inline operator const std::string & () const    {   return m_string;    }
        inline operator const std::wstring & () const   {   return m_wstring;   }

        inline const std::string & toStdString() const      {   return m_string;    }
        inline const std::wstring & toStdWString() const    {   return m_wstring;   }

        inline bool operator==( const std::string & other ) const   {   return m_string == other;   }
        inline bool operator!=( const std::string & other ) const   {   return !( *this == other ); }
        inline bool operator==( const std::wstring & other ) const  {   return m_wstring == other;  }
        inline bool operator!=( const std::wstring & other ) const  {   return !( *this == other ); }
        inline bool operator==( const UniString & other ) const     {   return *this == other.m_string; }
        inline bool operator!=( const UniString & other ) const     {   return !( *this == other ); }
        // *INDENT-ON*   For AStyle tool

        UniString & operator=( const char * other )
        {
            m_string = other;
            m_wstring = std::wstring( m_string.begin(), m_string.end() );
            return * this;
        }
        UniString & operator=( const std::string & other )
        {
            m_string = other;
            m_wstring = std::wstring( m_string.begin(), m_string.end() );
            return * this;
        }
        UniString & operator=( const wchar_t * other )
        {
            m_wstring = other;
            m_string = UTF8Encoder::From_wstring( other );
            return * this;
        }
        UniString & operator=( const std::wstring & other )
        {
            m_wstring = other;
            m_string = UTF8Encoder::From_wstring( other );
            return * this;
        }

        friend std::ostream & operator<<( std::ostream & os, const UniString & str );

    private:
        std::string     m_string;
        std::wstring    m_wstring;
};

inline std::ostream & operator<<( std::ostream & os, const UniString & str )
{
    os << str.m_string;
    return os;
}

/// @brief	Possible chart types
enum EChartTypes
{
    CHART_NONE = -1,
    CHART_LINEAR = 0,
    CHART_BAR,
    CHART_SCATTER,
    CHART_PIE,
};

/// @brief  Possible border attributes
enum EBorderStyle
{
    BORDER_NONE = 0,
    BORDER_THIN,
    BORDER_MEDIUM,
    BORDER_DASHED,
    BORDER_DOTTED,
    BORDER_THICK,
    BORDER_DOUBLE,
    BORDER_HAIR,
    BORDER_MEDIUM_DASHED,
    BORDER_DASH_DOT,
    BORDER_MEDIUM_DASH_DOT,
    BORDER_DASH_DOT_DOT,
    BORDER_MEDIUM_DASH_DOT_DOT,
    BORDER_SLANT_DASH_DOT
};

/// @brief  Additional font attributes
enum EFontAttributes
{
    FONT_NORMAL     = 0,
    FONT_BOLD       = 1,
    FONT_ITALIC     = 2,
    FONT_UNDERLINED = 4,
    FONT_STRIKE     = 8,
    FONT_OUTLINE    = 16,
    FONT_SHADOW     = 32,
    FONT_CONDENSE   = 64,
    FONT_EXTEND     = 128,
};

/// @brief  Fill`s pattern type enumeration
enum EPatternType
{
    PATTERN_NONE = 0,
    PATTERN_SOLID,
    PATTERN_MEDIUM_GRAY,
    PATTERN_DARK_GRAY,
    PATTERN_LIGHT_GRAY,
    PATTERN_DARK_HORIZ,
    PATTERN_DARK_VERT,
    PATTERN_DARK_DOWN,
    PATTERN_DARK_UP,
    PATTERN_DARK_GRID,
    PATTERN_DARK_TRELLIS,
    PATTERN_LIGHT_HORIZ,
    PATTERN_LIGHT_VERT,
    PATTERN_LIGHT_DOWN,
    PATTERN_LIGHT_UP,
    PATTERN_LIGHT_GRID,
    PATTERN_LIGHT_TRELLIS,
    PATTERN_GRAY_125,
    PATTERN_GRAY_0625
};

/// @brief	Text horizontal alignment enumerations
enum EAlignHoriz
{
    ALIGN_H_NONE = 0,
    ALIGN_H_LEFT,
    ALIGN_H_CENTER,
    ALIGN_H_RIGHT
};

/// @brief	Text vertical alignment enumerations
enum EAlignVert
{
    ALIGN_V_NONE = 0,
    ALIGN_V_TOP,
    ALIGN_V_CENTER,
    ALIGN_V_BOTTOM
};

/// @brief	Number styling most general enumeration
enum ENumericStyle
{
    NUMSTYLE_GENERAL = 0,
    NUMSTYLE_NUMERIC,
    NUMSTYLE_PERCENTAGE,
    NUMSTYLE_EXPONENTIAL,
    NUMSTYLE_FINANCIAL,
    NUMSTYLE_MONEY,
    NUMSTYLE_DATE,
    NUMSTYLE_TIME,
    NUMSTYLE_DATETIME,
};

/// @brief	Number can be colored differently depending on whether it positive or negavite or zero
enum ENumericStyleColor
{
    NUMSTYLE_COLOR_DEFAULT = 0,
    NUMSTYLE_COLOR_BLACK,
    NUMSTYLE_COLOR_GREEN,
    NUMSTYLE_COLOR_WHITE,
    NUMSTYLE_COLOR_BLUE,
    NUMSTYLE_COLOR_MAGENTA,
    NUMSTYLE_COLOR_YELLOW,
    NUMSTYLE_COLOR_CYAN,
    NUMSTYLE_COLOR_RED
};

/// @brief  Font describes a font that can be added into final document stylesheet
/// @see    EFontAttributes
class Font
{
    public:
        UniString name;		///< font name (there is no enumeration or preset values, it should be used carefully)
        std::string color;	///< color format: AARRGGBB - (alpha, red, green, blue). If empty default theme is used
        int32_t size;		///< font size
        int32_t attributes;	///< combination of additinal font flags (EFontAttributes)
        bool theme;			///< theme if true then color is not taken into account

    public:
        Font()
        {
            Clear();
        }

        void Clear();

        bool operator==( const Font & _font ) const;
};

/// @brief  Fill describes a fill that can be added into final document stylesheet
/// @note   Current version describes the pattern fill only
/// @see    EPatternType
class Fill
{
    public:
        EPatternType patternType;	///< patternType
        std::string fgColor;		///< fgColor foreground color format: AARRGGBB - (alpha, red, green, blue). Can be left unset
        std::string bgColor;		///< bgColor background color format: AARRGGBB - (alpha, red, green, blue). Can be left unset

    public:
        Fill() : patternType( PATTERN_NONE ), fgColor( "" ), bgColor( "" ) {}

        void Clear();

        bool operator==( const Fill & _fill ) const;
};

/// @brief  Border describes a border style that can be added into final document stylesheet
/// @see    EBorderStyle
class Border
{
    public:
        /// @brief	BorderItem describes border items (e.g. left, right, bottom, top, diagonal sides)
        struct BorderItem
        {
            EBorderStyle style;		///< style border style
            std::string color;		///< colour border colour format: AARRGGBB - (alpha, red, green, blue). Can be left unset

            BorderItem() : style( BORDER_NONE ), color( "" ) {}

            void Clear()
            {
                style = BORDER_NONE;
                color = "";
            }

            bool operator==( const BorderItem & _borderItem ) const
            {
                return ( ( color == _borderItem.color ) && ( style == _borderItem.style ) );
            }
        };

    public:
        BorderItem left;			///< left left side style
        BorderItem right;			///< right right side style
        BorderItem bottom;			///< bottom bottom side style
        BorderItem top;				///< top top side style
        BorderItem diagonal;		///< diagonal diagonal side style
        bool isDiagonalUp;			///< isDiagonalUp indicates whether this diagonal border should be used
        bool isDiagonalDown;		///< isDiagonalDown indicates whether this diagonal border should be used

    public:
        Border()
        {
            Clear();
        }

        void Clear();

        bool operator==( const Border & _border ) const;
};

/// @brief	NumFormat helps to create a customized number format
class NumFormat
{
    public:
        mutable size_t id;					///< id internal format number (must not be changed manually)

        std::string formatString;				///< formatString final format form. If set manually, then all other options are ignored

        ENumericStyle numberStyle;			///< numberStyle most general style definition from enumeration
        ENumericStyleColor positiveColor;	///< positiveColor positive numbers color (switched off by default)
        ENumericStyleColor negativeColor;	///< negativeColor negative numbers color (switched off by default)
        ENumericStyleColor zeroColor;		///< zeroColor zero color (switched off by default)
        bool showThousandsSeparator;		///< thousandsSeparator specifies whether to show thousands separator
        size_t numberOfDigitsAfterPoint;	///< numberOfDigitsAfterPoint number of digits after point

    public:
        NumFormat()
        {
            Clear();
        }

        void Clear();

        bool operator==( const NumFormat & _num ) const;
};

/// @brief	Cell coordinate structure
class CellCoord
{
        static const uint8_t ColStrOffset = 11;

    public:
        typedef char TConvBuf[ 24 ];    // Max string for $Col$Row\0
        static const uint32_t   MaxRows = 1048576, MaxCols = 16384; // Excel limits
        uint32_t row;	///< row (starts from 1)
        uint32_t col;	///< col (starts from 0)

        CellCoord() : row( 1 ), col( 0 ) {}
        CellCoord( uint32_t _r, uint32_t _c ) : row( _r ), col( _c ) {}

        void Clear()
        {
            row = 1;
            col = 0;
        }

        // Transforms the row and the column numerics from uint32_t to coordinate string format
        std::string ToString() const;
        char * ToString( TConvBuf & Buffer ) const;
        // Transforms the row and the column numerics from uint32_t to coordinate string format with '$' symbols
        std::string ToStringAbs() const;
        char * ToStringAbs( TConvBuf & Buffer ) const;

    private:
        template< bool AbsColAndRow >
        inline char * ToString( char * Buffer ) const
        {
            const int32_t iAlphLen = 26;
            const char * szAlph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            Buffer[ CellCoord::ColStrOffset ] = szAlph[ col % iAlphLen ];
            char * StartPtr = Buffer + CellCoord::ColStrOffset;
            int32_t div = col / iAlphLen;
            while( div != 0 )
            {
                div--;
                StartPtr--;
                * StartPtr = szAlph[ div % iAlphLen ];
                div /= iAlphLen;
            }
            if( AbsColAndRow )
            {
                StartPtr--;
                * StartPtr = '$';
                Buffer[ CellCoord::ColStrOffset + 1 ] = '$';
            }
            sprintf( Buffer + CellCoord::ColStrOffset + ( AbsColAndRow ? 2 : 1 ), "%u", row );
            return StartPtr;
        }
};

/// @brief	Column width structure
class ColumnWidth
{
    public:
        uint32_t colFrom;	///< column range from (starts from 0)
        uint32_t colTo;		///< column range to (starts from 0)
        float width;		///< specified width

    public:
        ColumnWidth() : colFrom( 0 ), colTo( 0 ), width( 15 ) {}
        ColumnWidth( uint32_t min, uint32_t max, float w ) : colFrom( min ), colTo( max ), width( w ) {}
};

class CellDataStr
{
    public:
        std::string value;
        size_t style_id;

    public:
        CellDataStr() : value( "" ), style_id( 0 ) {}
        CellDataStr( const char * pStr ) : value( pStr ), style_id( 0 ) {}
        CellDataStr( const std::string & _str ) : value( _str ), style_id( 0 ) {}
        CellDataStr( const std::string & _str, size_t _style_id ) : value( _str ), style_id( _style_id ) {}

        CellDataStr & operator=( const CellDataStr & obj )
        {
            value = obj.value;
            style_id = obj.style_id;
            return *this;
        }

        CellDataStr & operator=( const char * pStr )
        {
            value = pStr;
            return *this;
        }

        CellDataStr & operator=( const std::string & _str )
        {
            value = _str;
            return *this;
        }

        CellDataStr( const std::wstring & _str ) : value( UTF8Encoder::From_wstring( _str ) ), style_id( 0 ) {}
        CellDataStr( const std::wstring & _str, size_t _style_id ) : value( UTF8Encoder::From_wstring( _str ) ), style_id( _style_id ) {}

        CellDataStr & operator=( const wchar_t * pStr )
        {
            value = UTF8Encoder::From_wstring( pStr );
            return *this;
        }

        CellDataStr & operator=( const std::wstring & _str )
        {
            value = UTF8Encoder::From_wstring( _str );
            return *this;
        }
};	///< cell data:style pair

class CellDataTime
{
    public:
        size_t style_id;

        inline CellDataTime() : style_id( 0 ), m_xlsx_val( 0.0 ) {}
        inline CellDataTime( time_t _val, size_t _style_id = 0 ) : style_id( _style_id ), m_xlsx_val( From_time_t( _val ) ) {}

        // Year must be in the range 1900 to 9999, month must be in the range 1 to 12, and day must be in the range 1 to 31.
        // Hour must be in the range 0 to 23, minute and second must be in the range 0 to 59, and millisecond must be in the range 0 to 999.
        inline CellDataTime( uint16_t year, uint16_t month, uint16_t day, size_t _style_id = 0 ) :
            style_id( _style_id ), m_xlsx_val( FromGregorian( year, month, day, 0, 0, 0 ) ) {}
        inline CellDataTime( uint16_t year, uint16_t month, uint16_t day, uint16_t hour, uint16_t minute, uint16_t second, uint16_t millisecond = 0, size_t _style_id = 0 ) :
            style_id( _style_id ), m_xlsx_val( FromGregorian( year, month, day, hour, minute, second, millisecond ) ) {}

        inline CellDataTime & operator=( const CellDataTime & obj )
        {
            style_id = obj.style_id;
            m_xlsx_val = obj.m_xlsx_val;
            return *this;
        }

        inline CellDataTime & operator=( time_t _val )
        {
            return SetDateTime( _val );
        }

        inline CellDataTime & SetDateTime( time_t val )
        {
            m_xlsx_val = From_time_t( val );
            return * this;
        }
        // Year must be in the range 1900 to 9999, month must be in the range 1 to 12, and day must be in the range 1 to 31.
        // Hour must be in the range 0 to 23, minute and second must be in the range 0 to 59, and millisecond must be in the range 0 to 999.
        inline CellDataTime & SetDateTime( uint16_t year, uint8_t month, uint8_t day,
                                           uint8_t hour = 0, uint8_t minute = 0, uint8_t second = 0, uint16_t millisecond = 0 )
        {
            m_xlsx_val = FromGregorian( year, month, day, hour, minute, second, millisecond );
            return * this;
        }

        inline double XlsxValue() const
        {
            return m_xlsx_val;
        }

#ifdef _WIN32
        inline CellDataTime( const SYSTEMTIME & win_st, size_t _style_id = 0 ) : style_id( _style_id ), m_xlsx_val( FromWinSystemTime( win_st ) ) {}

        // *INDENT-OFF*   For AStyle tool
        inline CellDataTime & SetDateTime( const SYSTEMTIME & win_st )  { m_xlsx_val = FromWinSystemTime( win_st ); return * this; }
        inline CellDataTime & operator=( const SYSTEMTIME & win_st )    { return SetDateTime( win_st ); }
        // *INDENT-ON*   For AStyle tool
#endif

#if defined( QT_VERSION ) && ( QT_VERSION >= 0x040000 )
        inline CellDataTime( const QTime & t, size_t _style_id = 0 ) : style_id( _style_id ), m_xlsx_val( FromQTime( t ) ) {}
        inline CellDataTime( const QDate & d, size_t _style_id = 0 ) : style_id( _style_id ), m_xlsx_val( FromQDate( d ) ) {}
        inline CellDataTime( const QDateTime & dt, size_t _style_id = 0 ) : style_id( _style_id ), m_xlsx_val( FromQDateTime( dt ) ) {}

        // *INDENT-OFF*   For AStyle tool
        inline CellDataTime & SetTime( const QTime & t )            { m_xlsx_val = FromQTime( t );      return * this; }
        inline CellDataTime & SetDate( const QDate & d )            { m_xlsx_val = FromQDate( d );      return * this; }
        inline CellDataTime & SetDateTime( const QDateTime & dt )   { m_xlsx_val = FromQDateTime( dt ); return * this; }

        inline CellDataTime & operator=( const QTime & t )      { return SetTime( t ); }
        inline CellDataTime & operator=( const QDate & d )      { return SetDate( d ); }
        inline CellDataTime & operator=( const QDateTime & dt ) { return SetDateTime( dt ); }
        // *INDENT-ON*   For AStyle tool
#endif

    private:
        double m_xlsx_val;

        static double From_time_t( time_t val );
        // Year must be in the range 1900 to 9999, month must be in the range 1 to 12, and day must be in the range 1 to 31.
        // Hour must be in the range 0 to 23, minute and second must be in the range 0 to 59, and millisecond must be in the range 0 to 999.
        static double FromGregorian( uint16_t year, uint16_t month, uint16_t day, uint16_t hour, uint16_t minute, uint16_t second, uint16_t millisecond = 0 );

#ifdef _WIN32
        static inline double FromWinSystemTime( const SYSTEMTIME & win_st )
        {
            return FromGregorian( win_st.wYear, win_st.wMonth, win_st.wDay, win_st.wHour, win_st.wMinute, win_st.wSecond, win_st.wMilliseconds );
        }
#endif

#if defined( QT_VERSION ) && ( QT_VERSION >= 0x040000 )
        static inline double FromQTime( const QTime & t )
        {
            assert( t.isValid() );
            return FromGregorian( 1900, 1, 1, t.hour(), t.minute(), t.second(), t.msec() );
        }
        static inline double FromQDate( const QDate & d )
        {
            assert( d.isValid() );
            return FromGregorian( d.year(), d.month(), d.day(), 0, 0, 0 );
        }
        static inline double FromQDateTime( const QDateTime & dt )
        {
            assert( dt.isValid() );
            const QDate d = dt.date();
            const QTime t = dt.time();
            return FromGregorian( d.year(), d.month(), d.day(), t.hour(), t.minute(), t.second(), t.msec() );
        }
#endif
};	///< cell data:style pair

class CellDataInt
{
    public:
        int32_t value;
        size_t style_id;

    public:
        CellDataInt() : value( 0 ), style_id( 0 ) {}
        CellDataInt( int32_t _val ) : value( _val ), style_id( 0 ) {}
        CellDataInt( int32_t _val, size_t _style_id ) : value( _val ), style_id( _style_id ) {}

        CellDataInt & operator=( const CellDataInt & obj )
        {
            value = obj.value;
            style_id = obj.style_id;
            return *this;
        }

        CellDataInt & operator=( int32_t _val )
        {
            value = _val;
            return *this;
        }
};	///< cell data:style pair
class CellDataUInt
{
    public:
        uint32_t value;
        size_t style_id;

    public:
        CellDataUInt() : value( 0 ), style_id( 0 ) {}
        CellDataUInt( uint32_t _val ) : value( _val ), style_id( 0 ) {}
        CellDataUInt( uint32_t _val, size_t _style_id ) : value( _val ), style_id( _style_id ) {}

        CellDataUInt & operator=( const CellDataUInt & obj )
        {
            value = obj.value;
            style_id = obj.style_id;
            return *this;
        }

        CellDataUInt & operator=( uint32_t _val )
        {
            value = _val;
            return *this;
        }
};	///< cell data:style pair
class CellDataDbl
{
    public:
        double value;
        size_t style_id;

    public:
        CellDataDbl() : value( 0.0 ), style_id( 0 ) {}
        CellDataDbl( double _val ) : value( _val ), style_id( 0 ) {}
        CellDataDbl( double _val, size_t _style_id ) : value( _val ), style_id( _style_id ) {}

        CellDataDbl & operator=( const CellDataDbl & obj )
        {
            value = obj.value;
            style_id = obj.style_id;
            return *this;
        }

        CellDataDbl & operator=( double _val )
        {
            value = _val;
            return *this;
        }
};	///< cell data:style pair
class CellDataFlt
{
    public:
        float value;
        size_t style_id;

    public:
        CellDataFlt() : value( 0.0 ), style_id( 0 ) {}
        CellDataFlt( float _val ) : value( _val ), style_id( 0 ) {}
        CellDataFlt( float _val, size_t _style_id ) : value( _val ), style_id( _style_id ) {}

        CellDataFlt & operator=( const CellDataFlt & obj )
        {
            value = obj.value;
            style_id = obj.style_id;
            return *this;
        }

        CellDataFlt & operator=( float _val )
        {
            value = _val;
            return *this;
        }
};	///< cell data:style pair

/// @brief	This structure describes comment item that can added to a cell
struct Comment
{
    size_t sheetIndex;									///< sheetIndex internal page index (must not be changed manually)
    std::list<std::pair<Font, UniString> > contents;    ///< contents set of contents with specified fonts
    CellCoord cellRef;									///< cellRef reference to the cell
    std::string fillColor;								///< fillColor comment box background colour (format: #RRGGBB)
    bool isHidden;										///< isHidden determines if comments box is hidden
    int x;												///< x absolute position in pt (can be left unset)
    int y;												///< y absolute position in pt (can be left unset)
    int width;											///< width width in pt (can be left unset)
    int height;											///< height height in pt (can be left unset)

    Comment()
    {
        Clear();
    }

    void Clear();

    void AddContent( const Font & afont, const UniString & astring )
    {
        contents.push_back( std::pair<Font, UniString>( afont, astring ) );
    }

    bool operator < ( const Comment & _comm ) const
    {
        return ( sheetIndex < _comm.sheetIndex );
    }
};

/// @brief  Style describes a set of styling parameter that can be used into final document
/// @see    EBorder
/// @see    EAlignHoriz
/// @see    EAlignVert
class Style
{
    public:
        Font font;				///< font structure object describes font
        Fill fill;				///< fill structure object describes fill
        Border border;			///< border combination of border attributes
        NumFormat numFormat;	///< numFormat structure object describes numeric cell representation
        EAlignHoriz horizAlign;	///< horizAlign cell content horizontal alignment value
        EAlignVert vertAlign;	///< vertAlign cell content vertical alignment value
        bool wrapText;			///< wrapText text wrapping property
        int textRotation;       ///< textRotation angle in degree from -90 to 90

    public:
        Style()
        {
            Clear();
        }

        void Clear();
};

/// @brief  This structure represents a handle to manage newly adding styles to avoid dublicating
class StyleList
{
    public:
        static const size_t BUILT_IN_STYLES_NUMBER = 164;
        static const size_t STYLE_LINK_NUMBER = 4;

        enum
        {
            STYLE_LINK_BORDER = 0,
            STYLE_LINK_FONT,
            STYLE_LINK_FILL,
            STYLE_LINK_NUM_FORMAT
        };

        struct StylePosInfo
        {
            EAlignHoriz horizAlign;
            EAlignVert  vertAlign;
            bool        wrapText;
            int         textRotation;

            inline bool isAllDefault() const
            {
                return ( horizAlign == ALIGN_H_NONE ) && ( vertAlign == ALIGN_V_NONE ) && ! wrapText && ( textRotation == 0 );
            }
        };

    private:
        size_t m_fmtLastId;				///< m_fmtLastId format counter. There are 164 (0 - 163) built-in numeric formats

        std::vector<Border> m_borders;	///< borders set of values represent styled borders
        std::vector<Font> m_fonts;		///< fonts set of fonts to be declared
        std::vector<Fill> m_fills;		///< fills set of fills to be declared
        std::vector<NumFormat> m_nums;	///< nums set of number formats to be declared
        std::vector<std::vector<size_t> > m_styleIndexes;///< styleIndexes vector of a number triplet contains links to style parts:
        ///         first - border id in borders
        ///         second - font id in fonts
        ///         third - fill id in fills
        ///			fourth - number format id in nums

        std::vector< StylePosInfo > m_stylePos;///< stylePos vector of a number triplet contains style`s alignments and wrap sign:

public:
        StyleList();

        // *INDENT-OFF*   For AStyle tool
        /// @brief	For internal use (at the book saving)
        inline const std::vector<Border> & GetBorders() const { return m_borders; }

        /// @brief	For internal use (at the book saving)
        inline const std::vector<Font> & GetFonts()	const { return m_fonts; }

        /// @brief	For internal use (at the book saving)
        inline const std::vector<Fill> & GetFills()	const { return m_fills; }

        /// @brief	For internal use (at the book saving)
        inline const std::vector<NumFormat> & GetNumFormats() const { return m_nums; }

        /// @brief	For internal use (at the book saving)
        inline const std::vector<std::vector<size_t> > & GetIndexes() const { return m_styleIndexes; }

        /// @brief	For internal use (at the book saving)
        inline const std::vector< StylePosInfo > & GetPositions() const { return m_stylePos; }
        // *INDENT-ON*   For AStyle tool

        /// @brief  Adds a new style into collection if it is not exists yet
        /// @param  style Reference to the Style structure object
        /// @return Style index that should be used at data appending to a data sheet
        /// @note   If returned value is 0 - this is a default normal style and it is optional
        ///         whether is will be added into column description or not
        ///         (but better not to add to reduce size and resource consumption)
        size_t Add( const Style & style );
};

//Struct for the drawing (chart, image) position description
struct DrawingPoint
{
    uint32_t col;	///< Column (starts from 0)
    uint32_t colOff;///< Column Offset (starts from 0)
    uint32_t row;	///< Row (starts from 0)
    uint32_t rowOff;///< Row Offset (starts from 0)

    DrawingPoint() {}
    DrawingPoint( uint32_t acol, uint32_t arow ) :
        col( acol ), colOff( 0 ), row( arow ), rowOff( 0 ) {}
    DrawingPoint( uint32_t acol, uint32_t acolOff, uint32_t arow, uint32_t arowOff ) :
        col( acol ), colOff( acolOff ), row( arow ), rowOff( arowOff ) {}
};

//Class for image description
class CImage
{
    public:
        enum ImageType
        {
            unknown, gif, jpg, jpeg, png, tif, tiff
        };

        std::string LocalPath;      //Path to the file in the system
        std::string InternalName;   //Name of the file in XLSX
        ImageType   Type;           //Image type (extension)
        uint16_t    Width;          //Width of the image
        uint16_t    Height;         //Height of the image

        static const uint16_t   PointByPixel = 9525;    // Points count by one pixel

        CImage( const std::string & LocPath, const std::string & TempPath, ImageType AType, uint16_t AWidth, uint16_t AHeight ) :
            LocalPath( LocPath ), InternalName( TempPath ), Type( AType ), Width( AWidth ), Height( AHeight ) {}
};

// Base class for CWorksheet and CChartsheet
class CSheet
{
    public:
        virtual ~CSheet();
        virtual const UniString & GetTitle() const = 0;

        inline size_t GetIndex() const
        {
            return m_index;
        }

    protected:
        const size_t  m_index;  ///< current sheet number

        CSheet( size_t SheetIndex ) : m_index( SheetIndex ) {}
};

}	// namespace SimpleXlsx

#endif // XLSX_SIMPLE_XLSX_DEF_H
