/*
  SimpleXlsxWriter
  Copyright (C) 2012-2018 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

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
#include <ctime>
#include <fstream>
#include <list>
#include <map>
#include <sstream>
#include <string>
#include <vector>
#include <utility>

#include "../UTF8Encoder.hpp"

#define SIMPLE_XLSX_VERSION	"0.33"

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

        void Clear()
        {
            size = 11;
            name = "Calibri";
            theme = true;
            color = "";
            attributes = FONT_NORMAL;
        }

        bool operator==( const Font & _font ) const
        {
            return ( ( size == _font.size ) && ( name == _font.name ) && ( theme == _font.theme ) &&
                     ( color == _font.color ) && ( attributes == _font.attributes ) );
        }
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

        void Clear()
        {
            patternType = PATTERN_NONE;
            fgColor = "";
            bgColor = "";
        }

        bool operator==( const Fill & _fill ) const
        {
            return ( ( patternType == _fill.patternType ) && ( fgColor == _fill.fgColor ) && ( bgColor == _fill.bgColor ) );
        }
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

        void Clear()
        {
            isDiagonalUp = false;
            isDiagonalDown = false;

            left.Clear();
            right.Clear();
            bottom.Clear();
            top.Clear();
        }

        bool operator==( const Border & _border ) const
        {
            return ( ( isDiagonalUp == _border.isDiagonalUp ) && ( isDiagonalDown == _border.isDiagonalDown ) &&
                     ( left == _border.left ) && ( right == _border.right ) &&
                     ( bottom == _border.bottom ) && ( top == _border.top ) );
        }
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

        void Clear()
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

        bool operator==( const NumFormat & _num ) const
        {
            return ( ( formatString == _num.formatString ) &&
                     ( numberStyle == _num.numberStyle ) && ( numberOfDigitsAfterPoint == _num.numberOfDigitsAfterPoint ) &&
                     ( positiveColor == _num.positiveColor ) && ( negativeColor == _num.negativeColor ) &&
                     ( zeroColor == _num.zeroColor ) && ( showThousandsSeparator == _num.showThousandsSeparator ) );
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

    public:
        Style()
        {
            Clear();
        }

        void Clear()
        {
            font.Clear();
            fill.Clear();
            border.Clear();
            horizAlign = ALIGN_H_NONE;
            vertAlign = ALIGN_V_NONE;
            wrapText = false;
        }
};

/// @brief	Cell coordinate structure
class CellCoord
{
    public:
        uint32_t row;	///< row (starts from 1)
        uint32_t col;	///< col (starts from 0)

    public:
        CellCoord() : row( 1 ), col( 0 ) {}
        CellCoord( uint32_t _r, uint32_t _c ) : row( _r ), col( _c ) {}

        void Clear()
        {
            row = 1;
            col = 0;
        }

        //Transforms the row and the column numerics from int32_t to coordinate string format
        std::string ToString() const
        {
            const int32_t iAlphLen = 26;
            const char * szAlph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            std::string strCol;
            strCol.append( 1, szAlph[ col % iAlphLen] ); // last
            int32_t div = col / iAlphLen;
            while( true )
            {
                if( div == 0 ) break;

                div--;
                strCol.insert( strCol.begin(), szAlph[ div % iAlphLen ] );
                div /= iAlphLen;
            }
            //return strCol + std::to_string( row );
            std::ostringstream ConvStream;
            ConvStream << row;
            return strCol + ConvStream.str();
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
        time_t value;
        size_t style_id;

    public:
        CellDataTime() : value( 0 ), style_id( 0 ) {}
        CellDataTime( time_t _val ) : value( _val ), style_id( 0 ) {}

        CellDataTime( time_t _val, size_t _style_id ) : value( _val ), style_id( _style_id ) {}

        CellDataTime & operator=( const CellDataTime & obj )
        {
            value = obj.value;
            style_id = obj.style_id;
            return *this;
        }

        CellDataTime & operator=( time_t _val )
        {
            value = _val;
            return *this;
        }
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

    void Clear()
    {
        contents.clear();
        cellRef.Clear();
        fillColor = "#FFEFD5";	// papaya whip
        isHidden = true;
        x = y = 50;
        width = height = 100;
    }

    void AddContent( const Font & afont, const UniString & astring )
    {
        contents.push_back( std::pair<Font, UniString>( afont, astring ) );
    }

    bool operator < ( const Comment & _comm ) const
    {
        return ( sheetIndex < _comm.sheetIndex );
    }
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

    private:
        size_t fmtLastId;				///< m_fmtLastId format counter. There are 164 (0 - 163) built-in numeric formats

        std::vector<Border> borders;	///< borders set of values represent styled borders
        std::vector<Font> fonts;		///< fonts set of fonts to be declared
        std::vector<Fill> fills;		///< fills set of fills to be declared
        std::vector<NumFormat> nums;	///< nums set of number formats to be declared
        std::vector<std::vector<size_t> > styleIndexes;///< styleIndexes vector of a number triplet contains links to style parts:
        ///         first - border id in borders
        ///         second - font id in fonts
        ///         third - fill id in fills
        ///			fourth - number format id in nums
        std::vector< std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> > stylePos;///< stylePos vector of a number triplet contains style`s alignments and wrap sign:
        ///         first - EAlignHoriz value
        ///         second - EAlignVert value
        ///         third - wrap boolean value
    public:
        StyleList()
        {
            fmtLastId = BUILT_IN_STYLES_NUMBER;

            borders.clear();
            fonts.clear();
            fills.clear();
            nums.clear();
            styleIndexes.clear();
            stylePos.clear();
        }

            // *INDENT-OFF*   For AStyle tool
            /// @brief	For internal use (at the book saving)
            inline const std::vector<Border> & GetBorders() const { return borders; }

            /// @brief	For internal use (at the book saving)
            inline const std::vector<Font> & GetFonts()	const { return fonts; }

            /// @brief	For internal use (at the book saving)
            inline const std::vector<Fill> & GetFills()	const { return fills; }

            /// @brief	For internal use (at the book saving)
            inline const std::vector<NumFormat> & GetNumFormats() const { return nums; }

            /// @brief	For internal use (at the book saving)
            inline const std::vector<std::vector<size_t> > & GetIndexes() const { return styleIndexes; }

            /// @brief	For internal use (at the book saving)
            inline const std::vector< std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> > & GetPositions() const { return stylePos; }
            // *INDENT-ON*   For AStyle tool

        /// @brief  Adds a new style into collection if it is not exists yet
        /// @param  style Reference to the Style structure object
        /// @return Style index that should be used at data appending to a data sheet
        /// @note   If returned value is 0 - this is a default normal style and it is optional
        ///         whether is will be added into column description or not
        ///         (but better not to add to reduce size and resource consumption)
        size_t Add( const Style & style )
        {
            std::vector<size_t> styleLinks( STYLE_LINK_NUMBER );

            // Check border existance
            bool addItem = true;
            for( size_t i = 0; i < borders.size(); i++ )
            {
                if( borders[i] == style.border )
                {
                    addItem = false;
                    styleLinks[STYLE_LINK_BORDER] = i;
                    break;
                }
            }

            // Add border if it is not in collection yet
            if( addItem )
            {
                borders.push_back( style.border );
                styleLinks[STYLE_LINK_BORDER] = borders.size() - 1;
            }

            // Check font existance
            addItem = true;
            for( size_t i = 0; i < fonts.size(); i++ )
            {
                if( fonts[i] == style.font )
                {
                    addItem = false;
                    styleLinks[STYLE_LINK_FONT] = i;
                    break;
                }
            }

            // Add font if it is not in collection yet
            if( addItem )
            {
                fonts.push_back( style.font );
                styleLinks[STYLE_LINK_FONT] = fonts.size() - 1;
            }

            // Check fill existance
            addItem = true;
            for( size_t i = 0; i < fills.size(); i++ )
            {
                if( fills[i] == style.fill )
                {
                    addItem = false;
                    styleLinks[STYLE_LINK_FILL] = i;
                    break;
                }
            }

            // Add fill if it is not in collection yet
            if( addItem )
            {
                fills.push_back( style.fill );
                styleLinks[STYLE_LINK_FILL] = fills.size() - 1;
            }

            // Check number format existance
            addItem = true;
            for( size_t i = 0; i < nums.size(); i++ )
            {
                if( nums[i] == style.numFormat )
                {
                    addItem = false;
                    styleLinks[STYLE_LINK_NUM_FORMAT] = nums[ i ].id;
                    break;
                }
            }

            // Add number format if it is not in collection yet
            if( addItem )
            {
                if( style.numFormat.id >= BUILT_IN_STYLES_NUMBER )
                {
                    styleLinks[STYLE_LINK_NUM_FORMAT] = fmtLastId;
                    style.numFormat.id = fmtLastId++;
                }
                else
                {
                    styleLinks[STYLE_LINK_NUM_FORMAT] = nums.size();
                }

                nums.push_back( style.numFormat );
            }

            // Check style combination existance
            for( size_t i = 0; i < styleIndexes.size(); i++ )
            {
                if( styleIndexes[i] == styleLinks &&
                        stylePos[i].first.first == style.horizAlign &&
                        stylePos[i].first.second == style.vertAlign &&
                        stylePos[i].second == style.wrapText )
                    return i;
            }

            std::pair<std::pair<EAlignHoriz, EAlignVert>, bool> pos;
            pos.first.first = style.horizAlign;
            pos.first.second = style.vertAlign;
            pos.second = style.wrapText;
            stylePos.push_back( pos );

            styleIndexes.push_back( styleLinks );
            return styleIndexes.size() - 1;
        }
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

}	// namespace SimpleXlsx

#endif // XLSX_SIMPLE_XLSX_DEF_H
