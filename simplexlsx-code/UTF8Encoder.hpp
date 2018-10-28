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

#ifndef UTF8ENCODER_H
#define UTF8ENCODER_H

#include <assert.h>
#include <stdint.h>
#include <string>

class UTF8Encoder
{
    public:
        static inline std::string From_wchar_t( int32_t Code )
        {
            assert( ( Code >= 0 ) && "Code can not be negative." );
            if( Code <= 0x7F ) return std::string( 1, static_cast<char>( Code ) );
            else
            {
                unsigned char FirstByte = 0xC0;          //First byte with mask
                unsigned char Buffer[ 7 ] = { 0 };       //Buffer for UTF-8 bytes, null-ponter string
                char * Ptr = reinterpret_cast<char *>( & Buffer[ 5 ] ); //Ptr to Buffer, started from end
                unsigned char LastValue = 0x1F;          //Max value for implement to the first byte
                while( true )
                {
                    * Ptr = static_cast<char>( ( Code & 0x3F ) | 0x80 );    //Making new value with format 10xxxxxx
                    Ptr--;                                  //Move Ptr to new position
                    Code >>= 6;                             //Shifting Code
                    if( Code <= LastValue ) break;          //if Code can set to FirstByte => break;
                    LastValue >>= 1;                        //Make less max value
                    FirstByte = ( FirstByte >> 1 ) | 0x80;  //Modify first byte
                }
                * Ptr = static_cast<char>( FirstByte | Code );  //Making first byte of UTF-8 sequence
                return std::string( Ptr );
            }
        }

        static inline std::string From_wstring( const std::wstring & Source )
        {
            std::string Result;
            for( std::wstring::const_iterator iter = Source.begin(); iter != Source.end(); iter++ )
                Result.append( From_wchar_t( * iter ) );
            return Result;
        }
};


#endif // UTF8ENCODER_H
