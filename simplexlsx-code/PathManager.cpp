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

#include <algorithm>
#include <cerrno>
#include <cstdio>
#include <cstring>
#include <fstream>
#include <iostream>

#include <stdint.h>

#include "PathManager.hpp"

#ifdef _WIN32
#include <windows.h>
#include <direct.h>
#else
#include <unistd.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <sys/syscall.h>
#endif

namespace SimpleXlsx
{

    //Creating all necessary subdirectories and copy image file
    bool PathManager::RegisterImage( const std::string & LocalPath, const std::string & XLSX_Path )
    {
        std::ifstream Source( LocalPath.c_str(), std::ios::binary );
        if( Source.is_open() )
        {
            const std::string InternalPath = RegisterFile( XLSX_Path );
            std::ofstream Destination( InternalPath.c_str(), std::ios::binary );
            if( Destination.is_open() )
            {
                Destination << Source.rdbuf();
                return true;
            }
        }
        return false;
    }

    //Deletes all temporary files and directories which have been created
    void PathManager::ClearTemp()
    {
        for( std::vector< std::string >::const_iterator it = m_contentFiles.begin(); it != m_contentFiles.end(); it++ )
            remove( ( m_temp_path + ( * it ) ).c_str() );
        m_contentFiles.clear();
        for( std::vector< std::string >::const_reverse_iterator it = m_temp_dirs.rbegin(); it != m_temp_dirs.rend(); it++ )
#ifdef _WIN32
            _rmdir( ( * it ).c_str() );
#else
            rmdir( ( *it ).c_str() );
#endif
        m_temp_dirs.clear();
    }

    // ****************************************************************************
    /// @brief  Function to create nested directories` tree
    /// @param  dirName directories` tree to be created
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool PathManager::MakeDirectory( const std::string & dirName )
    {
        std::string part = "";
        size_t oldPointer = 0;
        size_t currPointer = 0;
        int32_t res = -1;

        for( size_t i = 0; i < dirName.length(); i++ )
        {
            if( ( dirName.at( i ) == '\\' ) || ( dirName.at( i ) == '/' ) )
            {
                part += dirName.substr( oldPointer, currPointer - oldPointer ) + "/";
                oldPointer = currPointer + 1;
#ifdef _WIN32
                std::replace( part.begin(), part.end(), '/', '\\' );
                res = _mkdir( part.c_str() );
#else
                res = mkdir( part.c_str(), 0777 );
#endif
                if( res == 0 ) m_temp_dirs.push_back( part );  //Remember the created subdirectories
                if( ( res == -1 ) && ( errno == ENOENT ) ) return false;
            }
            currPointer++;
        }
        return true;
    }

#if ! defined( _WIN32 )     //Linux with Unicode
    std::string PathManager::PathEncode( const wchar_t * Path )
    {
        mbstate_t MbState;
        const wchar_t * Ptr = Path;
        //mbrlen( NULL, 0, & MbState );
        std::memset( & MbState, 0, sizeof( MbState ) );
        std::size_t BufSize = wcsrtombs( NULL, & Ptr, 0, & MbState );
        if( ( BufSize == static_cast<std::size_t>( -1 ) ) || ( BufSize == 0 ) ) return "";
        std::vector<char> MbVector( BufSize + 1, 0 );
        wcsrtombs( MbVector.data(), & Ptr, BufSize, & MbState );
        MbVector[ BufSize ] = '\0';
        return MbVector.data();
    }

#else                       //Windows with Unicode
    std::string PathManager::PathEncode( const wchar_t * Path )
    {
        int Length = static_cast< int >( wcslen( Path ) );
        int ResultSize = WideCharToMultiByte( CP_ACP, 0, Path, Length, NULL, 0, NULL, NULL );
        std::vector<char> AnsiFileName( ResultSize + 1, 0 );
        WideCharToMultiByte( CP_ACP, 0, Path, Length, AnsiFileName.data(), ResultSize, NULL, NULL );
        return AnsiFileName.data();
    }
#endif

}
