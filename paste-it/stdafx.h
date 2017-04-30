// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define NOMINMAX
#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>

// C RunTime Header Files
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>

#define _HAS_CXX17 1
#include <ctime>
#include <codecvt>
#include <filesystem>
#include <fstream>
#include <memory>
#include <optional>

#include <fmt/format.h>
#include <fmt/time.h>

#include <comutil.h>
#include <atlbase.h>
#include <shlobj.h>
#include <exdisp.h>