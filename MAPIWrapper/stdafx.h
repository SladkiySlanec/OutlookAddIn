// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"
#define USES_IID_IMAPISession 
#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

//STD
#include <string>
#include <vector>
#include <list>
#include <set>
#include <map>
#include <memory>
#include <sstream>
#include <iomanip>

////WIN
//#include <windows.h>	// to remove shellapi32.dll from import table
//#include <Ntsecapi.h>
//#include <Wincrypt.h>
//#include <assert.h>
//
#include "resource.h"
#include <atlbase.h>
//
#include <atlctl.h>

#include "ToDoList.hpp"

//com
#include <atlcom.h>

//MAPI
#include <InitGuid.h>
#define USES_IID_IMAPIFolder
#include <MAPIX.h>
#include <Mapidefs.h>
#include <Mapiutil.h> 
#include <MAPIGuid.h> 
#include "EdkMdb.h"
#include "EdkGuid.h"

#include "SmartPointers\MapiSmartPtr.h"
#include "SmartPointers\SmartPtr.hpp"
#include "Errors\EventLogReporter.h"
#include "Errors\errors.hpp"
#include "Errors\ComException.h"
