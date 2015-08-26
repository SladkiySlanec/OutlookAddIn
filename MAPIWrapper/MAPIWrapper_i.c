

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.00.0595 */
/* at Sun Aug 23 06:33:04 2015
 */
/* Compiler settings for MAPIWrapper.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0595 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IMapiFolderWrp,0x92BCCF25,0x803C,0x4312,0x89,0x59,0xAA,0x58,0x66,0x12,0xA2,0x18);


MIDL_DEFINE_GUID(IID, IID_IMapiTableWrp,0x2F5010CC,0xB8D4,0x4330,0x89,0xDB,0xD3,0x6A,0xC4,0xF1,0x23,0x7A);


MIDL_DEFINE_GUID(IID, LIBID_MAPIWrapperLib,0x24043FC4,0x3E9C,0x48BD,0x95,0x63,0x3C,0x76,0x96,0x0C,0x20,0x73);


MIDL_DEFINE_GUID(IID, DIID__IMapiFolderWrpEvents,0xD88741BD,0x32E9,0x4747,0x86,0x64,0x87,0x21,0x8C,0x30,0x2B,0x87);


MIDL_DEFINE_GUID(CLSID, CLSID_MapiFolderWrp,0x02AE975F,0x6795,0x463E,0xBE,0xD8,0x6F,0xFC,0x1B,0x74,0xD9,0x58);


MIDL_DEFINE_GUID(CLSID, CLSID_MapiTable,0x8CD1E521,0xF1E1,0x49A2,0xA3,0x1B,0x89,0xE4,0x10,0xB7,0xBD,0xBA);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



