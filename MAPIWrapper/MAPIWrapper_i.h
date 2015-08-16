

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0595 */
/* at Sun Aug 16 12:17:36 2015
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


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __MAPIWrapper_i_h__
#define __MAPIWrapper_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IMapiFolderWrp_FWD_DEFINED__
#define __IMapiFolderWrp_FWD_DEFINED__
typedef interface IMapiFolderWrp IMapiFolderWrp;

#endif 	/* __IMapiFolderWrp_FWD_DEFINED__ */


#ifndef __IMapiSubFolders_FWD_DEFINED__
#define __IMapiSubFolders_FWD_DEFINED__
typedef interface IMapiSubFolders IMapiSubFolders;

#endif 	/* __IMapiSubFolders_FWD_DEFINED__ */


#ifndef ___IMapiFolderWrpEvents_FWD_DEFINED__
#define ___IMapiFolderWrpEvents_FWD_DEFINED__
typedef interface _IMapiFolderWrpEvents _IMapiFolderWrpEvents;

#endif 	/* ___IMapiFolderWrpEvents_FWD_DEFINED__ */


#ifndef __MapiFolderWrp_FWD_DEFINED__
#define __MapiFolderWrp_FWD_DEFINED__

#ifdef __cplusplus
typedef class MapiFolderWrp MapiFolderWrp;
#else
typedef struct MapiFolderWrp MapiFolderWrp;
#endif /* __cplusplus */

#endif 	/* __MapiFolderWrp_FWD_DEFINED__ */


#ifndef __MapiSubFolders_FWD_DEFINED__
#define __MapiSubFolders_FWD_DEFINED__

#ifdef __cplusplus
typedef class MapiSubFolders MapiSubFolders;
#else
typedef struct MapiSubFolders MapiSubFolders;
#endif /* __cplusplus */

#endif 	/* __MapiSubFolders_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MAPIWrapper_0000_0000 */
/* [local] */ 




extern RPC_IF_HANDLE __MIDL_itf_MAPIWrapper_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MAPIWrapper_0000_0000_v0_0_s_ifspec;

#ifndef __IMapiFolderWrp_INTERFACE_DEFINED__
#define __IMapiFolderWrp_INTERFACE_DEFINED__

/* interface IMapiFolderWrp */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IMapiFolderWrp;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("92BCCF25-803C-4312-8959-AA586612A218")
    IMapiFolderWrp : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetAllSubFolders( 
            /* [in] */ VARIANT mapiObject,
            /* [retval][out] */ IMapiSubFolders **ppSubFolders) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE MoveTo( 
            /* [in] */ VARIANT session) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IMapiFolderWrpVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMapiFolderWrp * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMapiFolderWrp * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMapiFolderWrp * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMapiFolderWrp * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMapiFolderWrp * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMapiFolderWrp * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMapiFolderWrp * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetAllSubFolders )( 
            IMapiFolderWrp * This,
            /* [in] */ VARIANT mapiObject,
            /* [retval][out] */ IMapiSubFolders **ppSubFolders);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *MoveTo )( 
            IMapiFolderWrp * This,
            /* [in] */ VARIANT session);
        
        END_INTERFACE
    } IMapiFolderWrpVtbl;

    interface IMapiFolderWrp
    {
        CONST_VTBL struct IMapiFolderWrpVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMapiFolderWrp_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMapiFolderWrp_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMapiFolderWrp_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMapiFolderWrp_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMapiFolderWrp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMapiFolderWrp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMapiFolderWrp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IMapiFolderWrp_GetAllSubFolders(This,mapiObject,ppSubFolders)	\
    ( (This)->lpVtbl -> GetAllSubFolders(This,mapiObject,ppSubFolders) ) 

#define IMapiFolderWrp_MoveTo(This,session)	\
    ( (This)->lpVtbl -> MoveTo(This,session) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMapiFolderWrp_INTERFACE_DEFINED__ */


#ifndef __IMapiSubFolders_INTERFACE_DEFINED__
#define __IMapiSubFolders_INTERFACE_DEFINED__

/* interface IMapiSubFolders */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IMapiSubFolders;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("B27D2DC8-FEF4-4F07-BA6B-9E423BC2657D")
    IMapiSubFolders : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetNextItem( 
            /* [out] */ BSTR *DisplayName) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ ULONG *pVal) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE GetNextItemPro( 
            /* [in] */ VARIANT ppsaMyArray,
            /* [out][in] */ ULONG *size,
            /* [out] */ VARIANT *buffer) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IMapiSubFoldersVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IMapiSubFolders * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IMapiSubFolders * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IMapiSubFolders * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IMapiSubFolders * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IMapiSubFolders * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IMapiSubFolders * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IMapiSubFolders * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetNextItem )( 
            IMapiSubFolders * This,
            /* [out] */ BSTR *DisplayName);
        
        /* [id][propget] */ HRESULT ( STDMETHODCALLTYPE *get_Count )( 
            IMapiSubFolders * This,
            /* [retval][out] */ ULONG *pVal);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *GetNextItemPro )( 
            IMapiSubFolders * This,
            /* [in] */ VARIANT ppsaMyArray,
            /* [out][in] */ ULONG *size,
            /* [out] */ VARIANT *buffer);
        
        END_INTERFACE
    } IMapiSubFoldersVtbl;

    interface IMapiSubFolders
    {
        CONST_VTBL struct IMapiSubFoldersVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IMapiSubFolders_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IMapiSubFolders_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IMapiSubFolders_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IMapiSubFolders_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IMapiSubFolders_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IMapiSubFolders_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IMapiSubFolders_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IMapiSubFolders_GetNextItem(This,DisplayName)	\
    ( (This)->lpVtbl -> GetNextItem(This,DisplayName) ) 

#define IMapiSubFolders_get_Count(This,pVal)	\
    ( (This)->lpVtbl -> get_Count(This,pVal) ) 

#define IMapiSubFolders_GetNextItemPro(This,ppsaMyArray,size,buffer)	\
    ( (This)->lpVtbl -> GetNextItemPro(This,ppsaMyArray,size,buffer) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IMapiSubFolders_INTERFACE_DEFINED__ */



#ifndef __MAPIWrapperLib_LIBRARY_DEFINED__
#define __MAPIWrapperLib_LIBRARY_DEFINED__

/* library MAPIWrapperLib */
/* [version][uuid] */ 

typedef /* [helpstring] */ 
enum MapiPropTags
    {
        DisplayName	= 0x3001001f,
        ContentCount	= 0x36020003,
        AssociatedContentCount	= 0x36170003,
        HasSubfolders	= 0x360a000b,
        ParentSourceKey	= 0x65e10102
    } 	MapiPropTags;


EXTERN_C const IID LIBID_MAPIWrapperLib;

#ifndef ___IMapiFolderWrpEvents_DISPINTERFACE_DEFINED__
#define ___IMapiFolderWrpEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IMapiFolderWrpEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__IMapiFolderWrpEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("D88741BD-32E9-4747-8664-87218C302B87")
    _IMapiFolderWrpEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IMapiFolderWrpEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IMapiFolderWrpEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IMapiFolderWrpEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IMapiFolderWrpEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IMapiFolderWrpEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IMapiFolderWrpEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IMapiFolderWrpEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IMapiFolderWrpEvents * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _IMapiFolderWrpEventsVtbl;

    interface _IMapiFolderWrpEvents
    {
        CONST_VTBL struct _IMapiFolderWrpEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IMapiFolderWrpEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IMapiFolderWrpEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IMapiFolderWrpEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IMapiFolderWrpEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IMapiFolderWrpEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IMapiFolderWrpEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IMapiFolderWrpEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IMapiFolderWrpEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_MapiFolderWrp;

#ifdef __cplusplus

class DECLSPEC_UUID("02AE975F-6795-463E-BED8-6FFC1B74D958")
MapiFolderWrp;
#endif

EXTERN_C const CLSID CLSID_MapiSubFolders;

#ifdef __cplusplus

class DECLSPEC_UUID("8CD1E521-F1E1-49A2-A31B-89E410B7BDBA")
MapiSubFolders;
#endif
#endif /* __MAPIWrapperLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long *, unsigned long            , VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserMarshal(  unsigned long *, unsigned char *, VARIANT * ); 
unsigned char * __RPC_USER  VARIANT_UserUnmarshal(unsigned long *, unsigned char *, VARIANT * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long *, VARIANT * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


