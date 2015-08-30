// MapiFolderWrp.h : Declaration of the CMapiFolderWrp

#pragma once
#include "resource.h"       // main symbols



#include "MAPIWrapper_i.h"
#include "_IMapiFolderWrpEvents_CP.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CMapiFolderWrp

class ATL_NO_VTABLE CMapiFolderWrp :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMapiFolderWrp, &CLSID_MapiFolderWrp>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CMapiFolderWrp>,
	public CProxy_IMapiFolderWrpEvents<CMapiFolderWrp>,
	public IDispatchImpl<IMapiFolderWrp, &IID_IMapiFolderWrp, &LIBID_MAPIWrapperLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CMapiFolderWrp()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_MAPIFOLDERWRP)

DECLARE_NOT_AGGREGATABLE(CMapiFolderWrp)

BEGIN_COM_MAP(CMapiFolderWrp)
	COM_INTERFACE_ENTRY(IMapiFolderWrp)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CMapiFolderWrp)
	CONNECTION_POINT_ENTRY(__uuidof(_IMapiFolderWrpEvents))
END_CONNECTION_POINT_MAP()
// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	STDMETHOD(GetSubFolders)(VARIANT mapiObject, VARIANT_BOOL goDeep, IMapiTableWrp ** ppSubFolders);
	STDMETHOD(MoveTo)(VARIANT mapiSession);
private:
	HRESULT Move(LPMAPIFOLDER pSrcFodler, LPMAPIFOLDER pTrgFodler);

};

OBJECT_ENTRY_AUTO(__uuidof(MapiFolderWrp), CMapiFolderWrp)
