// MapiSubFolders.h : Declaration of the CMapiSubFolders

#pragma once
#include "resource.h"       // main symbols

#include "SmartPointers\MapiSmartPtr.h"

#include "MAPIWrapper_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CMapiSubFolders

class ATL_NO_VTABLE CMapiSubFolders :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CMapiSubFolders, &CLSID_MapiSubFolders>,
	public IDispatchImpl<IMapiSubFolders, &IID_IMapiSubFolders, &LIBID_MAPIWrapperLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CMapiSubFolders()
	{
		m_isSetSeekRowAndColumns = false;
		m_counter = 0;
		m_pRows= NULL;
	}

	void Init(IMAPITable * pTable);
DECLARE_REGISTRY_RESOURCEID(IDR_MAPISUBFOLDERS)


BEGIN_COM_MAP(CMapiSubFolders)
	COM_INTERFACE_ENTRY(IMapiSubFolders)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

private:
	void SetSeekRowAndColumns();
private:
	CComPtr<IMAPITable> m_ptrTable;
	bool m_isSetSeekRowAndColumns;
	int m_counter;
	SmartSRowSet m_pRows;
	
public:
	STDMETHOD (GetNextItemPro)(VARIANT ppsaMyArray, ULONG* size, VARIANT *buffer);
	STDMETHOD(GetNextItem)(BSTR* DisplayName);
	STDMETHOD(get_Count)(ULONG* pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(MapiSubFolders), CMapiSubFolders)
