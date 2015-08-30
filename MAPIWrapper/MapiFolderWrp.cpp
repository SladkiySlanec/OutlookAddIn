// MapiFolderWrp.cpp : Implementation of CMapiFolderWrp

#include "stdafx.h"
#include "MapiFolderWrp.h"
#include "MapiTableWrp.h"
#include ".\errors\checks.h"
#include ".\errors\BaseException.h"
#include "MapiSessionWrp.h"
#include "Utils\StringHelper.hpp"
#include "Utils\arithmetic.hpp"

// CMapiFolderWrp

STDMETHODIMP CMapiFolderWrp::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* const arr[] = 
	{
		&IID_IMapiFolderWrp
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

// 1. hr Error handler
// 2. ComError creator

void OutPut(LPCWSTR szStr)
{
	time_t rawtime;
	time (&rawtime);
	char szBuff[100];
	ctime_s (szBuff, 100, &rawtime);
	CComBSTR szTime(szBuff);
	OutputDebugString(szTime);
	OutputDebugString(L"	");
	OutputDebugString(szStr);
	OutputDebugString(L"\n");
}
STDMETHODIMP CMapiFolderWrp::GetSubFolders(VARIANT mapiObject,  VARIANT_BOOL goDeep, IMapiTableWrp ** ppSubFolders)
{
	HRESULT hr = E_FAIL;
	try
	{
		MAPIINIT_0 mi = { MAPI_INIT_VERSION, MAPI_NT_SERVICE };
		hr = MAPIInitialize(&mi);
		check_mapi_err(hr);
		CComPtr<IMAPIFolder> ptrMapiFolder;

		SmartSPropValue props;

		hr = mapiObject.punkVal->QueryInterface( IID_IMAPIFolder, (LPVOID*)&ptrMapiFolder);
		hr = HrGetOneProp((IMAPIFolder*)ptrMapiFolder, PR_DISPLAY_NAME, &props); 
		check_mapi_err(hr);

		CComPtr<IMAPITable> pTable;
		ULONG flags = MAPI_UNICODE; 
		if(goDeep)
		{
			flags|=CONVENIENT_DEPTH;
		}
		hr = ptrMapiFolder->GetHierarchyTable(flags, &pTable);
		check_mapi_err(hr);

		CComObject<CMapiTableWrp> * pSubFolders = NULL;
		CComObject<CMapiTableWrp>::CreateInstance(&pSubFolders);
		pSubFolders->Init(pTable);
		CComPtr<IMapiTableWrp> ptrSubFolders(pSubFolders);
		*ppSubFolders = ptrSubFolders.Detach();
	
	}
	catch(BaseException ex)
	{
	
	}
	CATCH_SEH

	return hr ;

	////THROW_MAPI_ERR(pRoot, hr);
	//hr = pTable->SeekRow(BOOKMARK_BEGINNING,0,NULL);
	////THROW_MAPI_ERR(pTable, hr);

	//SizedSPropTagArray(2, sptCols) = { 2, { PR_ENTRYID, PR_DISPLAY_NAME } };
 //   hr = pTable->SetColumns((LPSPropTagArray)&sptCols , 0);
	////THROW_MAPI_ERR(pTable, hr);
	//OutPut(L"Enumareation");
	//int count = 0;
	//while(true)
	//{
	//	LPSRowSet pRows= NULL;
	//	hr = pTable->QueryRows(255, 0, &pRows);
	//	//THROW_MAPI_ERR(pTable, hr);
	//	if(pRows->cRows == 0)
	//		break;
	//	for(ULONG i = 0; i < pRows->cRows; i++)
	//	{
	//		for(int j = 0; j<2; j++)
	//		{
	//			if(pRows->aRow[i].lpProps[j].ulPropTag == PR_DISPLAY_NAME)
	//			{
	//				//LPSPropValue pEid;
	//				//pEid.Alloc(sizeof(SPropValue) );
	//				count++;
	//				OutPut(pRows->aRow[i].lpProps[j].Value.lpszW);


	//				//hr = ::PropCopyMore(pEid, pRows->aRow[i].lpProps, MAPIAllocateMore, pEid);
	//				//THROW_MAPI_ERR0(hr);
	//				
	//				//lstEids.push_back(pEid);
	//			}
	//		}

	//	}
	//	if(pRows != NULL)
	//		MAPIFreeBuffer(pRows);
	//}



	//return S_OK;
}


STDMETHODIMP CMapiFolderWrp::MoveTo(VARIANT mapiSession)
{
	HRESULT hr = E_FAIL;
	try
	{
		CMapiSessionWrp::GlobalInit(mapiSession);		
		LPMAPISESSION pSession = CMapiSessionWrp::GetSession();

		SBinary srcFodlerID = StringHelper::HexStringToSBinary(L"000000001A447390AA6611CD9BC800AA002FC45A0300523F116A08351D4F8B2B79F8C8742119000471E9F7B40000");
		SmartPtr<BYTE> srcFodlerIDPtr(srcFodlerID.lpb);
		
		SBinary newParentFodlerID = StringHelper::HexStringToSBinary(L"000000001A447390AA6611CD9BC800AA002FC45A0300523F116A08351D4F8B2B79F8C8742119000473B68A1B0000");
		SmartPtr<BYTE> newParentFodlerIDPtr(newParentFodlerID.lpb);

		ULONG type;
		CComPtr<IMAPIFolder> ptrSrcFodler;
		hr = pSession->OpenEntry(srcFodlerID.cb,(LPENTRYID)srcFodlerID.lpb, NULL, MAPI_BEST_ACCESS, &type,(LPUNKNOWN* ) &ptrSrcFodler);
		check_mapi_err(hr);

		CComPtr<IMAPIFolder> ptrTrgFodler;
		hr = pSession->OpenEntry(newParentFodlerID.cb,(LPENTRYID)newParentFodlerID.lpb, NULL, MAPI_BEST_ACCESS, &type, (LPUNKNOWN* )&ptrTrgFodler);
		check_mapi_err(hr);

		Move(ptrSrcFodler, ptrTrgFodler);
		//
	}
	catch(BaseException ex)
	{
	
	}
	CATCH_SEH

	return hr ;
}

HRESULT CMapiFolderWrp::Move(LPMAPIFOLDER pSrcFodler, LPMAPIFOLDER pTrgFodler)
{
	HRESULT hr = E_FAIL;
	SmartSPropValue ptrSrcFolderProps;
	SizedSPropTagArray(5, tags) = { 5, { PR_CHANGE_KEY, PR_PREDECESSOR_CHANGE_LIST, PR_PARENT_SOURCE_KEY,  
										PR_SOURCE_KEY,	PR_DISPLAY_NAME,} };
	ULONG ulCount;
	hr = pSrcFodler->GetProps((LPSPropTagArray)&tags, 0, &ulCount, &ptrSrcFolderProps);
	check_mapi_err(hr);
	check_arg(ulCount == 5);

	check_arg(ptrSrcFolderProps[0].ulPropTag == PR_CHANGE_KEY);
	check_arg(ptrSrcFolderProps[1].ulPropTag == PR_PREDECESSOR_CHANGE_LIST);
	check_arg(ptrSrcFolderProps[2].ulPropTag == PR_PARENT_SOURCE_KEY);

	SmartPtr<WCHAR> ptrChangeKey = StringHelper::SBinaryToHexString(&ptrSrcFolderProps[0].Value.bin);

	SmartPtr<WCHAR> ptrChangeList = StringHelper::SBinaryToHexString(&ptrSrcFolderProps[1].Value.bin);
	std::wstring szChangeList(ptrChangeList);
	std::wstring szChangeKey(ptrChangeKey);

	std::size_t pos = szChangeList.rfind(szChangeKey);

	//Increment PR_CHANGE_KEY and PR_CHANGE_KEY in PR_PREDECESSOR_CHANGE_LIST
	Arithmetic::IncrementLast4Bytes(&ptrSrcFolderProps[1].Value.bin.lpb[pos],szChangeKey.length());
	Arithmetic::IncrementLast4Bytes(ptrSrcFolderProps[0].Value.bin.lpb, ptrSrcFolderProps[0].Value.bin.cb);

	//change a parent ID on folder 
	SmartSPropValue ptrDestParentSK;
    hr = ::HrGetOneProp(pTrgFodler, PR_SOURCE_KEY, &ptrDestParentSK);
	check_mapi_err(hr);
	SBinary& currentParentSK = ptrSrcFolderProps[2].Value.bin;
    SBinary tempBuffSK = currentParentSK;
	currentParentSK = ptrDestParentSK->Value.bin;

	 CComPtr<IExchangeImportHierarchyChanges> ptrDestCollector;
     hr = pTrgFodler->OpenProperty(PR_COLLECTOR, &IID_IExchangeImportHierarchyChanges,
            NULL, MAPI_DEFERRED_ERRORS, (LPUNKNOWN*)(&ptrDestCollector));
	 check_mapi_err(hr);

    //apply move
	 hr = ptrDestCollector->ImportFolderChange(ulCount, ptrSrcFolderProps.m_pMapiItem);
    
	ptrSrcFolderProps[2].Value.bin= tempBuffSK;

	return hr;
}