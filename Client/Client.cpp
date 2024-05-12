// Client.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "ObjBase.h"
#include "COM.h"
#include "../Server/include/JSON_Share/I_JSON_Share.h"

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL);
 
	IUnknown* unknown = NULL;
	HRESULT hr_unknown = CoCreateInstance(
		CLSID_JSON_Share,
		NULL, 
		CLSCTX_INPROC_SERVER, 
		IID_IUnknown, 
		(void**)&unknown
	);
	printf("got CoCreateInstance result %d\n", hr_unknown);

	I_JSON_Share * json_share = NULL;
	HRESULT hr_json_share = unknown->QueryInterface(IID_JSON_Share, (void**)&json_share);

	if (SUCCEEDED(hr_json_share)) {
		int value;
		json_share->Get(2, value);
		printf("json_share->Get(2, value) %i \n", value);
	};

	json_share->Release();
	CoFreeUnusedLibraries();
	return 0;
}

