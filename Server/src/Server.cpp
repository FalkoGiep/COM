
#include "stdafx.h"
#include "ClassFactory.h"
#include "JSON_Share/JSON_Share.h"
//#include "Registry.h"

extern GUID CLSID_JSON_Share;
ULONG g_ServerLocks = 0;
ULONG g_Components = 0;


__control_entrypoint(DllExport)
	STDAPI DllCanUnloadNow(void)
{
	HRESULT rc = E_UNEXPECTED;
	if ((g_ServerLocks == 0) && (g_Components == 0))
		rc = S_OK;
	else
		rc = S_FALSE;
	return rc;
};

_Check_return_
	STDAPI
	DllGetClassObject(_In_ REFCLSID clsid, _In_ REFIID iid, _Outptr_ LPVOID FAR *ppv)
{
	HRESULT rc = E_UNEXPECTED;

	if (clsid == CLSID_JSON_Share)
	{
		ClassFactory<JSON_Share> *cf = new ClassFactory<JSON_Share>();
		rc = cf->QueryInterface(iid, ppv);
		cf->Release();
	}
	else
		rc = CLASS_E_CLASSNOTAVAILABLE;
	return rc;
};