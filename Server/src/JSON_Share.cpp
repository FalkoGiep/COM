#include "StdAfx.h"
#include "JSON_Share/JSON_Share.h"


extern ULONG g_Components;

JSON_Share::JSON_Share(void) : referenceCount(1)
{
	InterlockedIncrement((LONG*)&g_Components);
}

JSON_Share::~JSON_Share(void)
{
	InterlockedDecrement((LONG*)&g_Components);
}

#pragma region IUknown

HRESULT STDMETHODCALLTYPE JSON_Share::QueryInterface(REFIID riid, void ** ppv)
{
	HRESULT rc = S_OK;
	*ppv = NULL;
	if (riid == IID_IUnknown)
	{
		*ppv = (I_JSON_Share*)this;
	}
	else if (riid == IID_JSON_Share)
	{
		*ppv = (I_JSON_Share*)this;
	}
	else
	{
		rc = E_NOINTERFACE;
	}
	if (rc == S_OK)
		this->AddRef();
	return rc;
}

ULONG STDMETHODCALLTYPE JSON_Share::AddRef(void)
{
	InterlockedIncrement((LONG*)&(this->referenceCount));
	return this->referenceCount;
}

ULONG STDMETHODCALLTYPE JSON_Share::Release(void)
{
	InterlockedDecrement((LONG*)&(this->referenceCount));
	if (referenceCount == 0)
	{
		delete this;
		return 0;
	}
	else
		return referenceCount;
}

#pragma endregion

#pragma region I_JSON_Share 

HRESULT STDMETHODCALLTYPE JSON_Share::Get(
	const int ii,
	int& jj
) {
	jj = 2 * ii;
	return S_OK;
}

#pragma endregion