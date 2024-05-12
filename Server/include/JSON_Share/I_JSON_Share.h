#pragma once
#include <ObjBase.h>
// {5D98B0B9-00C5-49D8-A737-8591B3163147}
static const GUID IID_JSON_Share =
{ 0x5d98b0b9, 0xc5, 0x49d8, { 0xa7, 0x37, 0x85, 0x91, 0xb3, 0x16, 0x31, 0x47 } };

interface I_JSON_Share: public IUnknown
{
	virtual HRESULT STDMETHODCALLTYPE Get(
		const int ii,
		int &jj
	) = 0;
};
