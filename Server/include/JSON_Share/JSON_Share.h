#pragma once
#include "I_JSON_Share.h"

class JSON_Share: public I_JSON_Share
{
protected:
	long referenceCount;

public:
	JSON_Share(void);
	~JSON_Share(void);

	// IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv);
	virtual ULONG STDMETHODCALLTYPE AddRef(void);
	virtual ULONG STDMETHODCALLTYPE Release(void);

	// I_JSON_Share
	virtual HRESULT STDMETHODCALLTYPE Get(
		const int ii,
		int& jj
	);
	
};
