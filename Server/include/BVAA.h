#pragma once
#include "IBVAA.h"

class BVAA : public IBVAA_summer, public IBVAA_multiplier
{
protected:
	// Reference count
	long m_lRef;

public:
	BVAA(void);
	~BVAA(void);

	// IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
	virtual ULONG STDMETHODCALLTYPE AddRef(void);
	virtual ULONG STDMETHODCALLTYPE Release(void);

	// IBVAA_summer
	virtual HRESULT STDMETHODCALLTYPE Add(const double x, const double y, double &z);
	virtual HRESULT STDMETHODCALLTYPE Sub(const double x, const double y, double &z);

	// IBVAA_multiplier
	virtual HRESULT STDMETHODCALLTYPE Mul(const double x, const double y, double &z);
	virtual HRESULT STDMETHODCALLTYPE Div(const double x, const double y, double &z);
};
