#pragma once
#include <ObjBase.h>

// {59F4A881-5464-4409-9052-8C0D05828EFA}
static const GUID IID_IBVAB_power =
	{0x59f4a881, 0x5464, 0x4409, {0x90, 0x52, 0x8c, 0xd, 0x5, 0x82, 0x8e, 0xfa}};

interface IBVAB_power : IUnknown
{
	virtual HRESULT __stdcall Pow(
		const double x,
		const double y,
		double &z) = 0;
	virtual HRESULT __stdcall Log(
		const double x,
		const double y,
		double &z) = 0;
};