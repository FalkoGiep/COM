#pragma once
#include <ObjBase.h>

// {5219B44A-0874-449E-8611-B7080DBFA6AB}
static const GUID IID_IBVAA_summer =
	{0x5219b44a, 0x874, 0x449e, {0x86, 0x11, 0xb7, 0x8, 0xd, 0xbf, 0xa6, 0xab}};

interface IBVAA_summer : IUnknown
{
	virtual HRESULT __stdcall Add(
		const double x,
		const double y,
		double &z) = 0;
	virtual HRESULT __stdcall Sub(
		const double x,
		const double y,
		double &z) = 0;
};

// {8A2A00DD-8B8D-4898-B08E-000A6E40A2B5}
static const GUID IID_IBVAA_multiplier =
	{0x8a2a00dd, 0x8b8d, 0x4898, {0xb0, 0x8e, 0x0, 0xa, 0x6e, 0x40, 0xa2, 0xb5}};

interface IBVAA_multiplier : IUnknown
{
	virtual HRESULT __stdcall Mul(
		const double x,
		const double y,
		double &z) = 0;
	virtual HRESULT __stdcall Div(
		const double x,
		const double y,
		double &z) = 0;
};
