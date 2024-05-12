#pragma once
#include <ObjBase.h>

HMODULE Module_JSON_Share;
// {C0AA564B-56B1-437C-A392-88FE5511322A}
GUID CLSID_JSON_Share = { 0xc0aa564b, 0x56b1, 0x437c, { 0xa3, 0x92, 0x88, 0xfe, 0x55, 0x11, 0x32, 0x2a } };
const WCHAR FriendlyName_JSON_Share[] = L"JSON_Share";
const WCHAR VersionIndependentProgID_JSON_Share[] = L"COMServer.JSON_Share";
const WCHAR ProgID_JSON_Share[] = L"COMServer.JSON_Share.1";

STDAPI DllInstall(char *s);
STDAPI DllRegisterServer();
STDAPI DllUnregisterServer();
// This function will register a component in the Registry.
// The component calls this function from its DllRegisterServer function.
HRESULT RegisterServer(
    HMODULE hModule,
    const CLSID &clsid,
    const WCHAR *szFriendlyName,
    const WCHAR *szVerIndProgID,
    const WCHAR *szProgID);

// This function will unregister a component.  Components
// call this function from their DllUnregisterServer function.
HRESULT UnregisterServer(
    const CLSID &clsid,
    const WCHAR *szVerIndProgID,
    const WCHAR *szProgID);

BOOL setKeyAndValue(
    const WCHAR *szKey,
    const WCHAR *szSubkey,
    const WCHAR *szValue);

void CLSIDtochar(
    const CLSID &clsid,
    WCHAR *szCLSID,
    int length);

LONG recursiveDeleteKey(HKEY hKeyParent, const WCHAR* szKeyChild);


