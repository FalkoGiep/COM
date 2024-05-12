#include "stdafx.h"
#include <assert.h>
#include "Windows.h"
#include "Registry.h"




const int CLSID_STRING_SIZE = 39; // Size of GUID with braces
const int CUSTOM_MAX_PATH = MAX_PATH;
const WCHAR* key_prefix = L"CLSID\\";
const int CLSID_WITH_PREFIX_STRING_SIZE = sizeof(key_prefix) + 2 + CLSID_STRING_SIZE;

STDAPI DllRegisterServer()
{
	HRESULT rc = RegisterServer(
		Module_JSON_Share,
		CLSID_JSON_Share,
		FriendlyName_JSON_Share,
		VersionIndependentProgID_JSON_Share,
		ProgID_JSON_Share);
	return rc;
};

STDAPI DllUnregisterServer()
{
	HRESULT rc = UnregisterServer(
		CLSID_JSON_Share,
		VersionIndependentProgID_JSON_Share,
		ProgID_JSON_Share);

	return rc;
};

STDAPI DllInstall(char *s)
{
	return S_OK;
};

HRESULT RegisterServer(
	HMODULE hModule,
	const CLSID &clsid,
	const WCHAR *FriendlyName,
	const WCHAR *VersionIndependentProgID,
	const WCHAR *ProgID)
{
	// get the file path of the relevant module
	WCHAR ModuleFileName[CUSTOM_MAX_PATH];
	DWORD result = GetModuleFileName(hModule, ModuleFileName, CUSTOM_MAX_PATH);
	assert(result != 0);

	WCHAR szCLSID[CLSID_STRING_SIZE];
	CLSIDtochar(clsid, szCLSID, CLSID_STRING_SIZE);

	
	WCHAR szKey[CLSID_WITH_PREFIX_STRING_SIZE];
	wcscpy_s(szKey, key_prefix);
	wcscat_s(szKey, szCLSID);

	// set the CLSID key (https://learn.microsoft.com/en-us/windows/win32/com/clsid-key-hklm)
	setKeyAndValue(szKey, NULL, FriendlyName);
	setKeyAndValue(szKey, L"InprocServer32", ModuleFileName);
	setKeyAndValue(szKey, L"ProgID", ProgID);
	setKeyAndValue(szKey, L"VersionIndependentProgID", VersionIndependentProgID);
	// set the VersionIndependentProgID Key(https://learn.microsoft.com/en-us/windows/win32/com/-version-independent-progid--key)
	setKeyAndValue(VersionIndependentProgID, NULL, FriendlyName);
	setKeyAndValue(VersionIndependentProgID, L"CLSID", szCLSID);
	setKeyAndValue(VersionIndependentProgID, L"CurVer", ProgID);
	// set the ProgID key ((https://learn.microsoft.com/en-us/windows/win32/com/-progid--key)
	setKeyAndValue(ProgID, NULL, FriendlyName);
	setKeyAndValue(ProgID, L"CLSID", szCLSID);

	return S_OK;
}

HRESULT UnregisterServer(
	const CLSID &clsid,
	const WCHAR *szVerIndProgID,
	const WCHAR *szProgID)
{
	WCHAR szCLSID[CLSID_STRING_SIZE];
	CLSIDtochar(clsid, szCLSID, sizeof(szCLSID));

	WCHAR szKey[CLSID_WITH_PREFIX_STRING_SIZE];
	wcscpy_s(szKey, key_prefix);
	wcscat_s(szKey, szCLSID);

	LONG lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szKey);
	assert((lResult == ERROR_SUCCESS) || (lResult == ERROR_FILE_NOT_FOUND));
	lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szVerIndProgID);
	assert((lResult == ERROR_SUCCESS) || (lResult == ERROR_FILE_NOT_FOUND));
	lResult = recursiveDeleteKey(HKEY_CLASSES_ROOT, szProgID);
	assert((lResult == ERROR_SUCCESS) || (lResult == ERROR_FILE_NOT_FOUND));
	return S_OK;
}

void CLSIDtochar(
	const CLSID &clsid,
	WCHAR *szCLSID,
	int length)
{
	assert(length >= CLSID_STRING_SIZE);
	LPOLESTR wszCLSID = NULL;
	HRESULT hr = StringFromCLSID(clsid, &wszCLSID);
	assert(SUCCEEDED(hr));
	wcscpy_s(szCLSID, length, wszCLSID);
	CoTaskMemFree(wszCLSID);
}

LONG recursiveDeleteKey(
	HKEY hKeyParent,
	const WCHAR *lpszKeyChild)
{	
	HKEY hKeyChild;
	LONG lRes = RegOpenKeyEx(
		hKeyParent,
		lpszKeyChild,
		0,
		KEY_ALL_ACCESS,
		&hKeyChild);
	if (lRes != ERROR_SUCCESS)
	{
		return lRes;
	}
	FILETIME time;
	WCHAR szBuffer[256];
	DWORD dwSize = 256;
	while (
		RegEnumKeyEx(
			hKeyChild,
			0,
			szBuffer,
			&dwSize,
			NULL,
			NULL,
			NULL,
			&time) == S_OK)
	{
		lRes = recursiveDeleteKey(hKeyChild, szBuffer);
		if (lRes != ERROR_SUCCESS)
		{
			RegCloseKey(hKeyChild);
			return lRes;
		}
		dwSize = 256;
	}

	RegCloseKey(hKeyChild);
	return RegDeleteKey(hKeyParent, lpszKeyChild);
}

BOOL setKeyAndValue(
	const WCHAR *szKey,
	const WCHAR *szSubkey,
	const WCHAR *szValue)
{
	HKEY hKey;
	WCHAR szKeyBuf[1024];

	wcscpy_s(szKeyBuf, szKey);

	if (szSubkey != NULL)
	{
		wcscat_s(szKeyBuf, L"\\");
		wcscat_s(szKeyBuf, szSubkey);
	}
	long lResult = RegCreateKeyEx(
		HKEY_CLASSES_ROOT,
		szKeyBuf,
		0, 
		NULL, 
		REG_OPTION_NON_VOLATILE,
		KEY_ALL_ACCESS, 
		NULL,
		&hKey, 
		NULL
	);
	if (lResult != ERROR_SUCCESS)
	{
		return FALSE;
	}

	if (szValue != NULL)
	{
		RegSetValueEx(
			hKey,
			NULL,
			0,
			REG_SZ,
			(BYTE *)szValue,
			2 * wcslen(szValue) + 1
		);
	}
	RegCloseKey(hKey);
	return TRUE;
}
