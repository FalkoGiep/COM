[![Coverity Scan Build Status](https://scan.coverity.com/projects/30682/badge.svg)](https://scan.coverity.com/projects/falkogiep-com)
# Documenation

Test project to experiment with COM technology based on code from codeproject.

## Usage

### c++ server/client

Build with VS on win32. Register the resulting server.dll with regsvr32.

`regsvr32 server.dll`

### unregister dll

`regsvr32 /U server.dll`

### python client

Create a virtual environment of a 32-bit python version with comtypes installed.

## Simplified flow

- [CoCreateInstance](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance) is called using
  - a CLSID (for example: CLSID_BVAA)
  - a IID (for example: IID_IBVAA_summer)
- _[The DllGetClassObject function is called by the COM library as part of the work done by the CoCreateInstance() API](https://www.codeproject.com/Articles/901/Introduction-to-COM-Part-II-Behind-the-Scenes-of-a)_
- DllGetClassObject handles the request conform to the documentation of [windows](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-dllgetclassobject)
- Via the ClassFactory template an object is created and the pointer to the requested interface is returned

## Source

- [COM in C (codeproject)](https://www.codeproject.com/Articles/338268/COM-in-C)

## Resources

- [Comprehensible notes on Aggregation in COM](https://www.codeproject.com/Articles/5337/COM-Concept-Unleashing-Aggregation)
- [Explaination of the ClassFactory concept](https://www.codeproject.com/Articles/4685/Single-class-object-for-multiple-COM-classes)
- [COM server explained 2](https://www.codeproject.com/Articles/901/Introduction-to-COM-Part-II-Behind-the-Scenes-of-a)
- [Explaining the differences between most concepts and their relations](https://docwiki.embarcadero.com/RADStudio/Athens/en/Parts_of_a_COM_Application)

## Registry information

- [Microsoft: COM Registry Keys](https://learn.microsoft.com/en-us/windows/win32/com/com-registry-keys)
- [HKEY_CLASSES_ROOT = HKEY_LOCAL_MACHINE\SOFTWARE\Classes](https://learn.microsoft.com/en-us/windows/win32/sysinfo/hkey-classes-root-key)

### CLSID Key

[CLSID key](https://learn.microsoft.com/en-us/windows/win32/com/clsid-key-hklm)

- `Computer\HKEY_CLASSES_ROOT\WOW6432Node\CLSID\ (CLSID of COM object)`
  - Friendly name
  - InprocServer32: dll location
  - ProgID
  - VersionIndependentProgID

### ProgID Key

[ProgID Key](https://learn.microsoft.com/en-us/windows/win32/com/-progid--key)

- `Computer\HKEY_CLASSES_ROOT\ (ProgID)`
  - CLSID

### VersionIndependentProgID

[VersionIndependentProgID Key](https://learn.microsoft.com/en-us/windows/win32/com/-version-independent-progid--key)

- `Computer\HKEY_CLASSES_ROOT\ (VersionIndependentProgID)`
  - CurVer
