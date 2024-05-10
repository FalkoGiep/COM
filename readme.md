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
