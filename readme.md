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

## Source

- [COM in C (codeproject)](https://www.codeproject.com/Articles/338268/COM-in-C)

## Resources

- [Comprehensible notes on Aggregation in COM](https://www.codeproject.com/Articles/5337/COM-Concept-Unleashing-Aggregation)
- [Explaination of the ClassFactory concept](https://www.codeproject.com/Articles/4685/Single-class-object-for-multiple-COM-classes)
