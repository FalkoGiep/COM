

import comtypes.client as cc
import comtypes
import win32com.client as win32
import pythoncom
CLSID_BVAA = "{6942E971-6F95-44BC-B3A9-EFD270EB39C9}"
CLSID_BVAB = "{01ADBE7E-4EE3-4F36-9DCA-95722F897AE1}"

class ISummer(comtypes.IUnknown):
    _iid_ = comtypes.GUID("{5219B44A-0874-449E-8611-B7080DBFA6AB}")
    _methods_ = [
        comtypes.COMMETHOD([], comtypes.HRESULT, 'Add',
                           (['in'], comtypes.c_double, 'a'),
                           (['in'], comtypes.c_double, 'b'),
                           (['out', 'retval'], comtypes.POINTER(comtypes.c_double), 'c')),
        comtypes.COMMETHOD([], comtypes.HRESULT, 'Sub',
                            (['in'], comtypes.c_double, 'a'),
                            (['in'], comtypes.c_double, 'b'),
                            (['out', 'retval'], comtypes.POINTER(comtypes.c_double), 'c')),
    ]

progID_BVAA = "BVAA.Component.1"
progID_BVAB = "BVAB.Component.2"

if __name__ == "__main__":
    obj = cc.CreateObject(progID_BVAA, comtypes.CLSCTX_INPROC_SERVER, None, ISummer)
    print(obj.Add(1, 2))
    print(obj.Sub(1, 2))
