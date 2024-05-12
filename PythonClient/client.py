

import comtypes.client as cc
import comtypes
CLSID_JSON_Share = "{C0AA564B-56B1-437C-A392-88FE5511322A}"
progID_JSON_Share = "COMServer.JSON_Share.1"

class I_JSON_Share(comtypes.IUnknown):
    _iid_ = comtypes.GUID("{5D98B0B9-00C5-49D8-A737-8591B3163147}")
    _methods_ = [
        comtypes.COMMETHOD(
            [],
            comtypes.HRESULT,
            'Get',
            (
                ['in'],
                comtypes.c_int,
                'a'
            ),
            (
                ['out', 'retval'],
                comtypes.POINTER(comtypes.c_int),
                'b'
            )
        )
    ]


if __name__ == "__main__":
    obj = cc.CreateObject(progID_JSON_Share, comtypes.CLSCTX_INPROC_SERVER, None, I_JSON_Share)
    print(obj.Get(2))
