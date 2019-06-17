/*******************************************************************************
Copyright (c) 2012, Kim Gräsman
All rights reserved.

Released under the Modified BSD license. For details, please see LICENSE file.

*******************************************************************************/
#ifndef INCLUDED_COGETSERVERPID_H__
#define INCLUDED_COGETSERVERPID_H__

#include <objbase.h>

/** OBJREF structure that include PID at offset 52 bytes.
    Packed struct to make offsets deterministic.
    REF: https://en.wikipedia.org/wiki/OBJREF */
#pragma pack(push, 1)
struct COGETSERVERPID_OBJREFHDR {
    unsigned long  signature;   // Should be 'MEOW'
    unsigned long  flags;       // 1 (OBJREF_STANDARD)
    GUID           iid;         // IID_IUnknown
    /* Start of STDOBJREF */
    unsigned long  sorfFlags;   // SORF_ flags (see above)
    unsigned long  cPublicRefs; // count of references passed
    unsigned hyper oxid;        // oxid of server with this oid
    unsigned hyper oid;         // oid of object with this ipid
    /* Start of Interface Pointer IDentifier (IPID). A 128-bit number that uniquely
       identifies an interface on an object within an object exporter. */
    unsigned short field1;
    unsigned short field2;
    unsigned short pid;         // process ID (clamped on overflow)
    unsigned short field4;
    unsigned short fields[4];
};
#pragma pack(pop)

inline HRESULT CoGetServerPID(IUnknown* punk, DWORD* pdwPID)
{
  HRESULT hr;
  IUnknown* pProxyManager = NULL;
  IStream* pMarshalStream = NULL;
  HGLOBAL hg = NULL;
  LARGE_INTEGER zero = {0};

  if(pdwPID == NULL) return E_POINTER;
  if(punk == NULL) return E_INVALIDARG;

  /* Make sure this is a standard proxy, otherwise we can't make any
     assumptions about OBJREF wire format. */
  hr = punk->QueryInterface(IID_IProxyManager, (void**)&pProxyManager);
  if(FAILED(hr)) return hr;
  
  pProxyManager->Release();

  /* Marshal the interface to get a new OBJREF. */
  hr = ::CreateStreamOnHGlobal(NULL, TRUE, &pMarshalStream);
  if(SUCCEEDED(hr))
  {
    hr = ::CoMarshalInterface(pMarshalStream, IID_IUnknown, punk, 
      MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL);
    if(SUCCEEDED(hr))
    {
      /* We just created the stream so it's safe to go back to a raw pointer. */
      hr = ::GetHGlobalFromStream(pMarshalStream, &hg);
      if(SUCCEEDED(hr))
      {
        /* Start out pessimistic. */
        hr = RPC_E_INVALID_OBJREF;

        COGETSERVERPID_OBJREFHDR * pObjRefHdr = (COGETSERVERPID_OBJREFHDR*)GlobalLock(hg);
        if(pObjRefHdr != NULL)
        {
          /* Verify type and MEOW signature. */
          if ((pObjRefHdr->signature == 0x574f454d) && (pObjRefHdr->flags == 1))
          {
            /* detect clamped PID */
            if (pObjRefHdr->pid != 0xFFFF)
            {
              /* We got the remote PID! */
              *pdwPID = pObjRefHdr->pid;
              hr = S_OK;
            }
          }

          GlobalUnlock(hg);
        }
      }

      /* Rewind stream and release marshal data to keep refcount in order. */
      pMarshalStream->Seek(zero, SEEK_SET, NULL);
      CoReleaseMarshalData(pMarshalStream);
    }

    pMarshalStream->Release();
  }

  return hr;
}

#endif // INCLUDED_COGETSERVERPID_H__
