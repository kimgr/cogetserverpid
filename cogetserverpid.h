/*******************************************************************************
Copyright (c) 2012, Kim Gräsman
All rights reserved.

Released under the Modified BSD license. For details, please see LICENSE file.

*******************************************************************************/
#pragma once

#include <atlbase.h>

/* This structure represents the OBJREF up to the PID at offset 52 bytes.
   1-byte structure packing to make sure offsets are deterministic. */
#pragma pack(push, 1)
typedef struct tagCOGETSERVERPID_OBJREFHDR
{
  DWORD  signature;  /* Should be 'MEOW'. */
  BYTE  padding[48];
  USHORT  pid;
} COGETSERVERPID_OBJREFHDR;
#pragma pack(pop)

inline HRESULT CoGetServerPID(IUnknown* punk, DWORD* pdwPID)
{
  if(pdwPID == NULL) return E_POINTER;
  if(punk == NULL) return E_INVALIDARG;

  {
      /* Make sure this is a standard proxy, otherwise we can't make any
         assumptions about OBJREF wire format. */
      CComPtr<IUnknown> pProxyManager;
      HRESULT hr = punk->QueryInterface(IID_IProxyManager, (void**)&pProxyManager);
      if(FAILED(hr)) return hr;
  }

  /* Marshal the interface to get a new OBJREF. */
  CComPtr<IStream> pMarshalStream;
  HRESULT hr = ::CreateStreamOnHGlobal(NULL, TRUE, &pMarshalStream);
  if(SUCCEEDED(hr))
  {
    hr = ::CoMarshalInterface(pMarshalStream, IID_IUnknown, punk, 
      MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL);
    if(SUCCEEDED(hr))
    {
      /* We just created the stream so it's safe to go back to a raw pointer. */
      HGLOBAL hg = NULL;
      hr = ::GetHGlobalFromStream(pMarshalStream, &hg);
      if(SUCCEEDED(hr))
      {
        /* Start out pessimistic. */
        hr = RPC_E_INVALID_OBJREF;

        COGETSERVERPID_OBJREFHDR *pObjRefHdr = (COGETSERVERPID_OBJREFHDR*)GlobalLock(hg);
        if(pObjRefHdr != NULL)
        {
          /* Verify that the signature is MEOW. */
          if(pObjRefHdr->signature == 0x574f454d)
          {
            /* We got the remote PID! */
            *pdwPID = pObjRefHdr->pid;
            hr = S_OK;
          }

          GlobalUnlock(hg);
        }
      }

      /* Rewind stream and release marshal data to keep refcount in order. */
      LARGE_INTEGER zero = {0};
      pMarshalStream->Seek(zero, SEEK_SET, NULL);
      CoReleaseMarshalData(pMarshalStream);
    }
  }

  return hr;
}
