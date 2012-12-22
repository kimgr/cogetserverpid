/*******************************************************************************
 Copyright (c) 2012, Kim Gräsman
 All rights reserved.

 Released under the BSD license. For details, please see LICENSE file.

*******************************************************************************/
#ifndef __COGETSERVERPID_H__
#define __COGETSERVERPID_H__

#include <atlbase.h>

HRESULT CoGetServerPID(IUnknown* punk, DWORD* pdwPID)
{
  HRESULT hr;

  if(pdwPID == NULL) return E_POINTER;
  if(punk == NULL) return E_INVALIDARG;

  // This structure represents the OBJREF
  // up to the pid at offset 52 bytes
  typedef struct tagOBJREFHDR
  {
    DWORD  signature;  // Should be 'MEOW'
    BYTE  padding[48];
    USHORT  pid;
  } OBJREFHDR;

  CComPtr<IUnknown> spProxyManager;
  hr = punk->QueryInterface(IID_IProxyManager, (void**)&spProxyManager);
  if(FAILED(hr))
  {
    // No standard proxy, we can't make any assumptions about remote PID
    return hr;
  }

  // Marshal the interface to get a new OBJREF
  CComPtr<IStream> spMarshalStream;
  hr = ::CreateStreamOnHGlobal(NULL, TRUE, &spMarshalStream);
  if(FAILED(hr)) return hr;

  hr = ::CoMarshalInterface(spMarshalStream, IID_IUnknown, punk, MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL);
  if(FAILED(hr)) return hr;
  
  // We just created the stream, so go back to a raw pointer
  HGLOBAL hg = NULL;
  hr = ::GetHGlobalFromStream(spMarshalStream, &hg);
  if(SUCCEEDED(hr))
  {
    // Start out pessimistic
    hr = RPC_E_INVALID_OBJREF;

    OBJREFHDR *pObjRefHdr = (OBJREFHDR*)GlobalLock(hg);
    if(pObjRefHdr != NULL)
    {
      // Verify that the signature is MEOW
      if(pObjRefHdr->signature == 0x574f454d)
      {
        // We got the remote PID
        *pdwPID = pObjRefHdr->pid;
        hr = S_OK;
      }

      GlobalUnlock(hg);
    }
  }

  // Rewind stream and release marshal data to keep refcount in order
  LARGE_INTEGER li = {0};
  spMarshalStream->Seek(li, SEEK_SET, NULL);
  CoReleaseMarshalData(spMarshalStream);
  
  return hr;
}

#endif //__COGETSERVERPID_H__
