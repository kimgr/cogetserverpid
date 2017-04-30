unit uCoGetServerPID;

// Copyright (c) 2012, Jon Robertson
// All rights reserved.
//
// Released under the Modified BSD license. For details, please see LICENSE file.
//
// CoGetServerPID will take a COM Object and return the Process ID of the
// server that is hosting the object.
//
// This code was converted from Kim Gräsman's C version.  Please see Kim's web page for additional credits.
// The web page is at http://www.kontrollbehov.com/articles/cogetserverpid/
//
// Thanks to Robert Marquardt for assisting with the conversion.  Thanks to
// Project JEDI (http://www.delphi-jedi.org/) for all of their conversion work!
//
// I use this like so:
//
// var
//   WordApp: OleVariant;
//   WordPID: DWORD;
// begin
//   WordApp := CreateOleObject('Word.Application');
//   OleCheck(CoGetServerPID(FWordApp, dwWordPID));
//
// Although it's a hack, it works great!
// Disclaimer:  Actually, I've only tested it from Delphi 6.
//              Your mileage will vary.

interface

uses Windows;

function CoGetServerPID (const unk: IUnknown; var dwPID: DWORD): HRESULT;

implementation

uses ActiveX, ComObj, Variants;

type
  // Used to verify the COM object's signature and to sniff the PID out of the header.
  TComObjRefHdr = packed record
    signature: DWORD;  	// Should be 'MEOW'
    padding  : array[0..47] of Byte;
    pid      : WORD;
  end;
  PComObjRefHdr = ^TComObjRefHdr;

function CoGetServerPID(const unk: IUnknown; var dwPID: DWORD): HRESULT;
var
  PObjRefHdr       : PComObjRefHdr;
  spProxyManager   : IUnknown;
  spMarshalStream  : IStream;
  liNewPos         : Int64;
  HG               : HGLOBAL;

const
  IID_IUnknown: TGUID = (D1:$00000000;D2:$0000;D3:$0000;D4:($C0,$00,$00,$00,$00,$00,$00,$46));

begin

  if VarIsNULL(unk) then begin
    Result := E_INVALIDARG;
    Exit;
  end;

  // Check if not standard proxy.  We can't make any assumptions about remote PID
  Result := unk.QueryInterface(IID_IProxyManager, spProxyManager);
  if Failed(Result) then Exit;

  // Marshall the interface to get a new OBJREF
  Result := CreateStreamOnHGlobal(0, True, spMarshalStream);
  if Failed(Result) then Exit;

  Result := CoMarshalInterface(spMarshalStream, IID_IUnknown, unk, MSHCTX_INPROC, nil, MSHLFLAGS_NORMAL);
  if Failed(Result) then Exit;

  // We just created the stream.  So go back to a raw pointer.
  Result := GetHGlobalFromStream(spMarshalStream, HG);

  if Succeeded(Result) then try

    // Start out pessimistic
    Result := RPC_E_INVALID_OBJREF;

    PObjRefHdr := GlobalLock(HG);

    if Assigned(PObjRefHdr) then begin
      // Verify that the signature is MEOW
      if PObjRefHdr^.signature = $574f454d then begin
        // We got the remote PID
        dwPID  := PObjRefHdr^.pid;
        Result := S_OK;
      end;
    end;

  finally
    GlobalUnlock(HG);
  end;

  // Rewind stream and release marshal data to keep refcount in order
  liNewPos := 0;
  spMarshalStream.Seek(0, STREAM_SEEK_SET, liNewPos);
  CoReleaseMarshalData(spMarshalStream);

end;

end.

