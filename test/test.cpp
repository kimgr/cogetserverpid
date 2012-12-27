// test.cpp : Defines the entry point for the console application.
//

#include "cogetserverpid.h"
#include <atlbase.h>

int main(int argc, char* argv[])
{
  CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

  {
    CComPtr<IUnknown> word;
    HRESULT hr = word.CoCreateInstance(L"Word.Application", NULL, CLSCTX_LOCAL_SERVER);
    printf("Word.Application created: 0x%X\n", hr);

    DWORD wordPID = 0;
    hr = CoGetServerPID(word, &wordPID);
    printf("Word.Application PID: %d (0x%X)\n", wordPID, hr);
  }

  CoUninitialize();

	return 0;
}

