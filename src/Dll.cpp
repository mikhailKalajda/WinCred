//
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved.
//
// Standard dll required functions and class factory implementation.

#include <windows.h>
#include <unknwn.h>
#include "Dll.h"
#include "helpers.h"
#include "easylogging++.h"

INITIALIZE_EASYLOGGINGPP

TIMED_SCOPE(appTimer, "cred");

static long g_cRef = 0;   // global dll reference count
HINSTANCE g_hinst = NULL; // global dll hinstance

extern HRESULT CSample_CreateInstance(__in REFIID riid, __deref_out void** ppv);
EXTERN_C GUID CLSID_CSample;

class CClassFactory : public IClassFactory
{
public:
    CClassFactory() : _cRef(1)
    {
		LOG(INFO) << "CClassFactory::CClassFactory...";
	}

    // IUnknown
    IFACEMETHODIMP QueryInterface(__in REFIID riid, __deref_out void **ppv)
    {
		LOG(INFO) << "CClassFactory::QueryInterface...";

		static const QITAB qit[] =
        {
            QITABENT(CClassFactory, IClassFactory),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
		LOG(INFO) << "CClassFactory::AddRef...";

		return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
		LOG(INFO) << "CClassFactory::Release...";

		long cRef = InterlockedDecrement(&_cRef);
        if (!cRef)
            delete this;
        return cRef;
    }

    // IClassFactory
    IFACEMETHODIMP CreateInstance(__in IUnknown* pUnkOuter, __in REFIID riid, __deref_out void **ppv)
    {
		LOG(INFO) << "CClassFactory::CreateInstance...";

		HRESULT hr;
        if (!pUnkOuter)
        {
            hr = CSample_CreateInstance(riid, ppv);
        }
        else
        {
            *ppv = NULL;
            hr = CLASS_E_NOAGGREGATION;
        }
        return hr;
    }

    IFACEMETHODIMP LockServer(__in BOOL bLock)
    {
		LOG(INFO) << "CClassFactory::LockServer...";

		if (bLock)
        {
            DllAddRef();
        }
        else
        {
            DllRelease();
        }
        return S_OK;
    }

private:
    ~CClassFactory()
    {
    }
    long _cRef;
};

HRESULT CClassFactory_CreateInstance(__in REFCLSID rclsid, __in REFIID riid, __deref_out void **ppv)
{
	LOG(INFO) << "CClassFactory::CClassFactory_CreateInstance...";
	*ppv = NULL;

    HRESULT hr;

    if (CLSID_CSample == rclsid)
    {
        CClassFactory* pcf = new CClassFactory();
        if (pcf)
        {
            hr = pcf->QueryInterface(riid, ppv);
			LOG(INFO) << "pcf->QueryInterface..."  << (int)*ppv;
			pcf->Release();
        }
        else
        {
			LOG(INFO) << "CClassFactory::CClassFactory_CreateInstance - E_OUTOFMEMORY";

			hr = E_OUTOFMEMORY;
        }
    }
    else
    {
		LOG(INFO) << "CClassFactory::CClassFactory_CreateInstance - CLASS_E_CLASSNOTAVAILABLE: " << GuidToString(rclsid) << ", excepted: " << GuidToString(CLSID_CSample);

		hr = CLASS_E_CLASSNOTAVAILABLE;
    }
    return hr;
}

void DllAddRef()
{
	LOG(INFO) << "DllAddRef...";
	InterlockedIncrement(&g_cRef);
}

void DllRelease()
{
	LOG(INFO) << "DllRelease...";
	InterlockedDecrement(&g_cRef);
}

STDAPI DllCanUnloadNow()
{
	LOG(INFO) << "DllCanUnloadNow...";
	return (g_cRef > 0) ? S_FALSE : S_OK;
}

STDAPI DllGetClassObject(__in REFCLSID rclsid, __in REFIID riid, __deref_out void** ppv)
{
	LOG(INFO) << "DllGetClassObject...";
	return CClassFactory_CreateInstance(rclsid, riid, ppv);
}

STDAPI_(BOOL) DllMain(__in HINSTANCE hinstDll, __in DWORD dwReason, __in void *)
{
	LOG(INFO) << "Starting...";

	switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        DisableThreadLibraryCalls(hinstDll);
        break;
    case DLL_PROCESS_DETACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
        break;
    }

    g_hinst = hinstDll;
    return TRUE;
}

