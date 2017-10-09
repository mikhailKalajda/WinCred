#pragma once


#include "helpers.h"
#include <windows.h>
#include <strsafe.h>
#include <new>

#include "CVistaCredential.h"

#define MAX_CREDENTIALS 3
#define MAX_DWORD   0xffffffff        // maximum DWORD

class CVistaProvider : public ICredentialProvider
{
public:
    // IUnknown
    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return ++_cRef;
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long cRef = --_cRef;
        if (!cRef)
        {
            delete this;
        }
        return cRef;
    }

    IFACEMETHODIMP QueryInterface(_In_ REFIID riid, _COM_Outptr_ void **ppv)
    {
        static const QITAB qit[] =
        {
            QITABENT(CVistaProvider, ICredentialProvider), // IID_ICredentialProvider
            { 0 },
        };

        HRESULT hr = QISearch(this, qit, riid, ppv);

        return hr;
    }

public:
    IFACEMETHODIMP SetUsageScenario(CREDENTIAL_PROVIDER_USAGE_SCENARIO cpus, DWORD dwFlags);
    IFACEMETHODIMP SetSerialization(const CREDENTIAL_PROVIDER_CREDENTIAL_SERIALIZATION* pcpcs);

    IFACEMETHODIMP Advise(__in ICredentialProviderEvents* pcpe, UINT_PTR upAdviseContext);
    IFACEMETHODIMP UnAdvise();

    IFACEMETHODIMP GetFieldDescriptorCount(__out DWORD* pdwCount);
    IFACEMETHODIMP GetFieldDescriptorAt(DWORD dwIndex, __deref_out CREDENTIAL_PROVIDER_FIELD_DESCRIPTOR** ppcpfd);

    IFACEMETHODIMP GetCredentialCount(__out DWORD* pdwCount,
        __out DWORD* pdwDefault,
        __out BOOL* pbAutoLogonWithDefault);
    IFACEMETHODIMP GetCredentialAt(DWORD dwIndex,
        __out ICredentialProviderCredential** ppcpc);

    friend HRESULT CSampleProvider_CreateInstance(REFIID riid, __deref_out void** ppv);

public:
    CVistaProvider();
    __override ~CVistaProvider();

private:

    HRESULT _EnumerateOneCredential(__in DWORD dwCredientialIndex,
        __in PCWSTR pwzUsername
    );
    HRESULT _EnumerateSetSerialization();

    // Create/free enumerated credentials.
    HRESULT _EnumerateCredentials();
    void _ReleaseEnumeratedCredentials();
    void _CleanupSetSerialization();

private:
    LONG              _cRef;
    CVistaCredential *_rgpCredentials[MAX_CREDENTIALS];  // Pointers to the credentials which will be enumerated by 
                                                          // this Provider.
    DWORD                                   _dwNumCreds;
    KERB_INTERACTIVE_UNLOCK_LOGON*          _pkiulSetSerialization;
    DWORD                                   _dwSetSerializationCred; //index into rgpCredentials for the SetSerializationCred
    bool                                    _bAutoSubmitSetSerializationCred;
    CREDENTIAL_PROVIDER_USAGE_SCENARIO      _cpus;
};
