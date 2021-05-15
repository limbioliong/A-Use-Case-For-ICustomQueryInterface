// TestCPPClient.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
#include <windows.h>

// CreateSafeArrayFromBSTRArray()
// This function will create a SafeArray of unsigned shorts (i.e. 2-byte unsiged integers) 
// using the char elements found inside the first parameter "pCharArray".
//
void CreateSafeArrayFromCharArray
(
    OLECHAR* pCharArray,
    ULONG ulArraySize,
    SAFEARRAY** ppSafeArrayReceiver
)
{
    HRESULT hrRetTemp = S_OK;
    SAFEARRAY* pSAFEARRAYRet = NULL;
    SAFEARRAYBOUND rgsabound[1];
    ULONG ulIndex = 0;
    long lRet = 0;

    // Initialise receiver.
    if (ppSafeArrayReceiver)
    {
        *ppSafeArrayReceiver = NULL;
    }

    if (pCharArray)
    {
        rgsabound[0].lLbound = 0;
        rgsabound[0].cElements = ulArraySize;

        pSAFEARRAYRet = (SAFEARRAY*)SafeArrayCreate
        (
            (VARTYPE)VT_UI2,
            (unsigned int)1,
            (SAFEARRAYBOUND*)rgsabound
        );
    }

    for (ulIndex = 0; ulIndex < ulArraySize; ulIndex++)
    {
        long lIndexVector[1];

        lIndexVector[0] = ulIndex;

        SafeArrayPutElement
        (
            (SAFEARRAY*)pSAFEARRAYRet,
            (long*)lIndexVector,
            (void*)(&(pCharArray[ulIndex]))
        );
    }

    if (pSAFEARRAYRet)
    {
        *ppSafeArrayReceiver = pSAFEARRAYRet;
    }

    return;
}

bool CreateServer(LPOLESTR lpwszServerName, IDispatch** ppIDispatchReceiver)
{
    *ppIDispatchReceiver = NULL;

    CLSID clsid;

    HRESULT hrRet = ::CLSIDFromProgID(lpwszServerName, &clsid);

    hrRet = ::CoCreateInstance
    (
        clsid,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_IDispatch,
        (LPVOID*)ppIDispatchReceiver
    );

    if (SUCCEEDED(hrRet) && (*ppIDispatchReceiver != NULL))
    {
        return true;
    }
    else
    {
        return false;
    }
}

void CallTestMethod(IDispatch* pIDispatch)
{
    OLECHAR* wszMemberTestMethod = (OLECHAR*)L"TestMethod";
    DISPID dispidTestMethod = 0;

    HRESULT hrRet = pIDispatch->GetIDsOfNames
    (
        IID_NULL,
        &wszMemberTestMethod,
        1,
        LOCALE_SYSTEM_DEFAULT,
        &dispidTestMethod
    );
    
    VARIANTARG varCharArray = { 0 };
    VARIANT varResult = { 0 };
    EXCEPINFO exception = { 0 };
    UINT uArgErr = 0;    
    OLECHAR charArray[] = { L'a', L'b', L'c', L'd', L'e' };
    SAFEARRAY* pSafeArray = NULL;

    CreateSafeArrayFromCharArray
    (
        charArray, 5, &pSafeArray
    );    

    VariantInit(&varResult);
    VariantInit(&varCharArray);
    V_VT(&varCharArray) = VT_ARRAY | VT_UI2;
    V_ARRAY(&varCharArray) = pSafeArray;

    DISPPARAMS dp = { 0 };
    dp.cArgs = 1;
    dp.rgvarg = &varCharArray;

    VariantInit(&varResult);

    hrRet = pIDispatch->Invoke
    (
        dispidTestMethod,
        IID_NULL,
        LOCALE_SYSTEM_DEFAULT,
        DISPATCH_METHOD,
        &dp,
        &varResult,
        &exception,
        &uArgErr
    );

    SafeArrayDestroy(pSafeArray);
    pSafeArray = NULL;
}

void DoTest()
{
    ::CoInitialize(NULL);

    IDispatch* pIDispatch = NULL;

    bool bRetTemp = CreateServer((LPOLESTR)L"TestCSServer.TestClass", &pIDispatch);

    if (pIDispatch != NULL)
    {
        CallTestMethod(pIDispatch);
        pIDispatch->Release();
        pIDispatch = NULL;
    }

    ::CoUninitialize();
}

int main()
{
    DoTest();
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
