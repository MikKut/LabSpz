#include <iostream>
#include <Windows.h>
#include <WbemCli.h>
#include <vector>
#include <Psapi.h>
#include <stdio.h>
#include <sstream>
#include <comutil.h>
#include <tuple>
#include <tchar.h>
#include <wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "kernel32.lib")
using namespace std;

HANDLE hConsole;
int GetProcessorInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService);
int GetAllProcessorInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService);
int GetFiveProcessesWithMostThreads(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService);
int GetMSWordProcessInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService);
void PrintSuccess(const char* text);
void PrintFail(const char* text, HRESULT res);
HRESULT StopLowPriorityNotepadProcess(IWbemServices* pSvc);
HRESULT StopTotalCommanderChildProcess(IWbemServices* pSvc);
HRESULT Task05_01(IWbemServices* pSvc);
HRESULT Task05_02(IWbemServices* pSvc);
HRESULT Task05(IWbemServices* pSvc);

const wchar_t* ZhenyaPathToWord = L"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.exe";
const wchar_t* MishaPathToWord = ZhenyaPathToWord;

int main()
{    
    hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
    //First
    HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hRes)) {
        cout << "Unable to launch COM: 0x" << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("Con library has been successfully initialized");
    }

    if ((FAILED(hRes = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_CONNECT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, 0))))
    {
        cout << "Unable to initialize security: 0x" << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("Security layers has been successfully initialized");
    }

    //Second
    IWbemLocator* pLocator = NULL;
    if (FAILED(hRes = CoCreateInstance(CLSID_WbemLocator, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pLocator)))) {
        cout << "Unable to create a WbemLocator: " << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("WbemLocator has been successfully created");
    }

    //Third
    IWbemServices* pService = NULL;
    if (FAILED(hRes = pLocator->ConnectServer(BSTR(L"root\\CIMV2"), NULL, NULL, NULL, WBEM_FLAG_CONNECT_USE_MAX_WAIT, NULL, NULL, &pService))) {
        pLocator->Release();
        cout << "Unable to connect to \"CIMV2\": " << std::hex << hRes << endl;
        return 1;
    }
    else {
        PrintSuccess("Connection to server has been successfully created");
    }
    /*
    //Fourth

    // Task 1
    SetConsoleTextAttribute(hConsole, 13);
    cout << endl << "The First task: " << endl << endl;
    SetConsoleTextAttribute(hConsole, 7);

    ShowFullInfoAboutProcessor(hRes, pLocator, pService);

    // Task 2
    SetConsoleTextAttribute(hConsole, 13);
    cout << endl << "THe Second task: " << endl << endl;
    SetConsoleTextAttribute(hConsole, 7);

    ShowDescriptionAndNumberOfFunctionKeysOfKeyboard(hRes, pLocator, pService);

    // Task 3
    SetConsoleTextAttribute(hConsole, 13);
    cout << endl << "The Third task" << endl << endl;
    SetConsoleTextAttribute(hConsole, 7);

    ShowInfoAboutRunningProcess(hRes, pLocator, pService);

    // Task 4
    SetConsoleTextAttribute(hConsole, 13);
    cout << endl << "The Fourth task" << endl << endl;
    SetConsoleTextAttribute(hConsole, 7);

    GetInfoAboutProcByReadingSize();

    // Fifth
   

    system("pause");
    return 0;
    */
    GetProcessorInfo(hRes, pLocator, pService);
    GetAllProcessorInfo(hRes, pLocator, pService);
    GetMSWordProcessInfo(hRes, pLocator, pService);
    GetFiveProcessesWithMostThreads(hRes, pLocator, pService);
    Task05(pService);


    pService->Release();
    pLocator->Release();
    CoUninitialize();
    return 0;
}

void PrintFail(const char* text, HRESULT res) {
    SetConsoleTextAttribute(hConsole, 12);
    cout << text << std::hex << res << endl;
    SetConsoleTextAttribute(hConsole, 7);
}

void PrintSuccess(const char* text) {
    SetConsoleTextAttribute(hConsole, 10);
    cout << text << endl;
    SetConsoleTextAttribute(hConsole, 7);
}
//Task 1
int GetProcessorInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    cout << endl << endl << "Task 1: "<<endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT Manufacturer, PowerManagementSupported, Name FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* pclsObj;

    /*ULONG uReturn = 0;
    while (pEnumerator)
    {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
            &pclsObj, &uReturn);

        if (uReturn == 0)
        {
            break;
        }
        SAFEARRAY* sfArray;
        LONG lstart, lend;
        VARIANT vtProp;
        pclsObj->GetNames(0, WBEM_FLAG_ALWAYS, 0, &sfArray);
        hr = SafeArrayGetLBound(sfArray, 1, &lstart);
        if (FAILED(hr)) return hr;
        hr = SafeArrayGetUBound(sfArray, 1, &lend);
        if (FAILED(hr)) return hr;
        BSTR* pbstr;
        hr = SafeArrayAccessData(sfArray, (void HUGEP**) & pbstr);
        int nIdx = 0;
        if (SUCCEEDED(hr))
        {
            CIMTYPE pType;
            for (nIdx = lstart; nIdx <= lend; nIdx++)
            {
                hr = pclsObj->Get(pbstr[nIdx], 0, &vtProp, &pType, 0);
                if (vtProp.vt == VT_NULL)
                {
                    continue;
                }
                if (pType == CIM_STRING && pType != CIM_EMPTY && pType != CIM_ILLEGAL)
                {
                    wcout << "Property value: " << ' ' << " " << vtProp.bstrVal << endl;
                }

                VariantClear(&vtProp);

            }
            hr = SafeArrayUnaccessData(sfArray);
            if (FAILED(hr)) return hr;
        }       



        pclsObj->Release();

        cout << endl;
    }*/
    IWbemClassObject* clsObj = NULL;
    /*if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT PowerManagementSupported FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }*/
    int numElems;
    while ((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems)) != WBEM_S_FALSE)
    {
        if (FAILED(hRes)) {
            break;
        }
        VARIANT vRet;
        VariantInit(&vRet);
        if (SUCCEEDED(clsObj->Get(L"Manufacturer", 0, &vRet, NULL, NULL)))
        {
            std::wcout << L"Manufacturer: " << vRet.bstrVal << endl;
            VariantClear(&vRet);
        }
        if (SUCCEEDED(clsObj->Get(L"Name", 0, &vRet, NULL, NULL)))
        {
            wstring str(vRet.bstrVal, SysStringLen(vRet.bstrVal));
            auto str1 = str.substr(0, str.find(L"@"));
            auto str2 = str.substr(str.find(L"@") +1, str.size() - str.find(L"@"));
            std::wcout << L"Name: " << str1 << endl;
            std::wcout << L"Frequency: " << str2 << endl;
            VariantClear(&vRet);
        }
        if (SUCCEEDED(clsObj->Get(L"PowerManagementSupported", 0, &vRet, NULL, NULL)))
        {
            auto isSupported = L"";
            if (vRet.boolVal) {
                isSupported = L"yes";
            }
            else {
                isSupported = L"no";
            }
            std::wcout << L"PowerManagementSupported: " << isSupported << endl;
            VariantClear(&vRet);
        }

        clsObj->Release();
    }
}
//Task 2
int GetAllProcessorInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    cout << endl << endl << "Task 2: "<<endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }
    IWbemClassObject* pclsObj;
    ULONG uReturn = 0;
   while (pEnumerator)
   {
       HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
           &pclsObj, &uReturn);

       if (uReturn == 0)
       {
           break;
       }
       SAFEARRAY* sfArray;
       LONG lstart, lend;
       VARIANT vtProp;
       pclsObj->GetNames(0, WBEM_FLAG_ALWAYS, 0, &sfArray);
       hr = SafeArrayGetLBound(sfArray, 1, &lstart);
       if (FAILED(hr)) return hr;
       hr = SafeArrayGetUBound(sfArray, 1, &lend);
       if (FAILED(hr)) return hr;
       BSTR* pbstr;
       hr = SafeArrayAccessData(sfArray, (void HUGEP**) & pbstr);
       int nIdx = 0;
       if (SUCCEEDED(hr))
       {
           CIMTYPE pType;
           for (nIdx = lstart; nIdx <= lend; nIdx++)
           {
               hr = pclsObj->Get(pbstr[nIdx], 0, &vtProp, &pType, 0);
               if (vtProp.vt == VT_NULL)
               {
                   continue;
               }
               if (pType == CIM_STRING && pType != CIM_EMPTY && pType != CIM_ILLEGAL)
               {
                   wcout << "Property value: " << ' ' << " " << vtProp.bstrVal << endl;
               }

               VariantClear(&vtProp);

           }
           hr = SafeArrayUnaccessData(sfArray);
           if (FAILED(hr)) return hr;
       }



       pclsObj->Release();

       cout << endl;
   }
    IWbemClassObject* clsObj = NULL;
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT PowerManagementSupported FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }
}
//Task 3
void CreateMsWordProcess()
{
    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    CreateProcess(ZhenyaPathToWord, NULL, NULL, NULL, TRUE
        , IDLE_PRIORITY_CLASS, NULL, NULL, &si, &pi);

}

int ShowInfoAboutThreads(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService, int activeProcessId
    , int numberOfThreads)
{
    IEnumWbemClassObject* pEnumerator = NULL;

    stringstream oss;
    string queryStr = "SELECT * FROM WIN32_THREAD WHERE ProcessHandle=";
    oss << activeProcessId;
    queryStr += oss.str();

    BSTR query = _com_util::ConvertStringToBSTR(queryStr.c_str());
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), query, WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems;
    if (!FAILED(hRes))
    {
        while (numberOfThreads != 0)
        {
            if (FAILED((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems))) == false)
            {
                VARIANT vRet;
                VariantInit(&vRet);
                if (SUCCEEDED(clsObj->Get(L"ProcessHandle", 0, &vRet, NULL, NULL)))
                {
                    std::wcout << L"Id that created process: " << vRet.uintVal << endl;
                    VariantClear(&vRet);
                }
                if (SUCCEEDED(clsObj->Get(L"Priority", 0, &vRet, NULL, NULL)))
                {
                    std::wcout << L"Dynamics priority: " << vRet.uintVal << endl;
                    VariantClear(&vRet);
                }
                if (SUCCEEDED(clsObj->Get(L"PriorityBase", 0, &vRet, NULL, NULL)))
                {
                    std::wcout << L"Base priority: " << vRet.uintVal << endl;
                    VariantClear(&vRet);
                }
                if (SUCCEEDED(clsObj->Get(L"ElapsedTime", 0, &vRet, NULL, NULL)))
                {
                    std::wcout << L"Time spent: " << vRet.uintVal << endl;
                    VariantClear(&vRet);
                }
                if (SUCCEEDED(clsObj->Get(L"ThreadState", 0, &vRet, NULL, NULL)))
                {
                    std::wcout << L"State: " << vRet.uintVal << endl;
                    VariantClear(&vRet);
                }

                cout << endl;
            }

            numberOfThreads--;

        }
        pEnumerator->Release();
        if (clsObj != nullptr) {
            clsObj->Release();
        }
    }
    return 0;
}

int GetMSWordProcessInfo(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService)
{
    CreateMsWordProcess();
    IEnumWbemClassObject* pEnumerator = NULL;
    cout << endl << endl << "Task 3: " << endl;
    // CHANGE !!!
    // CHANGE !!!
    // CHANGE !!!
    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Process WHERE Name = 'WINWORD.EXE'")
        , WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }

    IWbemClassObject* clsObj = NULL;
    int numElems = 0;
    int activeProcessId = 0;
    int numberOfThreads = 0;
    if ((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems)) != WBEM_S_FALSE)
    {
        if (!FAILED(hRes))
        {
            VARIANT vRet;
            VariantInit(&vRet);
            if (SUCCEEDED(clsObj->Get(L"ExecutablePath", 0, &vRet, NULL, NULL)))
            {
                std::wcout << L"Path: " << vRet.bstrVal << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"Name", 0, &vRet, NULL, NULL)))
            {
                std::wcout << L"Name: " << vRet.bstrVal << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"Priority", 0, &vRet, NULL, NULL)))
            {
                std::wcout << L"Priority: " << vRet.uintVal << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"ProcessId", 0, &vRet, NULL, NULL)))
            {
                activeProcessId = vRet.uintVal;
                std::wcout << L"Id: " << activeProcessId << endl;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"ThreadCount", 0, &vRet, NULL, NULL)))
            {
                numberOfThreads = vRet.uintVal;
                std::wcout << L"Thread count: " << numberOfThreads << endl;
                VariantClear(&vRet);
            }
        }

        clsObj->Release();
    }
    pEnumerator->Release();

    cout << endl;

    ShowInfoAboutThreads(hRes, pLocator, pService, activeProcessId, numberOfThreads);

    cout << endl;

    return 0;
}
//Task 4
int GetFiveProcessesWithMostThreads(HRESULT hRes, IWbemLocator* pLocator, IWbemServices* pService) {
    // Get the list of process identifiers.
    cout << "Task 4: " << endl;

    DWORD ProcessIDs[1024], cbNeeded, cProcesses;
    //unsigned int i;

    if (!EnumProcesses(ProcessIDs, sizeof(ProcessIDs), &cbNeeded))
    {
        return 1;
    }

    // Calculate how many process identifiers were returned.

    cProcesses = cbNeeded / sizeof(DWORD);

    // Print the memory usage for each process
    vector<tuple<DWORD, DWORD, BSTR>> processes;
    IEnumWbemClassObject* pEnumerator = NULL;
    IWbemClassObject* clsObj = NULL;
    int numElems = 0;

    if (FAILED(hRes = pService->ExecQuery(BSTR(L"WQL"), BSTR(L"SELECT * FROM Win32_Process")
        , WBEM_FLAG_FORWARD_ONLY, NULL, &pEnumerator))) {
        pLocator->Release();
        pService->Release();
        cout << "Unable to retrive desktop monitors: " << std::hex << hRes << endl;
        return 1;
    }
    while ((hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems)) != WBEM_S_FALSE)
    {
        hRes = pEnumerator->Next(WBEM_INFINITE, 1, &clsObj, (ULONG*)&numElems);
        DWORD id;
        DWORD threadCount;
        BSTR name = BSTR(L"");
        if (!FAILED(hRes))
        {
            VARIANT vRet;
            VariantInit(&vRet);
            if (SUCCEEDED(clsObj->Get(L"Name", 0, &vRet, NULL, NULL)))
            {
                name = vRet.bstrVal;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"ProcessId", 0, &vRet, NULL, NULL)))
            {
                id = vRet.uintVal;
                VariantClear(&vRet);
            }
            if (SUCCEEDED(clsObj->Get(L"ThreadCount", 0, &vRet, NULL, NULL)))
            {
                threadCount = vRet.uintVal;
                VariantClear(&vRet);
            }
        }
        processes.push_back(make_tuple(threadCount, id, name));
        clsObj->Release();
    }

    for (size_t i = 0; i < processes.size(); i++)
    {
        for (size_t j = 0; j < processes.size() - i - 1; j++)
        {
            if (get<0>(processes[j]) < get<0>(processes[j + 1]))
            {
                swap(processes[j], processes[j + 1]);
            }
        }
    }
    cout << "5 processes with the most amount of threads: " << endl;
    for (size_t i = 0; i < 5; i++) {
        cout << endl;
        cout << "Process ID: " << get<1>(processes[i]) << endl;
        cout << "Threads: " << get<0>(processes[i]) << endl;
        wcout << "Name of process: " << get<2>(processes[i]) << endl;
    }

    /*for (i = 0; i < cProcesses; i++)
    {
        DWORD processID = ProcessIDs[i];
        HANDLE hProcess;
        PROCESS_MEMORY_COUNTERS pmc;


        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION |
            PROCESS_VM_READ,
            FALSE, processID);
        if (FAILED(hProcess))
            return 1;

        CloseHandle(hProcess);
    }*/

    return 0;
}
//Task 5a
HRESULT StopLowPriorityNotepadProcess(IWbemServices* pSvc)
{
    HRESULT hr = S_OK;

    static LPCTSTR lpszMethod = _T("Terminate");
    static LPCTSTR lpszClass = _T("Win32_Process");

    IWbemClassObject* pClsInParam = NULL;
    IWbemClassObject* pClsInParamInst = NULL;

    BSTR bszClsMoniker = NULL;
    VARIANT v;

    VariantInit(&v);

    IEnumWbemClassObject* pEnum = NULL;

    hr = pSvc->ExecQuery(
        (BSTR)_T("WQL"),
        (BSTR)_T("SELECT * FROM Win32_Process WHERE Name='notepad.exe'"),
        0,
        NULL, &pEnum
    );

    IWbemClassObject* pObj = NULL;
    IWbemClassObject* pClsDef = NULL;
    if (FAILED(hr))
        goto fail;

    hr = pSvc->GetObject(
        (BSTR)lpszClass, 0,
        NULL, &pClsDef, NULL
    );

    hr = pClsDef->GetMethod(
        lpszMethod, 0,
        &pClsInParam, NULL
    );

    hr = pClsInParam->SpawnInstance(0, &pClsInParamInst);

    V_VT(&v) = VT_UI4;
    V_UI4(&v) = 0;
    pClsInParamInst->Put(_T("Reason"), 0, &v, CIM_UINT32);

    while (1) {
        ULONG uRet = 0;
        pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);

        if (uRet == 0)
            break;

        pObj->Get(
            _T("Handle"), 0,
            &v, 0, 0
        );

        bszClsMoniker = SysAllocString(_T("Win32_Process.Handle='"));
        VarBstrCat(bszClsMoniker, V_BSTR(&v), &bszClsMoniker);
        VarBstrCat(bszClsMoniker, (BSTR)_T("'"), &bszClsMoniker);

        hr = pSvc->ExecMethod(
            bszClsMoniker, (BSTR)lpszMethod, 0,
            NULL, pClsInParamInst, NULL, NULL
        );

        SysFreeString(bszClsMoniker);

        if (FAILED(hr)) {
            _tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
            goto fail;
        }
    }

    PrintSuccess("Task 5a was completed successfully.");
    return hr;

fail:
    PrintFail("Task 5a - something goes wrong", hr);
    VariantClear(&v);
    pClsInParamInst->Release();
    pClsInParam->Release();
    pClsDef->Release();
    pObj->Release();
    pEnum->Release();
}
// Task 5b
HRESULT StopTotalCommanderChildProcess(IWbemServices* pSvc)
{
    HRESULT hr = S_OK;

    IEnumWbemClassObject* pEnum = NULL;
    IWbemClassObject* pClsInParam = NULL;

    IWbemClassObject* pClsInParamInst = NULL;

    static LPCTSTR lpszMethod = _T("Terminate");
    static LPCTSTR lpszClass = _T("Win32_Process");

    BSTR bszWQLQueryChild = NULL;

    VARIANT v;

    VariantInit(&v);

    IWbemClassObject* pClsDef = NULL;

    hr = pSvc->GetObject(
        (BSTR)lpszClass, 0,
        NULL, &pClsDef, NULL
    );

    hr = pClsDef->GetMethod(
        lpszMethod, 0,
        &pClsInParam, NULL
    );

    hr = pClsInParam->SpawnInstance(0, &pClsInParamInst);

    V_VT(&v) = VT_UI4;
    V_UI4(&v) = 0;

    pClsInParamInst->Put(_T("Reason"), 0, &v, CIM_UINT32);

    hr = pSvc->ExecQuery(
        (BSTR)_T("WQL"),
        (BSTR)_T("SELECT * ")
        _T("FROM Win32_Process ")
        _T("WHERE Name='totalcmd.exe' OR Name='totalcmd64.exe'"),
        0,
        NULL, &pEnum
    );

    IWbemClassObject* pObj = NULL;

    if (FAILED(hr))
        goto fail;

    while (1) {
        IEnumWbemClassObject* pEnumChild = NULL;
        IWbemClassObject* pObjChild = NULL;

        ULONG uRet = 0;

        bszWQLQueryChild = SysAllocString(
            _T("SELECT * ")
            _T("FROM Win32_Process ")
            _T("WHERE ParentProcessId=")
        );

        pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);

        if (uRet == 0)
            break;

        pObj->Get(
            _T("Handle"), 0,
            &v, 0, 0
        );

        VarBstrCat(bszWQLQueryChild, V_BSTR(&v), &bszWQLQueryChild);

        hr = pSvc->ExecQuery(
            (BSTR)_T("WQL"),
            bszWQLQueryChild,
            0,
            NULL, &pEnumChild
        );

        if (FAILED(hr))
            goto fail;

        while (1) {
            ULONG uRet = 0;
            pEnumChild->Next(WBEM_INFINITE, 1, &pObjChild, &uRet);

            if (uRet == 0)
                break;

            pObjChild->Get(
                _T("__PATH"), 0,
                &v, 0, 0
            );

            hr = pSvc->ExecMethod(
                V_BSTR(&v), (BSTR)lpszMethod, 0,
                NULL, pClsInParamInst, NULL, NULL
            );

            if (FAILED(hr)) {
                _tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
                goto fail;
            }
        }
        SysFreeString(bszWQLQueryChild);
    }

    PrintSuccess("Task 5b was completed successfully.");
    return hr;

fail:
    PrintFail("Task 5a - something goes wrong", hr);
    SysFreeString(bszWQLQueryChild);
    VariantClear(&v);
    pClsInParamInst->Release();
    pClsInParam->Release();
    pClsDef->Release();
    pObj->Release();
    pEnum->Release();
}

HRESULT Task05_01(IWbemServices* pSvc)
{
    HRESULT hr = S_OK;
    IEnumWbemClassObject* pEnum = NULL;
    IWbemClassObject* pObj = NULL;
    IWbemClassObject* pClsDef = NULL;
    IWbemClassObject* pClsInParam = NULL;
    IWbemClassObject* pClsInParamInst = NULL;
    static LPCTSTR lpszMethod = _T("Terminate");
    static LPCTSTR lpszClass = _T("Win32_Process");
    BSTR bszClsMoniker = NULL;
    VARIANT v;
    VariantInit(&v);

    std::wcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

    hr = pSvc->ExecQuery(
        (BSTR)_T("WQL"),
        (BSTR)_T("SELECT * ")
        _T("FROM Win32_Process ")
        _T("WHERE Name='notepad.exe' AND Priority='4'"),
        0,
        NULL, &pEnum
    );
    if (FAILED(hr))
        goto fail;

    hr = pSvc->GetObject(
        (BSTR)lpszClass, 0,
        NULL, &pClsDef, NULL
    );

    hr = pClsDef->GetMethod(
        lpszMethod, 0,
        &pClsInParam, NULL
    );

    hr = pClsInParam->SpawnInstance(0, &pClsInParamInst);

    V_VT(&v) = VT_UI4;
    V_UI4(&v) = 0;
    pClsInParamInst->Put(_T("Reason"), 0, &v, CIM_UINT32);

    while (1) {
        ULONG uRet = 0;
        pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
        if (uRet == 0)
            break;
        pObj->Get(
            _T("Handle"), 0,
            &v, 0, 0
        );
        bszClsMoniker = SysAllocString(_T("Win32_Process.Handle='"));
        VarBstrCat(bszClsMoniker, V_BSTR(&v), &bszClsMoniker);
        VarBstrCat(bszClsMoniker, (BSTR)_T("'"), &bszClsMoniker);

        hr = pSvc->ExecMethod(
            bszClsMoniker, (BSTR)lpszMethod, 0,
            NULL, pClsInParamInst, NULL, NULL
        );
        SysFreeString(bszClsMoniker);
        if (FAILED(hr)) {
            _tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
            goto fail;
        }
    }

    goto fail;
fail:
    VariantClear(&v);
    if(pClsInParamInst != nullptr) pClsInParamInst->Release();
    if (pClsInParam != nullptr) pClsInParam->Release();
    if (pClsDef != nullptr) pClsDef->Release();
    if (pObj != nullptr) pObj->Release();
    if (pEnum != nullptr) pEnum->Release();
    return hr;
}

static HRESULT Task05_02(IWbemServices* pSvc)
{
    HRESULT hr = S_OK;
    IEnumWbemClassObject* pEnum = NULL;
    IWbemClassObject* pObj = NULL;
    IWbemClassObject* pClsDef = NULL;
    IWbemClassObject* pClsInParam = NULL;
    IWbemClassObject* pClsInParamInst = NULL;
    static LPCTSTR lpszMethod = _T("Terminate");
    static LPCTSTR lpszClass = _T("Win32_Process");
    BSTR bszWQLQueryChild = NULL;
    VARIANT v;
    VariantInit(&v);

    std::wcout << _T("-- ") << _T(__FUNCTION__) << _T("\n");

    hr = pSvc->GetObject(
        (BSTR)lpszClass, 0,
        NULL, &pClsDef, NULL
    );

    hr = pClsDef->GetMethod(
        lpszMethod, 0,
        &pClsInParam, NULL
    );

    hr = pClsInParam->SpawnInstance(0, &pClsInParamInst);

    V_VT(&v) = VT_UI4;
    V_UI4(&v) = 0;
    pClsInParamInst->Put(_T("Reason"), 0, &v, CIM_UINT32);

    hr = pSvc->ExecQuery(
        (BSTR)_T("WQL"),
        (BSTR)_T("SELECT * ")
        _T("FROM Win32_Process ")
        _T("WHERE Name='totalcmd.exe' OR Name='totalcmd64.exe'"),
        0,
        NULL, &pEnum
    );
    if (FAILED(hr))
        goto fail;

    while (1) {
        IEnumWbemClassObject* pEnumChild = NULL;
        IWbemClassObject* pObjChild = NULL;
        ULONG uRet = 0;
        bszWQLQueryChild = SysAllocString(
            _T("SELECT * ")
            _T("FROM Win32_Process ")
            _T("WHERE ParentProcessId=")
        );
        pEnum->Next(WBEM_INFINITE, 1, &pObj, &uRet);
        if (uRet == 0)
            break;
        pObj->Get(
            _T("Handle"), 0,
            &v, 0, 0
        );

        VarBstrCat(bszWQLQueryChild, V_BSTR(&v), &bszWQLQueryChild);

        hr = pSvc->ExecQuery(
            (BSTR)_T("WQL"),
            bszWQLQueryChild,
            0,
            NULL, &pEnumChild
        );
        if (FAILED(hr))
            goto fail;

        while (1) {
            ULONG uRet = 0;
            pEnumChild->Next(WBEM_INFINITE, 1, &pObjChild, &uRet);
            if (uRet == 0)
                break;
            pObjChild->Get(
                _T("__PATH"), 0,
                &v, 0, 0
            );

            hr = pSvc->ExecMethod(
                V_BSTR(&v), (BSTR)lpszMethod, 0,
                NULL, pClsInParamInst, NULL, NULL
            );
            if (FAILED(hr)) {
                _tprintf_s(_T("ExecMethod failed, hr: %lX\n"), hr);
                goto fail;
            }
        }
        SysFreeString(bszWQLQueryChild);
    }

fail:
    SysFreeString(bszWQLQueryChild);
    VariantClear(&v);
    if (pClsInParamInst != nullptr) pClsInParamInst->Release();
    if (pClsInParam != nullptr) pClsInParam->Release();
    if (pClsDef != nullptr) pClsDef->Release();
    if (pObj != nullptr) pObj->Release();
    if (pEnum != nullptr) pEnum->Release();
    return hr;
}

HRESULT Task05(IWbemServices* pSvc)
{
    HRESULT hr = S_OK;
    hr = Task05_01(pSvc);
    hr = Task05_02(pSvc);
    return hr;
}