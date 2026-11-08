/*
 * BuyBox.exe — native launcher for the portable Windows build.
 *
 * Layout next to the exe:   python\python.exe   app\launch.py
 * Starts the embedded Python with launch.py (which picks a free port, opens the browser and
 * minimizes this window). The child runs inside a job object, so closing this window — or the
 * launcher dying — always stops the app; nothing is left running in the background.
 *
 * Build (any OS):  python -m ziglang cc -target x86_64-windows-gnu -O2 -municode -s \
 *                    launcher.c buybox.rc -o BuyBox.exe
 */
#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#include <windows.h>
#include <stdio.h>
#include <wchar.h>

#define APP_TITLE L"Buy Box Analyzer"

static void fail(const wchar_t *msg) {
    MessageBoxW(NULL, msg, APP_TITLE, MB_ICONERROR | MB_OK);
}

int wmain(int argc, wchar_t **argv) {
    (void)argc; (void)argv;
    wchar_t dir[32768];
    DWORD n = GetModuleFileNameW(NULL, dir, (DWORD)(sizeof dir / sizeof dir[0]));
    if (n == 0 || n >= sizeof dir / sizeof dir[0]) {
        fail(L"Cannot determine the program folder.");
        return 1;
    }
    wchar_t *slash = wcsrchr(dir, L'\\');
    if (slash) *slash = L'\0';

    wchar_t py[33000], script[33000], cmd[66000];
    swprintf(py, sizeof py / sizeof py[0], L"%ls\\python\\python.exe", dir);
    swprintf(script, sizeof script / sizeof script[0], L"%ls\\app\\launch.py", dir);
    if (GetFileAttributesW(py) == INVALID_FILE_ATTRIBUTES || GetFileAttributesW(script) == INVALID_FILE_ATTRIBUTES) {
        fail(L"فایل‌های برنامه پیدا نشد.\n"
             L"لطفاً کل فایل zip را Extract کنید و BuyBox.exe را از داخل پوشه استخراج‌شده اجرا کنید.\n\n"
             L"Program files not found. Extract the WHOLE zip first, then run BuyBox.exe from the extracted folder.");
        return 1;
    }

    SetConsoleOutputCP(CP_UTF8);
    SetConsoleTitleW(L"Buy Box Analyzer — تحلیل بای باکس (برای بستن برنامه این پنجره را ببندید)");
    SetEnvironmentVariableW(L"PYTHONUTF8", L"1");
    SetEnvironmentVariableW(L"PYTHONIOENCODING", L"utf-8");
    SetEnvironmentVariableW(L"PYTHONNOUSERSITE", L"1");
    SetEnvironmentVariableW(L"PYTHONHOME", NULL);
    SetEnvironmentVariableW(L"PYTHONPATH", NULL);
    SetEnvironmentVariableW(L"BUYBOX_LAUNCHER", L"exe");

    swprintf(cmd, sizeof cmd / sizeof cmd[0], L"\"%ls\" \"%ls\"", py, script);

    /* Kill the app together with this launcher (window closed, task ended, crash). */
    HANDLE job = CreateJobObjectW(NULL, NULL);
    if (job) {
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION info;
        ZeroMemory(&info, sizeof info);
        info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
        SetInformationJobObject(job, JobObjectExtendedLimitInformation, &info, sizeof info);
    }

    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof si);
    si.cb = sizeof si;
    ZeroMemory(&pi, sizeof pi);
    if (!CreateProcessW(py, cmd, NULL, NULL, FALSE, CREATE_SUSPENDED, NULL, dir, &si, &pi)) {
        wchar_t msg[512];
        swprintf(msg, 512, L"Could not start Python (error %lu).", GetLastError());
        fail(msg);
        return 1;
    }
    if (job) AssignProcessToJobObject(job, pi.hProcess);
    ResumeThread(pi.hThread);
    CloseHandle(pi.hThread);

    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD code = 0;
    GetExitCodeProcess(pi.hProcess, &code);
    CloseHandle(pi.hProcess);
    if (code != 0) {
        HWND w = GetConsoleWindow();
        if (w) ShowWindow(w, SW_RESTORE);
        wprintf(L"\n[!] The app stopped with an error (code %lu). See the messages above.\n"
                L"Press Enter to close this window...", code);
        getwchar();
    }
    if (job) CloseHandle(job);
    return (int)code;
}
