#include <initguid.h>
#include <windows.h>
#include <iostream>
#include <string>
#include <comdef.h>

#define _WIN32_MSM 200
#include <mergemod.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

int main(int argc, char* argv[]) {
    if (argc != 3) {
        std::cout << "Usage: " << argv[0] << " <msm_file> <output_dir>\n";
        return 1;
    }

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        std::cout << "Failed to initialize COM\n";
        return 1;
    }

    try {
        IMsmMerge2* merge = nullptr;
        hr = CoCreateInstance(CLSID_MsmMerge2, NULL, CLSCTX_INPROC_SERVER, IID_IMsmMerge2, reinterpret_cast<void**>(&merge));

        if (FAILED(hr))
            throw _com_error(hr);

        _bstr_t msmFile(argv[1]);
        _bstr_t outputDir(argv[2]);

        hr = merge->OpenModule(msmFile, {});
        if (FAILED(hr))
            throw _com_error(hr);

        IMsmStrings* extractedFiles = nullptr;
        hr = merge->ExtractFilesEx(outputDir, VARIANT_TRUE, &extractedFiles);
        if (SUCCEEDED(hr) && extractedFiles) {
            long count = 0;
            extractedFiles->get_Count(&count);
            std::cout << "Extracted " << count << " files:\n";

            for (long i = 0; i < count; i++) {
                BSTR fileName = nullptr;
                if (SUCCEEDED(extractedFiles->get_Item(i, &fileName))) {
                    std::wcout << fileName << std::endl;
                    SysFreeString(fileName);
                }
            }
            extractedFiles->Release();
        }

        if (FAILED(hr))
            throw _com_error(hr);

        merge->CloseModule();
        merge->Release();
        std::cout << "Files extracted successfully to: " << argv[2] << std::endl;
    }
    catch (const _com_error& e) {
        std::wcout << L"Error: " << e.ErrorMessage() << std::endl;
        return 1;
    }

    CoUninitialize();
    return 0;
}