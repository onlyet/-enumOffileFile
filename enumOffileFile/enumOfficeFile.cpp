#include "enumOfficeFile.h"
#include <Windows.h>
#include <QFileInfo>
#include <QDebug>

#pragma comment(lib,"Ole32.lib")
#pragma comment(lib,"OleAut32.lib")

EnumOfficeFile::EnumOfficeFile(QObject *parent)
{
}

QStringList EnumOfficeFile::enumOfficeFilePath(const QString& suffix)
{
    //CLSID clsid;
    CoInitialize(NULL);
    HRESULT hr = S_FALSE;
    IRunningObjectTable* pRot = NULL;
    IEnumMoniker* pEnumMoniker = NULL;
    IMoniker* pMoniker = NULL;
    QStringList filePathList;

    hr = GetRunningObjectTable(0, &pRot);
    if (FAILED(hr))
    {
        qDebug() << "GetRunningObjectTable failed";
        return filePathList;
    }
    hr = pRot->EnumRunning(&pEnumMoniker);
    if (FAILED(hr))
    {
        qDebug() << "EnumRunning failed";
        return filePathList;
    }

    hr = pEnumMoniker->Reset();
    if (FAILED(hr))
    {
        qDebug() << "Reset failed";
        return filePathList;
    }

    IBindCtx *pbc;
    CreateBindCtx(0, &pbc);

    while ((hr = pEnumMoniker->Next(1, &pMoniker, NULL)) == S_OK)
    {
        OLECHAR* szDisplayName;
        hr = pMoniker->GetDisplayName(pbc, NULL, &szDisplayName);
        if (FAILED(hr))
        {
            qDebug() << "GetDisplayName failed";
            return filePathList;
        }
        else
        {
            DWORD dwLen = WideCharToMultiByte(CP_ACP, 0, szDisplayName, -1, NULL, 0, NULL, FALSE);
            char* szResult = new char[dwLen];
            WideCharToMultiByte(CP_ACP, 0, szDisplayName, -1, szResult, dwLen, NULL, FALSE);

            QString path = QString::fromLocal8Bit(szResult);
            QFileInfo fileInfo(path);

            qDebug() << fileInfo.absoluteFilePath();

            if (suffix == "doc" && (fileInfo.suffix() == "doc" || fileInfo.suffix() == "docx" || fileInfo.suffix() == "wps"))
                filePathList.append(fileInfo.absoluteFilePath());
            else if (suffix == "ppt" && (fileInfo.suffix() == "ppt" || fileInfo.suffix() == "pptx" || fileInfo.suffix() == "dps"))
                filePathList.append(fileInfo.absoluteFilePath());

            delete[] szResult;
        }
        pMoniker->Release();
    }

    pbc->Release();
    pEnumMoniker->Release();
    pRot->Release();
    CoUninitialize();
    return filePathList;
}
