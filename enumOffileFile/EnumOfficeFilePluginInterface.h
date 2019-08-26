#pragma once

#include <QObject>
#include <QStringList>

class EnumOfficeFilePluginInterface : public QObject
{
public:

    //virtual ~DeviceInterface() {}

    virtual QStringList enumOfficeFilePath(const QString& suffix) = 0;
};

#define EnumOfficeFilePluginInterface_iid "cn.com.gdtech.WisdomClass.EnumOfficeFilePluginInterface/1.0"
Q_DECLARE_INTERFACE(EnumOfficeFilePluginInterface, EnumOfficeFilePluginInterface_iid)