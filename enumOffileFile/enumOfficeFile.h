#pragma once

#include <QObject>
#include "EnumOfficeFilePluginInterface.h"


class Q_DECL_EXPORT EnumOfficeFile : public EnumOfficeFilePluginInterface
{
    Q_OBJECT
    Q_PLUGIN_METADATA(IID EnumOfficeFilePluginInterface_iid FILE "EnumOfficeFile.json")
    Q_INTERFACES(EnumOfficeFilePluginInterface)

public:
    explicit EnumOfficeFile(QObject *parent = nullptr);
    virtual ~EnumOfficeFile() {}

    QStringList enumOfficeFilePath(const QString& suffix) override;

};


