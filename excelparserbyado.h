#ifndef EXCELPARSERBYADO_H
#define EXCELPARSERBYADO_H

#include <QAxObject>
#include <QDir>
#include <QObject>
class ExcelParserByADO : public QObject {
    Q_OBJECT
    public:
    ExcelParserByADO(QObject *parent = nullptr);

    void parse();

    signals:

    public slots:
};

#endif // EXCELPARSERBYADO_H
