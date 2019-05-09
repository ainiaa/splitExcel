#ifndef QSTRINGUTIL_H
#define QSTRINGUTIL_H

#include <QString>

#ifdef UNICODE
#define QStringToTCHAR(x) (wchar_t *)x.utf16()
#define PQStringToTCHAR(x) (wchar_t *)x->utf16()
#define TCHARToQString(x) QString::fromUtf16((x))
#define TCHARToQStringN(x, y) QString::fromUtf16((x), (y))
#else
#define QStringToTCHAR(x) x.local8Bit().constData()
#define PQStringToTCHAR(x) x->local8Bit().constData()
#define TCHARToQString(x) QString::fromLocal8Bit((x))
#define TCHARToQStringN(x, y) QString::fromLocal8Bit((x), (y))
#endif

class QStringUtil {
    public:
    QStringUtil();
    wchar_t *qstringToWchart(QString str);
    QString WchartToQstring(wchar_t *wchart);
    QString *WchartToQstringP(wchar_t *wchart);
};

#endif // QSTRINGUTIL_H
