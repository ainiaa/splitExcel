/****************************************************************************
** Meta object code from reading C++ file 'emailonlywindow.h'
**
** Created by: The Qt Meta Object Compiler version 67 (Qt 5.14.2)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include <memory>
#include "../emailonlywindow.h"
#include <QtCore/qbytearray.h>
#include <QtCore/qmetatype.h>
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'emailonlywindow.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 67
#error "This file was generated using the moc from 5.14.2. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
QT_WARNING_PUSH
QT_WARNING_DISABLE_DEPRECATED
struct qt_meta_stringdata_EmailOnlyWindow_t {
    QByteArrayData data[14];
    char stringdata0[241];
};
#define QT_MOC_LITERAL(idx, ofs, len) \
    Q_STATIC_BYTE_ARRAY_DATA_HEADER_INITIALIZER_WITH_OFFSET(len, \
    qptrdiff(offsetof(qt_meta_stringdata_EmailOnlyWindow_t, stringdata0) + ofs \
        - idx * sizeof(QByteArrayData)) \
    )
static const qt_meta_stringdata_EmailOnlyWindow_t qt_meta_stringdata_EmailOnlyWindow = {
    {
QT_MOC_LITERAL(0, 0, 15), // "EmailOnlyWindow"
QT_MOC_LITERAL(1, 16, 9), // "doProcess"
QT_MOC_LITERAL(2, 26, 0), // ""
QT_MOC_LITERAL(3, 27, 6), // "doSend"
QT_MOC_LITERAL(4, 34, 31), // "on_selectFilePushButton_clicked"
QT_MOC_LITERAL(5, 66, 13), // "changeGroupby"
QT_MOC_LITERAL(6, 80, 17), // "selectedSheetName"
QT_MOC_LITERAL(7, 98, 29), // "on_savePathPushButton_clicked"
QT_MOC_LITERAL(8, 128, 36), // "on_emailOnlySubmitPushButton_..."
QT_MOC_LITERAL(9, 165, 14), // "receiveMessage"
QT_MOC_LITERAL(10, 180, 7), // "msgType"
QT_MOC_LITERAL(11, 188, 6), // "result"
QT_MOC_LITERAL(12, 195, 27), // "on_gobackPushButton_clicked"
QT_MOC_LITERAL(13, 223, 17) // "showConfigSetting"

    },
    "EmailOnlyWindow\0doProcess\0\0doSend\0"
    "on_selectFilePushButton_clicked\0"
    "changeGroupby\0selectedSheetName\0"
    "on_savePathPushButton_clicked\0"
    "on_emailOnlySubmitPushButton_clicked\0"
    "receiveMessage\0msgType\0result\0"
    "on_gobackPushButton_clicked\0"
    "showConfigSetting"
};
#undef QT_MOC_LITERAL

static const uint qt_meta_data_EmailOnlyWindow[] = {

 // content:
       8,       // revision
       0,       // classname
       0,    0, // classinfo
       9,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       2,       // signalCount

 // signals: name, argc, parameters, tag, flags
       1,    0,   59,    2, 0x06 /* Public */,
       3,    0,   60,    2, 0x06 /* Public */,

 // slots: name, argc, parameters, tag, flags
       4,    0,   61,    2, 0x08 /* Private */,
       5,    1,   62,    2, 0x08 /* Private */,
       7,    0,   65,    2, 0x08 /* Private */,
       8,    0,   66,    2, 0x08 /* Private */,
       9,    2,   67,    2, 0x08 /* Private */,
      12,    0,   72,    2, 0x08 /* Private */,
      13,    0,   73,    2, 0x08 /* Private */,

 // signals: parameters
    QMetaType::Void,
    QMetaType::Void,

 // slots: parameters
    QMetaType::Void,
    QMetaType::Void, QMetaType::QString,    6,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void, QMetaType::Int, QMetaType::QString,   10,   11,
    QMetaType::Void,
    QMetaType::Void,

       0        // eod
};

void EmailOnlyWindow::qt_static_metacall(QObject *_o, QMetaObject::Call _c, int _id, void **_a)
{
    if (_c == QMetaObject::InvokeMetaMethod) {
        auto *_t = static_cast<EmailOnlyWindow *>(_o);
        Q_UNUSED(_t)
        switch (_id) {
        case 0: _t->doProcess(); break;
        case 1: _t->doSend(); break;
        case 2: _t->on_selectFilePushButton_clicked(); break;
        case 3: _t->changeGroupby((*reinterpret_cast< QString(*)>(_a[1]))); break;
        case 4: _t->on_savePathPushButton_clicked(); break;
        case 5: _t->on_emailOnlySubmitPushButton_clicked(); break;
        case 6: _t->receiveMessage((*reinterpret_cast< const int(*)>(_a[1])),(*reinterpret_cast< const QString(*)>(_a[2]))); break;
        case 7: _t->on_gobackPushButton_clicked(); break;
        case 8: _t->showConfigSetting(); break;
        default: ;
        }
    } else if (_c == QMetaObject::IndexOfMethod) {
        int *result = reinterpret_cast<int *>(_a[0]);
        {
            using _t = void (EmailOnlyWindow::*)();
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&EmailOnlyWindow::doProcess)) {
                *result = 0;
                return;
            }
        }
        {
            using _t = void (EmailOnlyWindow::*)();
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&EmailOnlyWindow::doSend)) {
                *result = 1;
                return;
            }
        }
    }
}

QT_INIT_METAOBJECT const QMetaObject EmailOnlyWindow::staticMetaObject = { {
    QMetaObject::SuperData::link<SplitSubWindow::staticMetaObject>(),
    qt_meta_stringdata_EmailOnlyWindow.data,
    qt_meta_data_EmailOnlyWindow,
    qt_static_metacall,
    nullptr,
    nullptr
} };


const QMetaObject *EmailOnlyWindow::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->dynamicMetaObject() : &staticMetaObject;
}

void *EmailOnlyWindow::qt_metacast(const char *_clname)
{
    if (!_clname) return nullptr;
    if (!strcmp(_clname, qt_meta_stringdata_EmailOnlyWindow.stringdata0))
        return static_cast<void*>(this);
    return SplitSubWindow::qt_metacast(_clname);
}

int EmailOnlyWindow::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = SplitSubWindow::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        if (_id < 9)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 9;
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        if (_id < 9)
            *reinterpret_cast<int*>(_a[0]) = -1;
        _id -= 9;
    }
    return _id;
}

// SIGNAL 0
void EmailOnlyWindow::doProcess()
{
    QMetaObject::activate(this, &staticMetaObject, 0, nullptr);
}

// SIGNAL 1
void EmailOnlyWindow::doSend()
{
    QMetaObject::activate(this, &staticMetaObject, 1, nullptr);
}
QT_WARNING_POP
QT_END_MOC_NAMESPACE
