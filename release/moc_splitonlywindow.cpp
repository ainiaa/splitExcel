/****************************************************************************
** Meta object code from reading C++ file 'splitonlywindow.h'
**
** Created by: The Qt Meta Object Compiler version 67 (Qt 5.14.2)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include <memory>
#include "../splitonlywindow.h"
#include <QtCore/qbytearray.h>
#include <QtCore/qmetatype.h>
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'splitonlywindow.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 67
#error "This file was generated using the moc from 5.14.2. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
QT_WARNING_PUSH
QT_WARNING_DISABLE_DEPRECATED
struct qt_meta_stringdata_SplitOnlyWindow_t {
    QByteArrayData data[13];
    char stringdata0[234];
};
#define QT_MOC_LITERAL(idx, ofs, len) \
    Q_STATIC_BYTE_ARRAY_DATA_HEADER_INITIALIZER_WITH_OFFSET(len, \
    qptrdiff(offsetof(qt_meta_stringdata_SplitOnlyWindow_t, stringdata0) + ofs \
        - idx * sizeof(QByteArrayData)) \
    )
static const qt_meta_stringdata_SplitOnlyWindow_t qt_meta_stringdata_SplitOnlyWindow = {
    {
QT_MOC_LITERAL(0, 0, 15), // "SplitOnlyWindow"
QT_MOC_LITERAL(1, 16, 9), // "doProcess"
QT_MOC_LITERAL(2, 26, 0), // ""
QT_MOC_LITERAL(3, 27, 31), // "on_selectFilePushButton_clicked"
QT_MOC_LITERAL(4, 59, 13), // "changeGroupby"
QT_MOC_LITERAL(5, 73, 17), // "selectedSheetName"
QT_MOC_LITERAL(6, 91, 29), // "on_savePathPushButton_clicked"
QT_MOC_LITERAL(7, 121, 36), // "on_splitOnlySubmitPushButton_..."
QT_MOC_LITERAL(8, 158, 14), // "receiveMessage"
QT_MOC_LITERAL(9, 173, 7), // "msgType"
QT_MOC_LITERAL(10, 181, 6), // "result"
QT_MOC_LITERAL(11, 188, 27), // "on_gobackPushButton_clicked"
QT_MOC_LITERAL(12, 216, 17) // "showConfigSetting"

    },
    "SplitOnlyWindow\0doProcess\0\0"
    "on_selectFilePushButton_clicked\0"
    "changeGroupby\0selectedSheetName\0"
    "on_savePathPushButton_clicked\0"
    "on_splitOnlySubmitPushButton_clicked\0"
    "receiveMessage\0msgType\0result\0"
    "on_gobackPushButton_clicked\0"
    "showConfigSetting"
};
#undef QT_MOC_LITERAL

static const uint qt_meta_data_SplitOnlyWindow[] = {

 // content:
       8,       // revision
       0,       // classname
       0,    0, // classinfo
       8,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       1,       // signalCount

 // signals: name, argc, parameters, tag, flags
       1,    0,   54,    2, 0x06 /* Public */,

 // slots: name, argc, parameters, tag, flags
       3,    0,   55,    2, 0x08 /* Private */,
       4,    1,   56,    2, 0x08 /* Private */,
       6,    0,   59,    2, 0x08 /* Private */,
       7,    0,   60,    2, 0x08 /* Private */,
       8,    2,   61,    2, 0x08 /* Private */,
      11,    0,   66,    2, 0x08 /* Private */,
      12,    0,   67,    2, 0x08 /* Private */,

 // signals: parameters
    QMetaType::Void,

 // slots: parameters
    QMetaType::Void,
    QMetaType::Void, QMetaType::QString,    5,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void, QMetaType::Int, QMetaType::QString,    9,   10,
    QMetaType::Void,
    QMetaType::Void,

       0        // eod
};

void SplitOnlyWindow::qt_static_metacall(QObject *_o, QMetaObject::Call _c, int _id, void **_a)
{
    if (_c == QMetaObject::InvokeMetaMethod) {
        auto *_t = static_cast<SplitOnlyWindow *>(_o);
        Q_UNUSED(_t)
        switch (_id) {
        case 0: _t->doProcess(); break;
        case 1: _t->on_selectFilePushButton_clicked(); break;
        case 2: _t->changeGroupby((*reinterpret_cast< QString(*)>(_a[1]))); break;
        case 3: _t->on_savePathPushButton_clicked(); break;
        case 4: _t->on_splitOnlySubmitPushButton_clicked(); break;
        case 5: _t->receiveMessage((*reinterpret_cast< const int(*)>(_a[1])),(*reinterpret_cast< const QString(*)>(_a[2]))); break;
        case 6: _t->on_gobackPushButton_clicked(); break;
        case 7: _t->showConfigSetting(); break;
        default: ;
        }
    } else if (_c == QMetaObject::IndexOfMethod) {
        int *result = reinterpret_cast<int *>(_a[0]);
        {
            using _t = void (SplitOnlyWindow::*)();
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&SplitOnlyWindow::doProcess)) {
                *result = 0;
                return;
            }
        }
    }
}

QT_INIT_METAOBJECT const QMetaObject SplitOnlyWindow::staticMetaObject = { {
    QMetaObject::SuperData::link<SplitSubWindow::staticMetaObject>(),
    qt_meta_stringdata_SplitOnlyWindow.data,
    qt_meta_data_SplitOnlyWindow,
    qt_static_metacall,
    nullptr,
    nullptr
} };


const QMetaObject *SplitOnlyWindow::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->dynamicMetaObject() : &staticMetaObject;
}

void *SplitOnlyWindow::qt_metacast(const char *_clname)
{
    if (!_clname) return nullptr;
    if (!strcmp(_clname, qt_meta_stringdata_SplitOnlyWindow.stringdata0))
        return static_cast<void*>(this);
    return SplitSubWindow::qt_metacast(_clname);
}

int SplitOnlyWindow::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = SplitSubWindow::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        if (_id < 8)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 8;
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        if (_id < 8)
            *reinterpret_cast<int*>(_a[0]) = -1;
        _id -= 8;
    }
    return _id;
}

// SIGNAL 0
void SplitOnlyWindow::doProcess()
{
    QMetaObject::activate(this, &staticMetaObject, 0, nullptr);
}
QT_WARNING_POP
QT_END_MOC_NAMESPACE
