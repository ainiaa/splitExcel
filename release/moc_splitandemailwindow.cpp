/****************************************************************************
** Meta object code from reading C++ file 'splitandemailwindow.h'
**
** Created by: The Qt Meta Object Compiler version 67 (Qt 5.14.2)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include <memory>
#include "../splitandemailwindow.h"
#include <QtCore/qbytearray.h>
#include <QtCore/qmetatype.h>
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'splitandemailwindow.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 67
#error "This file was generated using the moc from 5.14.2. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
QT_WARNING_PUSH
QT_WARNING_DISABLE_DEPRECATED
struct qt_meta_stringdata_SplitAndEmailWindow_t {
    QByteArrayData data[16];
    char stringdata0[277];
};
#define QT_MOC_LITERAL(idx, ofs, len) \
    Q_STATIC_BYTE_ARRAY_DATA_HEADER_INITIALIZER_WITH_OFFSET(len, \
    qptrdiff(offsetof(qt_meta_stringdata_SplitAndEmailWindow_t, stringdata0) + ofs \
        - idx * sizeof(QByteArrayData)) \
    )
static const qt_meta_stringdata_SplitAndEmailWindow_t qt_meta_stringdata_SplitAndEmailWindow = {
    {
QT_MOC_LITERAL(0, 0, 19), // "SplitAndEmailWindow"
QT_MOC_LITERAL(1, 20, 6), // "doSend"
QT_MOC_LITERAL(2, 27, 0), // ""
QT_MOC_LITERAL(3, 28, 9), // "doProcess"
QT_MOC_LITERAL(4, 38, 31), // "on_selectFilePushButton_clicked"
QT_MOC_LITERAL(5, 70, 29), // "on_savePathPushButton_clicked"
QT_MOC_LITERAL(6, 100, 13), // "changeGroupby"
QT_MOC_LITERAL(7, 114, 17), // "selectedSheetName"
QT_MOC_LITERAL(8, 132, 34), // "on_dataComboBox_currentTextCh..."
QT_MOC_LITERAL(9, 167, 4), // "arg1"
QT_MOC_LITERAL(10, 172, 17), // "showConfigSetting"
QT_MOC_LITERAL(11, 190, 14), // "receiveMessage"
QT_MOC_LITERAL(12, 205, 7), // "msgType"
QT_MOC_LITERAL(13, 213, 6), // "result"
QT_MOC_LITERAL(14, 220, 27), // "on_gobackPushButton_clicked"
QT_MOC_LITERAL(15, 248, 28) // "on_submit2PushButton_clicked"

    },
    "SplitAndEmailWindow\0doSend\0\0doProcess\0"
    "on_selectFilePushButton_clicked\0"
    "on_savePathPushButton_clicked\0"
    "changeGroupby\0selectedSheetName\0"
    "on_dataComboBox_currentTextChanged\0"
    "arg1\0showConfigSetting\0receiveMessage\0"
    "msgType\0result\0on_gobackPushButton_clicked\0"
    "on_submit2PushButton_clicked"
};
#undef QT_MOC_LITERAL

static const uint qt_meta_data_SplitAndEmailWindow[] = {

 // content:
       8,       // revision
       0,       // classname
       0,    0, // classinfo
      10,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       2,       // signalCount

 // signals: name, argc, parameters, tag, flags
       1,    0,   64,    2, 0x06 /* Public */,
       3,    0,   65,    2, 0x06 /* Public */,

 // slots: name, argc, parameters, tag, flags
       4,    0,   66,    2, 0x08 /* Private */,
       5,    0,   67,    2, 0x08 /* Private */,
       6,    1,   68,    2, 0x08 /* Private */,
       8,    1,   71,    2, 0x08 /* Private */,
      10,    0,   74,    2, 0x08 /* Private */,
      11,    2,   75,    2, 0x08 /* Private */,
      14,    0,   80,    2, 0x08 /* Private */,
      15,    0,   81,    2, 0x08 /* Private */,

 // signals: parameters
    QMetaType::Void,
    QMetaType::Void,

 // slots: parameters
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void, QMetaType::QString,    7,
    QMetaType::Void, QMetaType::QString,    9,
    QMetaType::Void,
    QMetaType::Void, QMetaType::Int, QMetaType::QString,   12,   13,
    QMetaType::Void,
    QMetaType::Void,

       0        // eod
};

void SplitAndEmailWindow::qt_static_metacall(QObject *_o, QMetaObject::Call _c, int _id, void **_a)
{
    if (_c == QMetaObject::InvokeMetaMethod) {
        auto *_t = static_cast<SplitAndEmailWindow *>(_o);
        Q_UNUSED(_t)
        switch (_id) {
        case 0: _t->doSend(); break;
        case 1: _t->doProcess(); break;
        case 2: _t->on_selectFilePushButton_clicked(); break;
        case 3: _t->on_savePathPushButton_clicked(); break;
        case 4: _t->changeGroupby((*reinterpret_cast< QString(*)>(_a[1]))); break;
        case 5: _t->on_dataComboBox_currentTextChanged((*reinterpret_cast< const QString(*)>(_a[1]))); break;
        case 6: _t->showConfigSetting(); break;
        case 7: _t->receiveMessage((*reinterpret_cast< const int(*)>(_a[1])),(*reinterpret_cast< const QString(*)>(_a[2]))); break;
        case 8: _t->on_gobackPushButton_clicked(); break;
        case 9: _t->on_submit2PushButton_clicked(); break;
        default: ;
        }
    } else if (_c == QMetaObject::IndexOfMethod) {
        int *result = reinterpret_cast<int *>(_a[0]);
        {
            using _t = void (SplitAndEmailWindow::*)();
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&SplitAndEmailWindow::doSend)) {
                *result = 0;
                return;
            }
        }
        {
            using _t = void (SplitAndEmailWindow::*)();
            if (*reinterpret_cast<_t *>(_a[1]) == static_cast<_t>(&SplitAndEmailWindow::doProcess)) {
                *result = 1;
                return;
            }
        }
    }
}

QT_INIT_METAOBJECT const QMetaObject SplitAndEmailWindow::staticMetaObject = { {
    QMetaObject::SuperData::link<SplitSubWindow::staticMetaObject>(),
    qt_meta_stringdata_SplitAndEmailWindow.data,
    qt_meta_data_SplitAndEmailWindow,
    qt_static_metacall,
    nullptr,
    nullptr
} };


const QMetaObject *SplitAndEmailWindow::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->dynamicMetaObject() : &staticMetaObject;
}

void *SplitAndEmailWindow::qt_metacast(const char *_clname)
{
    if (!_clname) return nullptr;
    if (!strcmp(_clname, qt_meta_stringdata_SplitAndEmailWindow.stringdata0))
        return static_cast<void*>(this);
    return SplitSubWindow::qt_metacast(_clname);
}

int SplitAndEmailWindow::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = SplitSubWindow::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        if (_id < 10)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 10;
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        if (_id < 10)
            *reinterpret_cast<int*>(_a[0]) = -1;
        _id -= 10;
    }
    return _id;
}

// SIGNAL 0
void SplitAndEmailWindow::doSend()
{
    QMetaObject::activate(this, &staticMetaObject, 0, nullptr);
}

// SIGNAL 1
void SplitAndEmailWindow::doProcess()
{
    QMetaObject::activate(this, &staticMetaObject, 1, nullptr);
}
QT_WARNING_POP
QT_END_MOC_NAMESPACE
