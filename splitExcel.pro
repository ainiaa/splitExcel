#-------------------------------------------------
#
# Project created by QtCreator 2018-12-20T14:56:04
#
#-------------------------------------------------

QT       += core gui axcontainer

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = splitExcel
TEMPLATE = app

# The following define makes your compiler emit warnings if you use
# any feature of Qt which has been marked as deprecated (the exact warnings
# depend on your compiler). Please consult the documentation of the
# deprecated API in order to know how to port your code away from it.
DEFINES += QT_DEPRECATED_WARNINGS

# You can also make your code fail to compile if you use deprecated APIs.
# In order to do so, uncomment the following line.
# You can also select to disable deprecated APIs only up to a certain version of Qt.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

CONFIG += c++11

SOURCES += \
        main.cpp \
        mainwindow.cpp \
        configsetting.cpp \
        config.cpp \
        common.cpp \
        processwindow.cpp \
        emailsender.cpp \
        emailsenderrunnable.cpp \
        splitonlywindow.cpp \
        testtimeer.cpp \
        emailtaskqueuedata.cpp \
        officehelper.cpp \
        qstringutil.cpp \
        excelbase.cpp \
        excelparser.cpp \
        excelparserbylib.cpp \
        excelparserbyoffice.cpp \
        excelparserbyofficerunnable.cpp \
        excelparserbylibrunnable.cpp \
        sourceexceldata.cpp \
    splitandemailwindow.cpp \
    emailonlywindow.cpp \
    splitsubwindow.cpp


HEADERS += \
        mainwindow.h \
        configsetting.h \
        config.h \
        common.h \
        processwindow.h \
        emailsender.h \
        emailsenderrunnable.h \
        splitonlywindow.h \
        testtimeer.h \
        emailtaskqueuedata.h \
        officehelper.h \
        qstringutil.h \
        excelbase.h \
        excelparser.h \
        excelparserbylib.h \
        excelparserbyoffice.h \
        excelparserbyofficerunnable.h \
        excelparserbylibrunnable.h \
        sourceexceldata.h \
        iexcelparserrunnable.h \
    splitandemailwindow.h \
    emailonlywindow.h \
    splitsubwindow.h

FORMS += \
        mainwindow.ui \
        configsetting.ui \
        processwindow.ui \
    splitonlywindow.ui \
    testtimeer.ui \
    splitandemailwindow.ui \
    emailonlywindow.ui

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

# qtxlsx
#include(xlsx/qtxlsx.pri)
QXLSX_PARENTPATH=QXlsx/         # current QXlsx path is . (. means curret directory)
QXLSX_HEADERPATH=QXlsx/header/  # current QXlsx header path is ./header/
QXLSX_SOURCEPATH=QXlsx/source/  # current QXlsx source path is ./source/
include(QXlsx/QXlsx.pri)


# stmp client
include(smtpclient/smtpclient.pri)

RC_FILE = splitExcel.rc

RESOURCES += \
    images.qrc
