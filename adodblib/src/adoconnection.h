#ifndef ADOCONENCTION_H
#define ADOCONENCTION_H

#include <QAxObject>
#include <QVariant>

class QTimer;

class AdoConnection : public QObject {
    Q_OBJECT
    public:
    explicit AdoConnection(QObject *parent = nullptr);
    bool open(const QString &connectString);
    bool open();
    bool execute(const QString &sql);
    QVariant connection();

    bool isOpen() const;
    void close();

    protected slots:
    void exception(int code, const QString &source, const QString &desc, const QString &help);
    void disconnect();

    private:
    QAxObject *object;
    QString openString;
    QTimer *timer;
};

#endif // ADOCONENCTION_H
