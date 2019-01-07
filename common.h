#ifndef COMMON_H
#define COMMON_H
class Common
{
public:
    static const int MsgTypeError  = 1;
    static const int MsgTypeWarn = 2;
    static const int MsgTypeInfo   = 3;
    static const int MsgTypeFail = 4;
    static const int MsgTypeSucc = 5;
    static const int MsgTypeFinish = 6;
    static const int MsgTypeEmailSendFinish = 7;

    static const int EmailColumnMinCnt =4;
};

#endif // COMMON_H
