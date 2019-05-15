#ifndef ADO_H 
#define ADO_H 
 
#include <QAxObject> 
 
#define adConnectUnspecified -1 
 
#define adStateClosed 0 
 
#define adOpenStatic  3 
#define adOpenDynamic 2 
 
#define adLockOptimistic 3 
 
#define adCmdText 1 
 
typedef long HRESULT; 
 
#define SUCCEEDED(hr) ((HRESULT)(hr) >= 0) 
#define FAILED(hr)    ((HRESULT)(hr) < 0) 
 
#define ADO_DELETE(p) do{if(p) delete (p); (p)=0;}while(0) 
 
#endif // ADO_H 