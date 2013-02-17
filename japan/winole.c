#include <stdio.h>
#include <malloc.h>
#include <windows.h>
#include <wchar.h>

// 参考 ruby.h
#define ALLOCA_N(type,n) (type*)alloca(sizeof(type)*(n))


IDispatch *mallocDispatch(){
    return (struct IDispatch *)malloc( sizeof(struct IDispatch) );
}

// oleauto.h:#define V_VT(X) ((X)->vt)
// oleauto.h:#define V_ISBYREF(X) (V_VT(X)&VT_BYREF)
IDispatch *Variant2Dispatch(VARIANT *pVariant){
    IDispatch *pDispatch;
    if (V_ISBYREF(pVariant))
        pDispatch = *V_DISPATCHREF(pVariant);
    else
        pDispatch = V_DISPATCH(pVariant);

    return pDispatch;
}
// 常にインタフェーステーブルへアクセスする。
// Open,Close などのコマンドからテーブルのディスパッチIDを求め実行する。
HRESULT ComInvoke( PVOID *p, wchar_t *ComString ,VARIANTARG *param,int nArgs, USHORT wFlags, VARIANT *result){
    IDispatch   *pDisp;
    DISPID      dispID;
    HRESULT     hr;
    unsigned    short *ucPtr;
    UINT        puArgErr = 0;
    EXCEPINFO   excepinfo;

    // http://msdn.microsoft.com/ja-jp/library/x6828bcx%28v=VS.80%29.aspx
    // Win32OLE 製作過程の雑記 : invoke メソッドの引数
    // http://homepage1.nifty.com/markey/ruby/win32ole/win32ole03.html#invoke-param
    DISPPARAMS  dispParams = { NULL, NULL, 0, 0 };
    dispParams.rgvarg            = param;  // 引数の配列への参照を表します。
    dispParams.rgdispidNamedArgs = NULL;   // 名前付き引数の dispID の配列(未使用)
    dispParams.cArgs             = nArgs;  // 引数の数を表します。
    dispParams.cNamedArgs        = 0;      // 名前付き引数の数 (未使用)
    // 参考:ruby win32ole.c  ole_invoke2 関数
    if (wFlags & DISPATCH_PROPERTYPUT) {
        dispParams.cNamedArgs = 1;
        dispParams.rgdispidNamedArgs    = ALLOCA_N( DISPID, 1 );
        dispParams.rgdispidNamedArgs[0] = DISPID_PROPERTYPUT;
    }

    memset( &excepinfo, 0, sizeof(EXCEPINFO));
    pDisp = (IDispatch   *)p;

    // コマンド文字列からディスパッチID取得
    ucPtr = SysAllocString( ComString );
    hr=pDisp->lpVtbl->GetIDsOfNames((IDispatch  *)pDisp, &IID_NULL, &ucPtr, 1, LOCALE_USER_DEFAULT, (DISPID*)&dispID);
    //wprintf(L"GetIDsOfNames nArgs:%d  %-10s = %04d hr:%08lx\n",  nArgs, ComString, dispID, hr);

    // ここが肝心のInvokeを実行する部分。
    VariantInit(result);
    hr = pDisp->lpVtbl->Invoke(
             pDisp,                    // 参考: Ruby付属の「OLE View」
             dispID,                   // arg1 - I4 dispidMember        [IN]
             &IID_NULL,                // arg2 - GUID riid              [IN]
             LOCALE_SYSTEM_DEFAULT,    // arg3 - UI4 lcid               [IN]
             wFlags,                   // arg4 - UI2 wFlags             [IN]
             &dispParams,              // arg5 - DISPPARAMS pdispparams [IN]
             result,                   // arg6 - VARIANT pvarResult     [OUT]
             &excepinfo,               // arg7 - EXCEPINFO pexcepinfo   [OUT]
             &puArgErr );              // arg8 - UINT puArgErr          [OUT]
     //wprintf(L"Invoke %-10s dispID:%4d hr:%08x puArgErr:%d\n",ComString, dispID, hr,puArgErr);
     SysFreeString(ucPtr);
     return hr;
}

// ProgID("Excel.Application")からCLSID({00024500-0000-0000-C000000000000046})
// を求め、CoCreateInstance APIを呼びます。
IDispatch *InstanceNew(wchar_t *ComName){
    IDispatch  *pDisp;
    BSTR       name;
    CLSID      clsid;
    HRESULT    hr=0;

    pDisp = mallocDispatch();
    name  = SysAllocString( ComName );
    hr    = CLSIDFromProgID(name, &clsid);
    // HRESULTは最上位ビットで OK ,NG を表現します。
    // FAILED は hr が 0 より小さいかどうかチェックするマクロ。
    if(FAILED(hr)) {
        hr = CLSIDFromString(name, &clsid);
    }
    hr = CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, &IID_IDispatch, (void **)&pDisp);
    SysFreeString(name);
    return pDisp;
}

wchar_t *Date2String(DATE date){
    wchar_t *buf;
    SYSTEMTIME st;

    VariantTimeToSystemTime(date, &st);
    buf = (wchar_t*)malloc(20 * sizeof(wchar_t));
    //swprintf(buf,L"%04d/%02d/%02d %02d:%02d:%02d",st.wYear,st.wMonth,st.wDay,st.wHour,st.wMinute,st.wSecond);
    return buf;
}
wchar_t *Number2String(long num){
    wchar_t *buf;
    buf = (wchar_t*)malloc(30 * sizeof(wchar_t));
    swprintf(buf,L"%d",num);
    return buf;
}

wchar_t *Double2String(double num){
    wchar_t *buf;
    buf = (wchar_t*)malloc(30 * sizeof(wchar_t));
    swprintf(buf,L"%f",num);
    return buf;
}

wchar_t *Variant2CWString(VARIANT *result){
    switch(V_VT(result)){
        case VT_EMPTY:
            return L"empty";
            break;
        case VT_NULL:
            return L"null";
            break;
        case VT_I2:  // short
            return Number2String((long)V_I2(result));
            break;
        case VT_I4:  // long
            return Number2String((long)V_I4(result));
            break;
        case VT_R4:  // float
            return Double2String(V_R4(result));
            break;
        case VT_R8:  // double
            return Double2String(V_R8(result));
            break;
        case VT_BOOL: //(True -1,False 0)
            return (V_BOOL(result) ? L"True" : L"False");
            break;
        case VT_BSTR:
            return (wchar_t*)(V_BSTR(result));
            break;
        case VT_DATE:
            return Date2String( V_DATE(result));
            break;
    }
}

// PropertyPut_S_S((void **)cell, L"Value",L"ほげ");
void PropertyPut_S_S(PVOID *pDisp, wchar_t *PropertyName, wchar_t *String){
    VARIANT     result;
    VARIANTARG  param[1];
    BSTR        bstr;

    bstr = SysAllocString(String);
    VariantInit(&param[0]);
    param[0].vt = VT_BSTR|VT_BYREF;
    param[0].pbstrVal = &bstr;
    ComInvoke((void **)pDisp, PropertyName, param, 1, DISPATCH_PROPERTYPUT, &result);
    VariantClear(&result);
    VariantClear(&param[0]);
    SysFreeString(bstr);
}

// GetAbsolutePathName メソッドをコールし、パス名を含めたファイル名を取得
wchar_t *GetPathName(IDispatch *fDisp, wchar_t *fileName){
    VARIANT    param, result;
    HRESULT    hr = 0;
    wchar_t *fullPathName;

    VariantInit(&param);
    param.vt      = VT_BSTR;
    param.bstrVal = SysAllocString(fileName);

    hr = ComInvoke((void **)fDisp, L"GetAbsolutePathName", &param, 1, DISPATCH_METHOD, &result);
    fullPathName  = (wchar_t*)malloc((SysStringLen(result.bstrVal)+1) * sizeof(wchar_t));
    wcscpy(fullPathName, result.bstrVal);
    SysFreeString(param.bstrVal);
    VariantClear(&param);
    VariantClear(&result);
    return fullPathName;
}
// Scripting.FileSystemObject を作りパス名を含めたファイル名を取得
// in  : fileName
// out : fullPathName
wchar_t *getFullPathName(wchar_t *fileName){
    return GetPathName(InstanceNew(L"Scripting.FileSystemObject"), fileName);
}

//   workBooks   = PropertyGet_S((void **)pExl, L"Workbooks");
IDispatch *PropertyGet_S( PVOID *parentDisp, wchar_t *ObjName){
    VARIANT    param, result;
    DISPID     dispID;
    HRESULT    hr = 0;

    VariantInit(&param);
    VariantInit(&result);
    param.vt = VT_EMPTY;
    hr = ComInvoke((void **)parentDisp, ObjName, &param, 0,DISPATCH_PROPERTYGET | DISPATCH_METHOD,&result);
    // wprintf(L"CreateNewObject   ObjName:%-14s hr:%08lx\n",ObjName,hr);
    VariantClear(&param);
    return Variant2Dispatch(&result);
}

//  sheet  = PropertyGet_S_N( (void **)workSheets, L"Item", 2); // 2 番目のシート
IDispatch *PropertyGet_S_N(PVOID *pDisp, wchar_t *str, long n){
    VARIANT     result;
    VARIANTARG  param[1];
    HRESULT     hr = 0;

    VariantInit(&param[0]);  param[0].vt = VT_I4;   param[0].lVal = n;
    ComInvoke((void **)pDisp, str, param, 1, DISPATCH_PROPERTYGET , &result);
    VariantClear(&param[0]);
    return Variant2Dispatch(&result);
}

// cell = PropertyGet_S_S((void **)sheet, L"Range", L"C2");
IDispatch *PropertyGet_S_S(PVOID *pDisp, wchar_t *str1, wchar_t *str2){
    VARIANT     result;
    VARIANTARG  param[1];
    BSTR        bstr;
    HRESULT     hr = 0;

    bstr = SysAllocString(str2);

    VariantInit(&param[0]);
    param[0].vt = VT_BSTR|VT_BYREF;
    param[0].pbstrVal = &bstr;
    ComInvoke((void **)pDisp, str1, param, 1, DISPATCH_PROPERTYGET , &result);
    VariantClear(&param[0]);
    SysFreeString(bstr);
    return Variant2Dispatch(&result);
}
// ver = ReadProperty((void **)pExl, L"Version");
wchar_t *ReadProperty(PVOID *pDisp, wchar_t *PropertyName){
    VARIANT    param, result;

    VariantInit(&param);
    param.vt      = VT_EMPTY;
    ComInvoke((void **)pDisp, PropertyName,&param, 0,DISPATCH_PROPERTYGET | DISPATCH_METHOD, &result);
    VariantClear(&param);
    return Variant2CWString(&result);
}


// call Method_S_S((void **)workBooks, "Open", "C:\\example.xls");
void Method_S_S(PVOID *pDisp, wchar_t *str1, wchar_t *str2){
    VARIANT     result;
    VARIANTARG  param[1];
    BSTR        bstr;

    bstr = SysAllocString(str2);
    VariantInit(&param[0]);  param[0].vt = VT_BSTR|VT_BYREF;param[0].pbstrVal = &bstr;
    ComInvoke((void **)pDisp, str1, param, 1, DISPATCH_METHOD, &result);
    SysFreeString(bstr);
    VariantClear(&result);
}


// call Method_S((void **)workBooks, L"Close");
// call Method_S((void **)pExl, L"Quit");
void Method_S(PVOID *pDisp, wchar_t *command){
    VARIANT    param, result;

    VariantInit(&param);
    param.vt      = VT_EMPTY;
    ComInvoke((void **)pDisp, command, &param, 0, DISPATCH_METHOD,&result);
    VariantClear(&param);
    VariantClear(&result);
}

// ReleaseObject((void **)pExl);
void ReleaseObject( PVOID *pDisp ){
    ((IDispatch  *)pDisp)->lpVtbl->Release( (void *)pDisp);
}


