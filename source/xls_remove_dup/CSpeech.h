// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSpeech 包装类

class CSpeech : public COleDispatchDriver
{
public:
	CSpeech(){} // 调用 COleDispatchDriver 默认构造函数
	CSpeech(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSpeech(const CSpeech& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ISpeech 方法
public:
	STDMETHOD(Speak)(LPCTSTR Text, VARIANT SpeakAsync, VARIANT SpeakXML, VARIANT Purge)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x7e1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Text, &SpeakAsync, &SpeakXML, &Purge);
		return result;
	}
	STDMETHOD(get_Direction)(XlSpeakDirection * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0xa8, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Direction)(XlSpeakDirection RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa8, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_SpeakCellOnEnter)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x8bb, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_SpeakCellOnEnter)(BOOL RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x8bb, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}

	// ISpeech 属性
public:

};
