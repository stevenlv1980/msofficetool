// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSpeech0 包装类

class CSpeech0 : public COleDispatchDriver
{
public:
	CSpeech0(){} // 调用 COleDispatchDriver 默认构造函数
	CSpeech0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSpeech0(const CSpeech0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Speech 方法
public:
	void Speak(LPCTSTR Text, VARIANT SpeakAsync, VARIANT SpeakXML, VARIANT Purge)
	{
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x7e1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Text, &SpeakAsync, &SpeakXML, &Purge);
	}
	long get_Direction()
	{
		long result;
		InvokeHelper(0xa8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Direction(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_SpeakCellOnEnter()
	{
		BOOL result;
		InvokeHelper(0x8bb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_SpeakCellOnEnter(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x8bb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// Speech 属性
public:

};
