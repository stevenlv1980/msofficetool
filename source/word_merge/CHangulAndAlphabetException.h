// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CHangulAndAlphabetException 包装类

class CHangulAndAlphabetException : public COleDispatchDriver
{
public:
	CHangulAndAlphabetException(){} // 调用 COleDispatchDriver 默认构造函数
	CHangulAndAlphabetException(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CHangulAndAlphabetException(const CHangulAndAlphabetException& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// HangulAndAlphabetException 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Index()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// HangulAndAlphabetException 属性
public:

};
