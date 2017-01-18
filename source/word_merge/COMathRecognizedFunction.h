// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COMathRecognizedFunction 包装类

class COMathRecognizedFunction : public COleDispatchDriver
{
public:
	COMathRecognizedFunction(){} // 调用 COleDispatchDriver 默认构造函数
	COMathRecognizedFunction(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COMathRecognizedFunction(const COMathRecognizedFunction& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// OMathRecognizedFunction 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Index()
	{
		long result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0xc8, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// OMathRecognizedFunction 属性
public:

};
