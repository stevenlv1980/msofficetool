// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COLEFormat0 包装类

class COLEFormat0 : public COleDispatchDriver
{
public:
	COLEFormat0(){} // 调用 COleDispatchDriver 默认构造函数
	COLEFormat0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COLEFormat0(const COLEFormat0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// OLEFormat 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Activate()
	{
		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH get_Object()
	{
		LPDISPATCH result;
		InvokeHelper(0x419, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_progID()
	{
		CString result;
		InvokeHelper(0x5f3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void Verb(VARIANT Verb)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x25e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Verb);
	}

	// OLEFormat 属性
public:

};
