// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CWatch0 包装类

class CWatch0 : public COleDispatchDriver
{
public:
	CWatch0(){} // 调用 COleDispatchDriver 默认构造函数
	CWatch0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWatch0(const CWatch0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Watch 方法
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
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	VARIANT get_Source()
	{
		VARIANT result;
		InvokeHelper(0xde, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}

	// Watch 属性
public:

};
