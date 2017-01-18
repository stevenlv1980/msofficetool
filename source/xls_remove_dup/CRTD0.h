// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRTD0 包装类

class CRTD0 : public COleDispatchDriver
{
public:
	CRTD0(){} // 调用 COleDispatchDriver 默认构造函数
	CRTD0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRTD0(const CRTD0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// RTD 方法
public:
	long get_ThrottleInterval()
	{
		long result;
		InvokeHelper(0x8c0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ThrottleInterval(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8c0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void RefreshData()
	{
		InvokeHelper(0x8c1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void RestartServers()
	{
		InvokeHelper(0x8c2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// RTD 属性
public:

};
