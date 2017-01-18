// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRTDUpdateEvent 包装类

class CRTDUpdateEvent : public COleDispatchDriver
{
public:
	CRTDUpdateEvent(){} // 调用 COleDispatchDriver 默认构造函数
	CRTDUpdateEvent(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRTDUpdateEvent(const CRTDUpdateEvent& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IRTDUpdateEvent 方法
public:
	void UpdateNotify()
	{
		InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long get_HeartbeatInterval()
	{
		long result;
		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_HeartbeatInterval(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Disconnect()
	{
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// IRTDUpdateEvent 属性
public:

};
