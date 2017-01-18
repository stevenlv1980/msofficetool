// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRoutingSlip0 包装类

class CRoutingSlip0 : public COleDispatchDriver
{
public:
	CRoutingSlip0(){} // 调用 COleDispatchDriver 默认构造函数
	CRoutingSlip0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRoutingSlip0(const CRoutingSlip0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// RoutingSlip 方法
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
	long get_Delivery()
	{
		long result;
		InvokeHelper(0x3bb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Delivery(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3bb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT get_Message()
	{
		VARIANT result;
		InvokeHelper(0x3ba, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Message(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x3ba, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	VARIANT get_Recipients(VARIANT Index)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x3b8, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, &Index);
		return result;
	}
	void put_Recipients(VARIANT Index, VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x3b8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &Index, &newValue);
	}
	VARIANT Reset()
	{
		VARIANT result;
		InvokeHelper(0x22b, DISPATCH_METHOD, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	BOOL get_ReturnWhenDone()
	{
		BOOL result;
		InvokeHelper(0x3bc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ReturnWhenDone(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x3bc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Status()
	{
		long result;
		InvokeHelper(0x3be, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Subject()
	{
		VARIANT result;
		InvokeHelper(0x3b9, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Subject(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x3b9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	BOOL get_TrackStatus()
	{
		BOOL result;
		InvokeHelper(0x3bd, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_TrackStatus(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x3bd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// RoutingSlip 属性
public:

};
