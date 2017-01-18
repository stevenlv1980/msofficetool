// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSortField0 包装类

class CSortField0 : public COleDispatchDriver
{
public:
	CSortField0(){} // 调用 COleDispatchDriver 默认构造函数
	CSortField0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSortField0(const CSortField0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SortField 方法
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
	long get_SortOn()
	{
		long result;
		InvokeHelper(0xab5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_SortOn(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xab5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_SortOnValue()
	{
		LPDISPATCH result;
		InvokeHelper(0xab6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Key()
	{
		LPDISPATCH result;
		InvokeHelper(0x9b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Order()
	{
		long result;
		InvokeHelper(0xc0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Order(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xc0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT get_CustomOrder()
	{
		VARIANT result;
		InvokeHelper(0xab7, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_CustomOrder(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xab7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	long get_DataOption()
	{
		long result;
		InvokeHelper(0xab8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DataOption(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xab8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Priority()
	{
		long result;
		InvokeHelper(0x3d9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Priority(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3d9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ModifyKey(LPDISPATCH Key)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xab9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Key);
	}
	void SetIcon(LPDISPATCH Icon)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xaba, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Icon);
	}

	// SortField 属性
public:

};
