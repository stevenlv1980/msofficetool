// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CMultiThreadedCalculation0 包装类

class CMultiThreadedCalculation0 : public COleDispatchDriver
{
public:
	CMultiThreadedCalculation0(){} // 调用 COleDispatchDriver 默认构造函数
	CMultiThreadedCalculation0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMultiThreadedCalculation0(const CMultiThreadedCalculation0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// MultiThreadedCalculation 方法
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
	BOOL get_Enabled()
	{
		BOOL result;
		InvokeHelper(0x258, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Enabled(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x258, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ThreadMode()
	{
		long result;
		InvokeHelper(0xace, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ThreadMode(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xace, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ThreadCount()
	{
		long result;
		InvokeHelper(0xacf, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ThreadCount(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xacf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// MultiThreadedCalculation 属性
public:

};
