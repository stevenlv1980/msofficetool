// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COMathAcc 包装类

class COMathAcc : public COleDispatchDriver
{
public:
	COMathAcc(){} // 调用 COleDispatchDriver 默认构造函数
	COMathAcc(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COMathAcc(const COMathAcc& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// OMathAcc 方法
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
	LPDISPATCH get_E()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	short get_Char()
	{
		short result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
		return result;
	}
	void put_Char(short newValue)
	{
		static BYTE parms[] = VTS_I2 ;
		InvokeHelper(0x68, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// OMathAcc 属性
public:

};
