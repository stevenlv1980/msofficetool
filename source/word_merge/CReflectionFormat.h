// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CReflectionFormat 包装类

class CReflectionFormat : public COleDispatchDriver
{
public:
	CReflectionFormat(){} // 调用 COleDispatchDriver 默认构造函数
	CReflectionFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CReflectionFormat(const CReflectionFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ReflectionFormat 方法
public:
	long get_Type()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Type(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// ReflectionFormat 属性
public:

};
