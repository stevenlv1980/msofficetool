// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CSoftEdgeFormat 包装类

class CSoftEdgeFormat : public COleDispatchDriver
{
public:
	CSoftEdgeFormat(){} // 调用 COleDispatchDriver 默认构造函数
	CSoftEdgeFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSoftEdgeFormat(const CSoftEdgeFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// SoftEdgeFormat 方法
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

	// SoftEdgeFormat 属性
public:

};
