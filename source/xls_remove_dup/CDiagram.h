// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDiagram 包装类

class CDiagram : public COleDispatchDriver
{
public:
	CDiagram(){} // 调用 COleDispatchDriver 默认构造函数
	CDiagram(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDiagram(const CDiagram& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Diagram 方法
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
	LPDISPATCH get_Nodes()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_AutoLayout()
	{
		long result;
		InvokeHelper(0x8c3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_AutoLayout(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8c3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Reverse()
	{
		long result;
		InvokeHelper(0x8c4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Reverse(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x8c4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_AutoFormat()
	{
		long result;
		InvokeHelper(0x72, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_AutoFormat(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x72, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Convert(long Type)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x416, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type);
	}
	void FitText()
	{
		InvokeHelper(0x900, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Diagram 属性
public:

};
