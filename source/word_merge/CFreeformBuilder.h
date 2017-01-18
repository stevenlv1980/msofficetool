// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CFreeformBuilder 包装类

class CFreeformBuilder : public COleDispatchDriver
{
public:
	CFreeformBuilder(){} // 调用 COleDispatchDriver 默认构造函数
	CFreeformBuilder(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFreeformBuilder(const CFreeformBuilder& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// FreeformBuilder 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void AddNodes(long SegmentType, long EditingType, float X1, float Y1, float X2, float Y2, float X3, float Y3)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3);
	}
	LPDISPATCH ConvertToShape(VARIANT * Anchor)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Anchor);
		return result;
	}

	// FreeformBuilder 属性
public:

};
