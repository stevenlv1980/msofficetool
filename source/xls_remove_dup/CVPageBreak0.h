// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CVPageBreak0 包装类

class CVPageBreak0 : public COleDispatchDriver
{
public:
	CVPageBreak0(){} // 调用 COleDispatchDriver 默认构造函数
	CVPageBreak0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CVPageBreak0(const CVPageBreak0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// VPageBreak 方法
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
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DragOff(long Direction, long RegionIndex)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x58c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Direction, RegionIndex);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Type(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Extent()
	{
		long result;
		InvokeHelper(0x58e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Location()
	{
		LPDISPATCH result;
		InvokeHelper(0x575, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void putref_Location(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x575, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
	}

	// VPageBreak 属性
public:

};
