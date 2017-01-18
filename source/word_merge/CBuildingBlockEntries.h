// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CBuildingBlockEntries 包装类

class CBuildingBlockEntries : public COleDispatchDriver
{
public:
	CBuildingBlockEntries(){} // 调用 COleDispatchDriver 默认构造函数
	CBuildingBlockEntries(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CBuildingBlockEntries(const CBuildingBlockEntries& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// BuildingBlockEntries 方法
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
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Name, long Type, LPCTSTR Category, LPDISPATCH Range, VARIANT * Description, long InsertOptions)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_BSTR VTS_DISPATCH VTS_PVARIANT VTS_I4 ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Type, Category, Range, Description, InsertOptions);
		return result;
	}

	// BuildingBlockEntries 属性
public:

};
