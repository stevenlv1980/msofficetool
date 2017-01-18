// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CServerViewableItems0 包装类

class CServerViewableItems0 : public COleDispatchDriver
{
public:
	CServerViewableItems0(){} // 调用 COleDispatchDriver 默认构造函数
	CServerViewableItems0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CServerViewableItems0(const CServerViewableItems0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ServerViewableItems 方法
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Add(VARIANT Obj)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Obj);
		return result;
	}
	void Delete(VARIANT Index)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Index);
	}
	void DeleteAll()
	{
		InvokeHelper(0x8bf, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Item(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH get__Default(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}

	// ServerViewableItems 属性
public:

};
