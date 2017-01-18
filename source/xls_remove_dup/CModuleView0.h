// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CModuleView0 包装类

class CModuleView0 : public COleDispatchDriver
{
public:
	CModuleView0(){} // 调用 COleDispatchDriver 默认构造函数
	CModuleView0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CModuleView0(const CModuleView0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ModuleView 方法
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
	LPDISPATCH get_Sheet()
	{
		LPDISPATCH result;
		InvokeHelper(0x2ef, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// ModuleView 属性
public:

};
