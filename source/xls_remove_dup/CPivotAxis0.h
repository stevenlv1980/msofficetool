// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CPivotAxis0 包装类

class CPivotAxis0 : public COleDispatchDriver
{
public:
	CPivotAxis0(){} // 调用 COleDispatchDriver 默认构造函数
	CPivotAxis0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPivotAxis0(const CPivotAxis0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// PivotAxis 方法
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
	LPDISPATCH get_PivotLines()
	{
		LPDISPATCH result;
		InvokeHelper(0xa7d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// PivotAxis 属性
public:

};
