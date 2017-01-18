// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CPivotLine0 包装类

class CPivotLine0 : public COleDispatchDriver
{
public:
	CPivotLine0(){} // 调用 COleDispatchDriver 默认构造函数
	CPivotLine0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPivotLine0(const CPivotLine0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// PivotLine 方法
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
	long get_LineType()
	{
		long result;
		InvokeHelper(0xa7b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Position()
	{
		long result;
		InvokeHelper(0x85, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_PivotLineCells()
	{
		LPDISPATCH result;
		InvokeHelper(0xa7c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// PivotLine 属性
public:

};
