// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDialogSheetView 包装类

class CDialogSheetView : public COleDispatchDriver
{
public:
	CDialogSheetView(){} // 调用 COleDispatchDriver 默认构造函数
	CDialogSheetView(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDialogSheetView(const CDialogSheetView& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// DialogSheetView 方法
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

	// DialogSheetView 属性
public:

};
