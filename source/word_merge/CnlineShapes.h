// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CnlineShapes 包装类

class CnlineShapes : public COleDispatchDriver
{
public:
	CnlineShapes(){} // 调用 COleDispatchDriver 默认构造函数
	CnlineShapes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CnlineShapes(const CnlineShapes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// InlineShapes 方法
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
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH AddPicture(LPCTSTR FileName, VARIANT * LinkToFile, VARIANT * SaveWithDocument, VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, LinkToFile, SaveWithDocument, Range);
		return result;
	}
	LPDISPATCH AddOLEObject(VARIANT * ClassType, VARIANT * FileName, VARIANT * LinkToFile, VARIANT * DisplayAsIcon, VARIANT * IconFileName, VARIANT * IconIndex, VARIANT * IconLabel, VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x18, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Range);
		return result;
	}
	LPDISPATCH AddOLEControl(VARIANT * ClassType, VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ClassType, Range);
		return result;
	}
	LPDISPATCH New(LPDISPATCH Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xc8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range);
		return result;
	}
	LPDISPATCH AddHorizontalLine(LPCTSTR FileName, VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, Range);
		return result;
	}
	LPDISPATCH AddHorizontalLineStandard(VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x69, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range);
		return result;
	}
	LPDISPATCH AddPictureBullet(LPCTSTR FileName, VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, Range);
		return result;
	}
	LPDISPATCH AddChart(long Type, VARIANT * Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x6b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, Range);
		return result;
	}

	// InlineShapes 属性
public:

};
