// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CFormatConditions 包装类

class CFormatConditions : public COleDispatchDriver
{
public:
	CFormatConditions(){} // 调用 COleDispatchDriver 默认构造函数
	CFormatConditions(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CFormatConditions(const CFormatConditions& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// FormatConditions 方法
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
	LPDISPATCH Item(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH Add(long Type, VARIANT Operator, VARIANT Formula1, VARIANT Formula2, VARIANT String, VARIANT TextOperator, VARIANT DateOperator, VARIANT ScopeType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0xb5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, &Operator, &Formula1, &Formula2, &String, &TextOperator, &DateOperator, &ScopeType);
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
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH AddColorScale(long ColorScaleType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa38, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ColorScaleType);
		return result;
	}
	LPDISPATCH AddDatabar()
	{
		LPDISPATCH result;
		InvokeHelper(0xa3a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH AddIconSetCondition()
	{
		LPDISPATCH result;
		InvokeHelper(0xa3b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH AddTop10()
	{
		LPDISPATCH result;
		InvokeHelper(0xa3c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH AddAboveAverage()
	{
		LPDISPATCH result;
		InvokeHelper(0xa3d, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH AddUniqueValues()
	{
		LPDISPATCH result;
		InvokeHelper(0xa3e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// FormatConditions 属性
public:

};
