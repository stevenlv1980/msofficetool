// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CColorScaleCriterion 包装类

class CColorScaleCriterion : public COleDispatchDriver
{
public:
	CColorScaleCriterion(){} // 调用 COleDispatchDriver 默认构造函数
	CColorScaleCriterion(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CColorScaleCriterion(const CColorScaleCriterion& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ColorScaleCriterion 方法
public:
	long get_Index()
	{
		long result;
		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
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
	VARIANT get_Value()
	{
		VARIANT result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Value(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	LPDISPATCH get_FormatColor()
	{
		LPDISPATCH result;
		InvokeHelper(0xa9d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// ColorScaleCriterion 属性
public:

};
