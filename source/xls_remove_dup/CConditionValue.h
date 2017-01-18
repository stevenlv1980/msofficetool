// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CConditionValue 包装类

class CConditionValue : public COleDispatchDriver
{
public:
	CConditionValue(){} // 调用 COleDispatchDriver 默认构造函数
	CConditionValue(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CConditionValue(const CConditionValue& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ConditionValue 方法
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
	void Modify(long newtype, VARIANT newvalue)
	{
		static BYTE parms[] = VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x62d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, newtype, &newvalue);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Value()
	{
		VARIANT result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}

	// ConditionValue 属性
public:

};
