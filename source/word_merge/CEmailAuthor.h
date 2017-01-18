// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CEmailAuthor 包装类

class CEmailAuthor : public COleDispatchDriver
{
public:
	CEmailAuthor(){} // 调用 COleDispatchDriver 默认构造函数
	CEmailAuthor(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEmailAuthor(const CEmailAuthor& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// EmailAuthor 方法
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
	LPDISPATCH get_Style()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// EmailAuthor 属性
public:

};
