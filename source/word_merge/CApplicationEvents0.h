// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CApplicationEvents0 包装类

class CApplicationEvents0 : public COleDispatchDriver
{
public:
	CApplicationEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	CApplicationEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents0(const CApplicationEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IApplicationEvents 方法
public:
	void Startup()
	{
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Quit()
	{
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DocumentChange()
	{
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// IApplicationEvents 属性
public:

};
