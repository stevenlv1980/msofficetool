// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// COCXEvents 包装类

class COCXEvents : public COleDispatchDriver
{
public:
	COCXEvents(){} // 调用 COleDispatchDriver 默认构造函数
	COCXEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COCXEvents(const COCXEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// OCXEvents 方法
public:
	void GotFocus()
	{
		InvokeHelper(0x800100e0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LostFocus()
	{
		InvokeHelper(0x800100e1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// OCXEvents 属性
public:

};
