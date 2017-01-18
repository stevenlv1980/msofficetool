// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDocumentEvents 包装类

class CDocumentEvents : public COleDispatchDriver
{
public:
	CDocumentEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CDocumentEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDocumentEvents(const CDocumentEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// DocumentEvents 方法
public:
	void New()
	{
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Open()
	{
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Close()
	{
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// DocumentEvents 属性
public:

};
