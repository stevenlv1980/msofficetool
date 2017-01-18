// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDiagramNode 包装类

class CDiagramNode : public COleDispatchDriver
{
public:
	CDiagramNode(){} // 调用 COleDispatchDriver 默认构造函数
	CDiagramNode(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDiagramNode(const CDiagramNode& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// DiagramNode 方法
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
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Children()
	{
		LPDISPATCH result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Shape()
	{
		LPDISPATCH result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Root()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Diagram()
	{
		LPDISPATCH result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Layout()
	{
		long result;
		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Layout(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x69, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_TextShape()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH AddNode(long Pos, long NodeType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Pos, NodeType);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void MoveNode(LPDISPATCH * TargetNode, long Pos)
	{
		static BYTE parms[] = VTS_PDISPATCH VTS_I4 ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, TargetNode, Pos);
	}
	void ReplaceNode(LPDISPATCH * TargetNode)
	{
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, TargetNode);
	}
	void SwapNode(LPDISPATCH * TargetNode, long Pos)
	{
		static BYTE parms[] = VTS_PDISPATCH VTS_I4 ;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, TargetNode, Pos);
	}
	LPDISPATCH CloneNode(BOOL copyChildren, LPDISPATCH * TargetNode, long Pos)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BOOL VTS_PDISPATCH VTS_I4 ;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, copyChildren, TargetNode, Pos);
		return result;
	}
	void TransferChildren(LPDISPATCH * ReceivingNode)
	{
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ReceivingNode);
	}
	LPDISPATCH NextNode()
	{
		LPDISPATCH result;
		InvokeHelper(0x11, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH PrevNode()
	{
		LPDISPATCH result;
		InvokeHelper(0x12, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// DiagramNode 属性
public:

};
