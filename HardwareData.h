
class HardwareData
{
public:
	CString	m_partname;		//零件名称
	CString	m_partno;		//零件代号
	CString	m_styleno;		//规格
	double	m_price;		//价格
	int		m_from;			//0:国产，1-进口   
	CString	m_factory;		//生产厂家
	CString	m_comment;		//备注 
	CString	m_dwgname;		//此五金件所使用的模型名称
public:
	HardwareData()
	{
		m_price = 0.0;
		m_from = 0;
		m_partname = _T("");
		m_partno = _T("");
		m_styleno = _T("");
		m_factory = _T("");
		m_comment = _T("");
		m_dwgname = _T("");
	}
	~HardwareData()
	{
	}
	HardwareData(const HardwareData &orig)
	{
		Copy(orig);
	}
	HardwareData &operator=(const HardwareData &orig)
	{
		Copy(orig);
		return (*this);
	}

protected:
	void Copy(const HardwareData &orig)
	{
		m_price = orig.m_price;
		m_from = orig.m_from;
		m_partname = orig.m_partname;
		m_partno = orig.m_partno;
		m_styleno = orig.m_styleno;
		m_factory = orig.m_factory;
		m_comment = orig.m_comment;
		m_dwgname = orig.m_dwgname;
	}
};

class BiasConnecterData:public HardwareData
{
public:
	double	m_d1;
	double	m_depth1;
	double	m_d2;
	double	m_depth2;
	double	m_d3;
	double	m_depth3;
	double	m_depth4;
	double	m_offset;	
public:
	BiasConnecterData()
	{
		m_d1 = 0.0;
		m_depth1 = 0.0;
		m_d2 = 0.0;
		m_depth2 = 0.0;
		m_d3 = 0.0;
		m_depth3 = 0.0;
		m_depth4 = 0.0;
		m_offset = 0.0;	
	}
	~BiasConnecterData(){};
};