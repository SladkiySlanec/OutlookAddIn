#pragma once

template <class T> class SmartPtr
{
public:
	SmartPtr(T * p)
	{
		m_pVal = p;
	}

	~SmartPtr() 
	{ 
		if(m_pVal)
		{
			delete[] m_pVal; 
			m_pVal = NULL;
		}
	}
	operator T*() const throw()
	{
		return m_pVal;
	}
	T * Val() 
	{
		return m_pVal; 
	}
	T * operator ->() 
	{ 
		return m_pVal; 
	}

protected:
	T * m_pVal;
};