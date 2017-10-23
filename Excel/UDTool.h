#include <vector>
#include<string>
using namespace std;

#pragma once
namespace UDExcel
{
  class a2w 
  {
    wchar_t  local[16];
    wchar_t* buffer;
    unsigned int nu;
    void init(const char* str, unsigned int n)
	{
		nu = MultiByteToWideChar(CP_THREAD_ACP,0,str,n,0,0);
		buffer = ( nu < (16-1) )? local: new wchar_t[nu+1];
        MultiByteToWideChar(CP_ACP,0,str,n,buffer,nu);
        buffer[nu] = 0;
	}
  public:
	  a2w(const char* str):buffer(0), nu(0)
      {
		  if(str)
			  init(str, (unsigned int)strlen(str) );
	  }

	  ~a2w() {  if(buffer != local) delete[] buffer;  }
	  unsigned int length() const { return nu; }
	  operator const wchar_t*() { return buffer; }
  };

  class w2a 
  {
	  char  local[16];
	  char* buffer;
	  unsigned int n;
	  
	  void init(const wchar_t* wstr, unsigned int nu)
      {
		  n = WideCharToMultiByte(CP_ACP,0,wstr,nu,0,0,0,0);
		  buffer = (n < (16-1))? local:new char[n+1];
		  WideCharToMultiByte(CP_ACP,0,wstr,nu,buffer,n,0,0);
		  buffer[n] = 0;
      }
  public:
	  w2a(const wchar_t* wstr):buffer(0),n(0)
      {
		  if(wstr)
			  init(wstr,(unsigned int)wcslen(wstr));
	  }
	  w2a(const std::wstring& wstr):buffer(0),n(0)
	  { 
		  init(wstr.c_str(),(unsigned int)wstr.length());
	  }
	  
	  ~w2a() { if(buffer != local) delete[] buffer;  }
	  unsigned int length() const { return n; }    
	  operator const char*() { return buffer; }
  };

  class  CI2A 
  {
	  char buffer[38];
  public:
	  CI2A(int n, int radix = 10)
	  {
		  _itoa(n, buffer, radix);
	  }
	  operator const char*()
	  {
		  return buffer; 
	  }
  };
  
  class   CF2A 
  {
	  char buffer[64];
  public: 
	  CF2A( double d, const char* units="", int fractional_digits=5 )
	  { 
		  _snprintf(buffer, 64, "%.*f%s", fractional_digits, d, units );
		  buffer[63] = 0;
	  }
	  
	  operator const char*()
	  { return buffer; }
  };

  class   CG2A 
  {
	  char buffer[64];
  public: 
	  CG2A( double d)
	  { 
		  _snprintf(buffer, 64, "%g", d );
		  buffer[63] = 0;
	  }
	  
	  operator const char*()
	  { return buffer; }
  };


  class CExcelExpection
  {
  public:
	  CExcelExpection(string sMsg){ m_sMsg = sMsg; }
	  string GetMsg(){return m_sMsg;}
  private:
	  string m_sMsg;
  };

}
