#include <nice/com.h>
#include <Windows.h>

namespace C6
{
  COMException::COMException(HRESULT hr_, const char* function_)
    : hr(hr_)
    , function(function_)
    , msg(nullptr)
  {
  }

  COMException::~COMException()
  {
    LocalFree(msg);
  }

  bool COMException::operator == (HRESULT hr_) const
  {
    return hr == hr_;
  }

  static char* FormatMessagef(const char* format, ...)
  {
    LPSTR msg = nullptr;
    va_list argptr;
    va_start(argptr, format);
    FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_STRING, format, 0, 0, reinterpret_cast<LPSTR>(&msg), 0, &argptr);
    va_end(argptr);
    return msg;
  }

  const char* COMException::what() const
  {
    if(!msg)
    {
      LPSTR readable_hr = nullptr;
      FormatMessageA(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr, hr, 0, reinterpret_cast<LPSTR>(&readable_hr), 0, nullptr);
      if(readable_hr)
      {
        msg = FormatMessagef("%1!s! (%2!s!)", readable_hr, function);
        LocalFree(readable_hr);
      }
      else
      {
        msg = FormatMessagef("HR 0x%1!08x! (%2!s!)", hr, function);
      }
    }
    return msg ? msg : "COMException";
  }

  CreateFileException::CreateFileException(const char* filename)
    : COMException(HRESULT_FROM_WIN32(GetLastError()), "CreateFileA")
  {
    char* full_msg = new char[strlen(function) + strlen(filename) + 3];
    sprintf(full_msg, "%s: %s", function, filename);
    function = full_msg;
  }

  CreateFileException::CreateFileException(const wchar_t* filename)
    : COMException(HRESULT_FROM_WIN32(GetLastError()), "CreateFileW")
  {
    char* full_msg = new char[strlen(function) + wcslen(filename) * 3 + 3];
    sprintf(full_msg, "%s: %S", function, filename);
    function = full_msg;
  }

  CreateFileException::~CreateFileException()
  {
    delete[] function;
  }

  ULARGE_INTEGER COMStream::seek(LARGE_INTEGER dlibMove, DWORD dwOrigin)
  {
    ULARGE_INTEGER libNewPosition;
    HRESULT hr = getRawInterface()->Seek(dlibMove, dwOrigin, &libNewPosition);
    if(FAILED(hr)) throw COMException(hr, "IStream::Seek");
    return libNewPosition;
  }

  void COMStream::setSize(ULARGE_INTEGER libNewSize)
  {
    HRESULT hr = getRawInterface()->SetSize(libNewSize);
    if(FAILED(hr)) throw COMException(hr, "IStream::SetSize");
  }

  std::tuple<ULARGE_INTEGER, ULARGE_INTEGER> COMStream::copyTo(COMStream& pstm, ULARGE_INTEGER cb)
  {
    ULARGE_INTEGER cbRead;
    ULARGE_INTEGER cbWritten;
    HRESULT hr = getRawInterface()->CopyTo(pstm.getRawInterface(), cb, &cbRead, &cbWritten);
    if(FAILED(hr)) throw COMException(hr, "IStream::CopyTo");
    return std::make_tuple(cbRead, cbWritten);
  }

  void COMStream::commit(DWORD grfCommitFlags)
  {
    HRESULT hr = getRawInterface()->Commit(grfCommitFlags);
    if(FAILED(hr)) throw COMException(hr, "IStream::Commit");
  }

  void COMStream::revert()
  {
    HRESULT hr = getRawInterface()->Revert();
    if(FAILED(hr)) throw COMException(hr, "IStream::Revert");
  }

  void COMStream::lockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
  {
    HRESULT hr = getRawInterface()->LockRegion(libOffset, cb, dwLockType);
    if(FAILED(hr)) throw COMException(hr, "IStream::LockRegion");
  }

  void COMStream::unlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
  {
    HRESULT hr = getRawInterface()->UnlockRegion(libOffset, cb, dwLockType);
    if(FAILED(hr)) throw COMException(hr, "IStream::UnlockRegion");
  }

  STATSTG COMStream::stat(DWORD grfStatFlag)
  {
    STATSTG stg;
    HRESULT hr = getRawInterface()->Stat(&stg, grfStatFlag);
    if(FAILED(hr)) throw COMException(hr, "IStream::Stat");
    return stg;
  }

  COMStream COMStream::clone()
  {
    IStream* stream;
    HRESULT hr = getRawInterface()->Clone(&stream);
    if(FAILED(hr)) throw COMException(hr, "IStream::Clone");
    return COMStream(stream);
  }
}
