#pragma once
#include <tuple>
#include <vector>
#include <exception>
#include <utility>
#include <Windows.h>
#pragma warning(push)
#pragma warning(disable: 4005) // Macro redefinition
#include <stdint.h>
#pragma warning(pop)
#include <type_traits>

namespace C6
{
  namespace internal
  {
    template <typename T1, typename T2>
    struct konst
    {
      typedef T1 T;
    };

    template <typename T>
    struct sizeof_or1
    {
      enum {size = sizeof(T)};
    };

    template <>
    struct sizeof_or1<void>
    {
      enum {size = 1};
    };

    template <>
    struct sizeof_or1<const void>
    {
      enum {size = 1};
    };

    template <typename T>
    struct bcount
    {
      enum {elem_size = sizeof_or1<std::conditional<std::is_array<T>::value,
        std::remove_all_extents<T>::type, std::remove_pointer<T>::type
        >::type>::size};
    };

    template <bool is_pod, typename T>
    struct ecount_base;

    template <typename T>
    struct ecount_base<false, T>
    {
      static auto size(T& x) -> decltype(x.size()) {return x.size();}
      static auto data(T& x) -> decltype(x.data()) {return x.data();}
    };

    template <typename T>
    struct ecount_base<true, T>
    {
      static size_t size(T x) {return 1;}
      static typename std::remove_reference<T>::type* data(T x) {return &x;}
    };

    template <typename T>
    struct ecount : ecount_base<std::is_pod<typename std::remove_reference<T>::type>::value, T>
    {
    };

    template <>
    struct ecount<nullptr_t>
    {
      static size_t size(nullptr_t) {return 0;}
      static nullptr_t data(nullptr_t) {return nullptr;}
    };

    template <>
    struct ecount<const wchar_t*>
    {
      static size_t size(const wchar_t* s) {return wcslen(s);}
      static const wchar_t* data(const wchar_t* s) {return s;}
    };

    template <typename T, size_t N>
    struct ecount<const T(&)[N]>
    {
      static size_t size(const T (&x)[N]) {return N;}
      static const T* data(const T (&x)[N]) {return x;}
    };

    template <typename T, size_t N>
    struct ecount<T(&)[N]>
    {
      static size_t size(T (&x)[N]) {return N;}
      static T* data(T (&x)[N]) {return x;}
    };
  }

  class COMException : public std::exception
  {
  public:
    explicit COMException(HRESULT hr, const char* function);
    ~COMException();
    bool operator == (HRESULT hr) const;
    
    const char* what() const;

  protected:
    HRESULT hr;
    const char* function;
    mutable char* msg;
  };

  class CreateFileException : public COMException
  {
  public:
    explicit CreateFileException(const char* filename);
    explicit CreateFileException(const wchar_t* filename);
    ~CreateFileException();
  };
  
  class COMObject
  {
  private:
    ::IUnknown* ptr;
    
    //! Explicit safe release of the underlying pointer.
    inline void release()
    {
      if(ptr != nullptr)
      {
        ptr->Release();
        ptr = nullptr;
      }
    }

    //! Get the underlying pointer and increment its reference count.
    inline IUnknown* addRef() const
    {
      if(ptr != nullptr)
        ptr->AddRef();
      return ptr;
    }
    
  public:
    //! Default constructor.
    COMObject() : ptr(nullptr) {}
    //! Nullary constructor.
    COMObject(nullptr_t) : ptr(nullptr) {}
    //! Explicit constructor from raw pointer.
    explicit COMObject(IUnknown* raw) : ptr(raw) {}
    //! Copy constructor.
    COMObject(const COMObject& copy_from) : ptr(copy_from.addRef()) {}
    //! Move constructor.
    COMObject(COMObject&& move_from) : ptr(move_from.ptr) {move_from.ptr = nullptr;}
    //! Destructor.
    ~COMObject() {release();}

    template <typename SmartInterface>
    SmartInterface queryInterface()
    {
      if(ptr == nullptr)
        return nullptr;
      decltype(SmartInterface().getRawInterface()) iface(nullptr);
      ptr->QueryInterface(&iface);
      return SmartInterface(iface);
    }
    
    //! Assignment operator.
    inline COMObject& operator= (const COMObject& copy_from) {COMObject copy(copy_from); swap(copy); return *this;}
    //! Move operator.
    inline COMObject& operator= (COMObject&& move_from) {swap(move_from); return *this;}
    //! Nullary assignment operator.
    inline COMObject& operator= (nullptr_t) {release(); return *this;}
    
    //! Get the underlying pointer.
    inline ::IUnknown* getRawInterface() {return ptr;}
    //! Get the underlying pointer.
    inline const ::IUnknown* getRawInterface() const {return ptr;}

    //! Efficient swap implementation
    inline void swap(COMObject& other) {std::swap(ptr, other.ptr);}
  };

  class COMStream : public COMObject
  {
  public:
    // Constructors
    COMStream() {}
    COMStream(nullptr_t) {}
    explicit COMStream(::IStream* raw) : COMObject(raw) {}
    COMStream(const COMStream& copy_from) : COMObject(copy_from) {}
    COMStream(COMStream&& move_from) : COMObject(std::move(move_from)) {}

    // Operators
    inline COMStream& operator= (const COMStream& copy_from) {COMObject::operator=(copy_from); return *this;}
    inline COMStream& operator= (COMStream&& move_from) {COMObject::operator=(std::move(move_from)); return *this;}
    _OPERATOR_BOOL() const {return getRawInterface() ? _CONVERTIBLE_TO_TRUE : 0;}

    // Utilities
    inline ::IStream* getRawInterface() {return static_cast<::IStream*>(COMObject::getRawInterface());}
    inline const ::IStream* getRawInterface() const {return static_cast<const ::IStream*>(COMObject::getRawInterface());}
    inline void swap(COMStream& other) {COMObject::swap(other);}

    // Methods
    template <typename T>
    unsigned long read(T&& destination)
    {
      unsigned long cb;
      auto data = C6::internal::ecount<T>::data(destination);
      cb = static_cast<unsigned long>(C6::internal::ecount<T>::size(destination) * C6::internal::bcount<decltype(data)>::elem_size);
      HRESULT hr = getRawInterface()->Read(std::move(data), cb, &cb);
      if(FAILED(hr)) throw COMException(hr, "IStream::Read");
      return cb;
    }

    template <typename T>
    unsigned long write(T&& source)
    {
      unsigned long cb;
      auto data = C6::internal::ecount<T>::data(source);
      cb = static_cast<unsigned long>(C6::internal::ecount<T>::size(source) * C6::internal::bcount<decltype(data)>::elem_size);
      HRESULT hr = getRawInterface()->Write(std::move(data), cb, &cb);
      if(FAILED(hr)) throw COMException(hr, "IStream::Write");
      return cb;
    }

    ULARGE_INTEGER seek(LARGE_INTEGER dlibMove, DWORD dwOrigin);
    void setSize(ULARGE_INTEGER libNewSize);
    std::tuple<ULARGE_INTEGER, ULARGE_INTEGER> copyTo(COMStream& pstm, ULARGE_INTEGER cb);
    void commit(DWORD grfCommitFlags);
    void revert();
    void lockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    void unlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
    STATSTG stat(DWORD grfStatFlag);
    COMStream clone();
  };
}