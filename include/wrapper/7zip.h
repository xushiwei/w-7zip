/* -------------------------------------------------------------------------
// WINX: a C++ template GUI library - MOST SIMPLE BUT EFFECTIVE
// 
// This file is a part of the WINX Library.
// The use and distribution terms for this software are covered by the
// Common Public License 1.0 (http://opensource.org/licenses/cpl.php)
// which can be found in the file CPL.txt at this distribution. By using
// this software in any fashion, you are agreeing to be bound by the terms
// of this license. You must not remove this notice, or any other, from
// this software.
// 
// Module: wrapper/7zip.h
// Creator: xushiwei
// Email: xushiweizh@gmail.com
// Date: 2009-7-2 23:57:09
// 
// $Id: 7zip.h,v 2009-7-2 23:57:09 xushiwei Exp $
// -----------------------------------------------------------------------*/
#ifndef WRAPPER_7ZIP_H
#define WRAPPER_7ZIP_H

#pragma warning(disable:4786)

// -------------------------------------------------------------------------

#ifndef __IARCHIVE_H
#include "7z/CPP/7zip/Archive/IArchive.h"
#endif

#ifndef assert
#include <assert.h>
#endif

#ifndef NS7ZIP_CALL
#define NS7ZIP_CALL
#endif

#ifndef NS7ZIP_ASSERT
#define NS7ZIP_ASSERT(e)	assert(e)
#endif

// -------------------------------------------------------------------------

namespace NS7zip {

//
// Create Object Instance
//
STDAPI CreateObject(const GUID* clsID, const GUID* interfaceID, void** outObject);

//
// Create InFileStream
//
STDAPI CreateInFileStream(LPCWSTR szFile, IInStream** ppstm);

//
// Create OutFileStream
//
STDAPI CreateOutFileStream(LPCWSTR szFile, IOutStream** ppstm);

//
// ToUInt64
//
inline UINT64 NS7ZIP_CALL ToUInt64(const PROPVARIANT& prop)
{
	switch (prop.vt)
	{
	case VT_UI1: return prop.bVal;
	case VT_UI2: return prop.uiVal;
	case VT_UI4: return prop.ulVal;
	case VT_UI8: return (UINT64)prop.uhVal.QuadPart;
	default:	 return 0;
	}
}

//
// Create InFileArchive
//
inline HRESULT NS7ZIP_CALL CreateInFileArchive(
	LPCWSTR szFile, const GUID* clsID, IArchiveOpenCallback* callback, IInArchive** pinArchive)
{
	struct __declspec(uuid("{23170F69-40C1-278A-0000-000600600000}")) IInArchive_;

    IInArchive* inArchive = NULL;
	HRESULT hr = CreateObject(clsID, &__uuidof(IInArchive_), (void **)&inArchive);
    if (hr != S_OK)
		return hr;
	
	IInStream* inStm = NULL;
	hr = CreateInFileStream(szFile, &inStm);
	if (hr != S_OK)
	{
		inArchive->Release();
		return hr;
	}
	
	hr = inArchive->Open(inStm, 0, callback);
	inStm->Release();
	if (hr != S_OK)
	{
		inArchive->Release();
		return hr;
	}
	*pinArchive = inArchive;
	return S_OK;
}

//
// class ArchiveOpenCallback
//
class ArchiveOpenCallback : public IArchiveOpenCallback
{
public:
	// IUnknown
	STDMETHOD(QueryInterface)(REFIID iid, void** ppv) { return E_NOINTERFACE; }
	STDMETHOD_(ULONG, AddRef)() { return 2; }
	STDMETHOD_(ULONG, Release)() { return 1; }
	
	// IArchiveOpenCallback
	STDMETHOD(SetTotal)(const UInt64 *files, const UInt64 *bytes) { return S_OK; }
	STDMETHOD(SetCompleted)(const UInt64 *files, const UInt64 *bytes) { return S_OK; }
	
	// ICryptoGetTextPassword
	STDMETHOD(CryptoGetTextPassword)(BSTR *aPassword) { return E_ABORT; }
};

//
// Create InFileArchive
//
inline HRESULT NS7ZIP_CALL CreateInFileArchive(
	LPCWSTR szFile, const GUID* clsID, IInArchive** pinArchive)
{
	ArchiveOpenCallback callback;
	return CreateInFileArchive(szFile, clsID, &callback, pinArchive);
}

//
// Is directory?
//
inline bool NS7ZIP_CALL IsDirectory(IInArchive* inArchive, UInt32 i)
{
	PROPVARIANT prop;
	PropVariantInit(&prop);
	inArchive->GetProperty(i, kpidIsDir, &prop);
	return prop.boolVal != VARIANT_FALSE;
}

//
// List Archive
//
template <class ArchiveListT>
inline void NS7ZIP_CALL ListArchiveFiles(IInArchive* inArchive, ArchiveListT& ns)
{
	typedef typename ArchiveListT::value_type ItemT;

	UINT32 numItems = 0;
	inArchive->GetNumberOfItems(&numItems);
	for (UINT32 i = 0; i < numItems; i++)
	{
		PROPVARIANT prop;
		PropVariantInit(&prop);

		// Is directory?
		inArchive->GetProperty(i, kpidIsDir, &prop);
		if (prop.boolVal != VARIANT_FALSE)
			continue;

		// Get name of file
		if (inArchive->GetProperty(i, kpidPath, &prop) == S_OK)
		{
			ns.insert(ItemT(prop.bstrVal, i));
			PropVariantClear(&prop);
		}
	}
}

//
// class ArchiveExtractCallback
//
class ArchiveExtractCallback : public IArchiveExtractCallback
{
private:
	UInt32 m_index;
	ISequentialOutStream* m_outStream;

public:
	ArchiveExtractCallback(UInt32 index, ISequentialOutStream* outStream)
		: m_index(index), m_outStream(outStream) {
	}

	// IUnknown
	STDMETHOD(QueryInterface)(REFIID iid, void** ppv) { return E_NOINTERFACE; }
	STDMETHOD_(ULONG, AddRef)() { return 2; }
	STDMETHOD_(ULONG, Release)() { return 1; }
	
	// IProgress
	STDMETHOD(SetTotal)(UInt64 size) { return S_OK; }
	STDMETHOD(SetCompleted)(const UInt64 *completeValue) { return S_OK; }
	
	// ICryptoGetTextPassword
	STDMETHOD(CryptoGetTextPassword)(BSTR *aPassword) { return E_ABORT; }
	
	// IArchiveExtractCallback
	STDMETHOD(GetStream)(UInt32 index, ISequentialOutStream** outStream, Int32 askExtractMode)
	{
		if (askExtractMode != NArchive::NExtract::NAskMode::kExtract) {
			*outStream = NULL;
			return S_OK;
		}

		NS7ZIP_ASSERT(index == m_index);
		*outStream = m_outStream;
		m_outStream->AddRef();
		return S_OK;
	}
	STDMETHOD(PrepareOperation)(Int32 askExtractMode) { return S_OK; }
	STDMETHOD(SetOperationResult)(Int32 resultEOperationResult) { return S_OK; }
};

inline HRESULT NS7ZIP_CALL ExtractArchiveFile(IInArchive* inArchive, UInt32 index, ISequentialOutStream* outStream)
{
	NS7ZIP_ASSERT(!IsDirectory(inArchive, index));
	ArchiveExtractCallback callback(index, outStream);
	return inArchive->Extract(&index, 1, NArchive::NExtract::NAskMode::kExtract, &callback);
}

inline HRESULT NS7ZIP_CALL ExtractArchiveFile(IInArchive* inArchive, UInt32 index, LPCWSTR szFile)
{
	IOutStream* pstm;
	HRESULT hr = CreateOutFileStream(szFile, &pstm);
	if (hr != S_OK)
		return hr;
	hr = ExtractArchiveFile(inArchive, index, pstm);
	pstm->Release();
	return hr;
}

//
// class LimitedMemoryStream
//
class LimitedMemoryStream : public ISequentialOutStream
{
private:
	char* m_buf;
	char* m_bufEnd;

public:
	LimitedMemoryStream(char* buf, UInt32 cb)
		: m_buf(buf), m_bufEnd(buf + cb) {
	}

	// IUnknown
	STDMETHOD(QueryInterface)(REFIID iid, void** ppv) { return E_NOINTERFACE; }
	STDMETHOD_(ULONG, AddRef)() { return 2; }
	STDMETHOD_(ULONG, Release)() { return 1; }

	// ISequentialOutStream
	STDMETHOD(Write)(const void *data, UInt32 size, UInt32 *processedSize)
	{
		if (size <= (UInt32)(m_bufEnd - m_buf)) {
			memcpy(m_buf, data, size);
			*processedSize = size;
			m_buf += size;
		}
		else {
			if (m_buf >= m_bufEnd)
				return E_ACCESSDENIED;
			memcpy(m_buf, data, m_bufEnd - m_buf);
			*processedSize = m_bufEnd - m_buf;
			m_buf = m_bufEnd;
		}
		return S_OK;
	}

	const char* end() const {
		return m_buf;
	}

	bool done() const {
		return m_buf == m_bufEnd;
	}
};

//
// class InArchive
//
class InArchive
{
private:
	IInArchive* inArchive;

public:
	InArchive() : inArchive(NULL) {}
	InArchive(LPCWSTR szFile, const GUID* clsID, IArchiveOpenCallback* callback)
	{
		inArchive = NULL;
		CreateInFileArchive(szFile, clsID, callback, &inArchive);
	}

	InArchive(LPCWSTR szFile, const GUID* clsID)
	{
		inArchive = NULL;
		ArchiveOpenCallback callback;
		CreateInFileArchive(szFile, clsID, &callback, &inArchive);
	}

	~InArchive()
	{
		if (inArchive)
			inArchive->Release();
	}

	bool NS7ZIP_CALL good() const {
		return inArchive != NULL;
	}

	HRESULT NS7ZIP_CALL open(LPCWSTR szFile, const GUID* clsID, IArchiveOpenCallback* callback) {
		NS7ZIP_ASSERT(inArchive == NULL);
		return CreateInFileArchive(szFile, clsID, callback, &inArchive);
	}

	HRESULT NS7ZIP_CALL open(LPCWSTR szFile, const GUID* clsID) {
		NS7ZIP_ASSERT(inArchive == NULL);
		ArchiveOpenCallback callback;
		return CreateInFileArchive(szFile, clsID, &callback, &inArchive);
	}

	void NS7ZIP_CALL close() {
		if (inArchive) {
			inArchive->Release();
			inArchive = NULL;
		}
	}

	template <class ArchiveListT>
	void NS7ZIP_CALL ListFiles(ArchiveListT& ns) const {
		ListArchiveFiles(inArchive, ns);
	}

	HRESULT NS7ZIP_CALL ExtractFile(UInt32 index, ISequentialOutStream* outStream) const {
		return ExtractArchiveFile(inArchive, index, outStream);
	}

	HRESULT NS7ZIP_CALL ExtractFile(UInt32 index, LPCWSTR szDestFile) const {
		return ExtractArchiveFile(inArchive, index, szDestFile);
	}

	template <class ArchiveListT>
	HRESULT NS7ZIP_CALL ExtractFile(const ArchiveListT& ns, LPCWSTR szSrcPath, ISequentialOutStream* outStream) const {
		ArchiveListT::const_iterator it = ns.find(szSrcPath);
		if (it == ns.end())
			return E_INVALIDARG;
		return ExtractArchiveFile(inArchive, (*it).second, outStream);
	}

	template <class ArchiveListT>
	HRESULT NS7ZIP_CALL ExtractFile(const ArchiveListT& ns, LPCWSTR szSrcPath, LPCWSTR szDestFile) const {
		ArchiveListT::const_iterator it = ns.find(szSrcPath);
		if (it == ns.end())
			return E_INVALIDARG;
		return ExtractArchiveFile(inArchive, (*it).second, szDestFile);
	}

	template <class ArchiveListT, class OutBufferT> // OutBufferT = std::vector<char>
	HRESULT NS7ZIP_CALL ExtractFile(const ArchiveListT& ns, LPCWSTR szSrcPath, OutBufferT& outBuffer) const
	{
		ArchiveListT::const_iterator it = ns.find(szSrcPath);
		if (it == ns.end())
			return E_INVALIDARG;
		const UInt32 index = (*it).second;
		PROPVARIANT prop;
		PropVariantClear(&prop);
		inArchive->GetProperty(index, kpidSize, &prop);
		const UInt32 size = (UInt32)ToUInt64(prop);
		outBuffer.resize(size);
		if (size == 0)
			return S_OK;
		LimitedMemoryStream stm(&outBuffer[0], size);
		return ExtractArchiveFile(inArchive, index, &stm);
	}
};

} // namespace

// -------------------------------------------------------------------------
// Link 7z.lib

#if !defined(Wrapper_Linked_7z)
#define Wrapper_Linked_7z
#pragma comment(lib, "7z")
#endif

// -------------------------------------------------------------------------

#endif /* WRAPPER_7ZIP_H */
