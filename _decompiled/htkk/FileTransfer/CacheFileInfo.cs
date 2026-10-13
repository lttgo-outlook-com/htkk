using System;
using Foundation.Base;

namespace FileTransfer;

public class CacheFileInfo : BaseObject
{
	public const int EN_NULL_STR = -1;

	public const int EN_NUM_FIELDS_NOT_MATCH = -2;

	private int block_len;

	private int num_blocks;

	private long file_size;

	public CacheFileInfo()
	{
		block_len = 0;
		num_blocks = 0;
		file_size = 0L;
		is_null = false;
	}

	public CacheFileInfo(int block_len, int num_blocks, long file_size)
	{
		this.block_len = block_len;
		this.num_blocks = num_blocks;
		this.file_size = file_size;
		is_null = false;
	}

	public CacheFileInfo(CacheFileInfo cache_info)
	{
		block_len = cache_info.block_len;
		num_blocks = cache_info.num_blocks;
		file_size = cache_info.file_size;
	}

	public string SerializeToString(char delimiter)
	{
		return block_len.ToString() + delimiter + num_blocks;
	}

	public long GetFileSize()
	{
		return file_size;
	}

	public int GetNumBlocks()
	{
		return num_blocks;
	}

	public int GetBlockLen()
	{
		return block_len;
	}

	public static CacheFileInfo LoadFromString(string str, char delimiter, out int errno)
	{
		if (str == null)
		{
			errno = -1;
			return null;
		}
		string[] array = str.Split(new char[1] { delimiter });
		if (array.Length != 3)
		{
			errno = -2;
			return null;
		}
		CacheFileInfo cacheFileInfo = null;
		try
		{
			int num = Convert.ToInt32(array[0]);
			int num2 = Convert.ToInt32(array[1]);
			long num3 = Convert.ToInt64(array[2]);
			cacheFileInfo = new CacheFileInfo(num, num2, num3);
		}
		catch
		{
			cacheFileInfo = null;
		}
		errno = 0;
		return cacheFileInfo;
	}
}
