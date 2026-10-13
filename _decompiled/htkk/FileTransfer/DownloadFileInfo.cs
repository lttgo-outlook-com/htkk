using System;
using Foundation.Base;
using Foundation.Utility;

namespace FileTransfer;

public class DownloadFileInfo : BaseObject
{
	public const int EN_NULL_STR = -1;

	public const int EN_NUM_FIELDS_NOT_MATCH = -2;

	public const int EN_SPLIT_FAILED = -3;

	private int file_id;

	private CacheFileInfo cache_file_info;

	private string dir_path_in_client;

	private string file_name;

	private bool is_read_only;

	private byte[] file_signature;

	private DownloadFileInfo(int file_id, CacheFileInfo cache_file_info, string dir_file_path, string file_name, bool is_read_only)
	{
		if (cache_file_info == null || dir_file_path == null || file_name == null)
		{
			SetNull();
			return;
		}
		this.file_id = file_id;
		this.cache_file_info = new CacheFileInfo(cache_file_info);
		dir_path_in_client = dir_file_path;
		this.file_name = file_name;
		this.is_read_only = is_read_only;
	}

	public DownloadFileInfo(DownloadFileInfo df_info)
	{
		if (df_info == null)
		{
			SetNull();
			return;
		}
		file_id = df_info.file_id;
		dir_path_in_client = df_info.dir_path_in_client;
		file_name = df_info.file_name;
		cache_file_info = df_info.cache_file_info;
		file_signature = ArrayUtil.Copy(df_info.file_signature);
		is_read_only = df_info.is_read_only;
	}

	public long GetFileSize()
	{
		if (cache_file_info == null)
		{
			return -1L;
		}
		return cache_file_info.GetFileSize();
	}

	public int GetFileSizeInKB()
	{
		if (cache_file_info == null)
		{
			return -1;
		}
		long fileSize = cache_file_info.GetFileSize();
		int result;
		try
		{
			result = (int)(fileSize / 1024);
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public int GetNumBlocks()
	{
		try
		{
			return cache_file_info.GetNumBlocks();
		}
		catch
		{
			return -1;
		}
	}

	public int GetBlockLen()
	{
		try
		{
			return cache_file_info.GetBlockLen();
		}
		catch
		{
			return -1;
		}
	}

	public void SetFileSignature(byte[] bytes)
	{
		file_signature = ArrayUtil.Copy(bytes);
	}

	public byte[] GetFileSignature()
	{
		return file_signature;
	}

	public int GetFileId()
	{
		return file_id;
	}

	public string GetDirPath()
	{
		return dir_path_in_client;
	}

	public string GetFileName()
	{
		return file_name;
	}

	public bool IsReadOnly()
	{
		return is_read_only;
	}

	public string GetFilePath(string root_prog_folder)
	{
		if (root_prog_folder == null || dir_path_in_client == null || file_name == null)
		{
			return null;
		}
		return (dir_path_in_client[dir_path_in_client.Length - 1] != '\\') ? (root_prog_folder + dir_path_in_client + "\\" + file_name) : (root_prog_folder + dir_path_in_client + file_name);
	}

	public static DownloadFileInfo[] LoadListFromStr(string str, char info_delimiter, char field_delimiver, out int errno)
	{
		if (str == null)
		{
			errno = -1;
			return null;
		}
		string[] array = str.Split(new char[1] { info_delimiter });
		DownloadFileInfo[] array2 = new DownloadFileInfo[array.Length];
		bool flag = false;
		errno = 0;
		for (int i = 0; i < array2.Length; i++)
		{
			array2[i] = LoadFromStr(array[i], field_delimiver, out errno);
			if (array2[i] == null)
			{
				flag = true;
				break;
			}
		}
		if (flag)
		{
			return null;
		}
		return array2;
	}

	public static DownloadFileInfo LoadFromStr(string str, char delimiter, out int errno)
	{
		if (str == null)
		{
			errno = -1;
			return null;
		}
		string[] array;
		try
		{
			array = str.Split(new char[1] { delimiter });
			if (array.Length < 5)
			{
				errno = -2;
				array = null;
			}
		}
		catch
		{
			array = null;
		}
		if (array == null)
		{
			errno = -3;
			return null;
		}
		DownloadFileInfo downloadFileInfo = null;
		try
		{
			string text = array[^1];
			string text2 = array[^2];
			string text3 = array[^3];
			string text4 = array[^4];
			int num = Convert.ToInt32(text4);
			bool flag;
			if (num < 0)
			{
				downloadFileInfo = null;
			}
			else if (ConvertReadOnly(text, out flag) != 0)
			{
				downloadFileInfo = null;
			}
			else
			{
				int length = str.Length - text4.Length - text3.Length - text2.Length - text.Length - 4;
				string str2 = str.Substring(0, length);
				int num2;
				CacheFileInfo cacheFileInfo = CacheFileInfo.LoadFromString(str2, delimiter, out num2);
				downloadFileInfo = ((cacheFileInfo != null) ? new DownloadFileInfo(num, cacheFileInfo, text3, text2, flag) : null);
			}
		}
		catch
		{
			downloadFileInfo = null;
		}
		if (downloadFileInfo == null)
		{
			errno = -10001;
		}
		errno = 0;
		return downloadFileInfo;
	}

	private static int ConvertReadOnly(string str, out bool is_read_only)
	{
		int result = 0;
		if (str == "1" || str == "0")
		{
			if (str == "1")
			{
				is_read_only = true;
			}
			else
			{
				is_read_only = false;
			}
		}
		else
		{
			result = -1;
			is_read_only = false;
		}
		return result;
	}

	private void SetNull()
	{
		cache_file_info = null;
		dir_path_in_client = null;
		file_name = null;
		is_null = true;
	}
}
