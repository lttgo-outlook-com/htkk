using System;
using System.IO;
using Foundation.Base;
using Foundation.Utility;

namespace FileTransfer;

internal class CurrVersionMgt : BaseObject
{
	private const string VERSION_CODE_FNAME = "version.inf";

	private const string DOWNLOAD_VERSION_RESULT_FNAME = "download_rslt.inf";

	public const int EN_ACCESS_DENIED = -1;

	public const int EN_PATH_TOO_LONG = -2;

	public const int EN_FILE_NOT_FOUND = -3;

	public const int EN_DIR_NOT_FOUND = -4;

	public const int EN_DRIVE_NOT_FOUND = -5;

	public const int EN_MAKE_DIR_FAILED = -6;

	private const int MAX_FOLDER_LEVEL = 20;

	private string root_prog_folder;

	private string version_folder;

	public CurrVersionMgt(string root_prog_folder, string version_folder)
	{
		if (root_prog_folder == null || version_folder == null)
		{
			is_null = true;
			return;
		}
		is_null = false;
		this.root_prog_folder = root_prog_folder;
		this.version_folder = version_folder;
	}

	public int CheckLastestVersion(string clientGuid, string tax_code, string curr_version, string tenCQTCapCuc, string maCQTCapCuc, string tenCQTNoiNop, string maCQTNoiNop, string admin_server_uri, out string lastest_version_code, out DownloadFileInfo[] down_file_info_list)
	{
		lastest_version_code = null;
		down_file_info_list = null;
		if (admin_server_uri == null)
		{
			return -1;
		}
		CheckLastestVersionClient checkLastestVersionClient = new CheckLastestVersionClient(admin_server_uri);
		string public_key;
		return checkLastestVersionClient.DoCheck(clientGuid, tax_code, curr_version, tenCQTCapCuc, maCQTCapCuc, tenCQTNoiNop, maCQTNoiNop, out lastest_version_code, out down_file_info_list, out public_key);
	}

	public string GetCurrVersion()
	{
		if (version_folder == null)
		{
			return null;
		}
		if (FileUtil.LoadFileContent(version_folder + "\\version.inf", out string file_content) != 0)
		{
			return null;
		}
		return file_content;
	}

	public DownloadFileInfo[] CompareCurrWithLastest(DownloadFileInfo[] down_file_info_list, string public_key)
	{
		if (down_file_info_list == null)
		{
			return null;
		}
		DownloadFileInfo[] array = new DownloadFileInfo[down_file_info_list.Length];
		bool[] array2 = new bool[down_file_info_list.Length];
		int num = 0;
		bool flag = false;
		for (int i = 0; i < down_file_info_list.Length; i++)
		{
			if (down_file_info_list[i] != null)
			{
				if (CheckFileExisted(down_file_info_list[i], public_key, out var existed) != 0)
				{
					flag = true;
				}
				else if (!existed)
				{
					array[num] = new DownloadFileInfo(down_file_info_list[i]);
					num++;
				}
				continue;
			}
			flag = true;
			break;
		}
		if (flag)
		{
			return null;
		}
		if (num == 0)
		{
			return null;
		}
		DownloadFileInfo[] array3 = new DownloadFileInfo[num];
		for (int i = 0; i < num; i++)
		{
			array3[i] = new DownloadFileInfo(array[i]);
		}
		return array3;
	}

	private int CheckFileExisted(DownloadFileInfo file_info, string public_key, out bool existed)
	{
		if (file_info == null)
		{
			existed = false;
			return -1;
		}
		string filePath = file_info.GetFilePath(root_prog_folder);
		if (filePath == null)
		{
			existed = false;
			return -1;
		}
		if (FileUtil.CheckExist(filePath, out var exist) != 0)
		{
			existed = false;
			return -1;
		}
		if (exist)
		{
			byte[] fileSignature = file_info.GetFileSignature();
			if (fileSignature == null)
			{
				existed = false;
			}
			else if (RSAUtil.VerifyFileSignature(filePath, fileSignature, public_key))
			{
				existed = true;
			}
			else
			{
				existed = false;
			}
		}
		else
		{
			existed = false;
		}
		return 0;
	}

	public int GetPrevDownloadResult(out bool is_ok, out string downloading_version_code)
	{
		is_ok = true;
		downloading_version_code = null;
		if (version_folder == null)
		{
			return -1;
		}
		string text = version_folder + "\\download_rslt.inf";
		if (FileUtil.CheckExist(text, out var exist) != 0)
		{
			return -1;
		}
		int result = 0;
		if (exist)
		{
			if (FileUtil.LoadFileContent(text, out string file_content) != 0)
			{
				result = -1;
			}
			else
			{
				downloading_version_code = file_content;
			}
			is_ok = false;
		}
		else
		{
			is_ok = true;
		}
		return result;
	}

	public int SetDownloadResult(string downloading_version_code, bool failed)
	{
		if (version_folder == null || downloading_version_code == null)
		{
			return -1;
		}
		string text = version_folder + "\\download_rslt.inf";
		int result = 0;
		if (failed)
		{
			if (FileUtil.SaveFileContent(text, downloading_version_code) != 0)
			{
				result = -1;
			}
		}
		else
		{
			FileUtil.RemoveFile(text);
		}
		return result;
	}

	public int UpdateVersion(string[] temp_file_paths, string version_code, DownloadFileInfo[] down_file_info_list)
	{
		if (down_file_info_list == null)
		{
			errno = -10001;
			return -1;
		}
		if (temp_file_paths == null)
		{
			errno = -10001;
			return -1;
		}
		if (temp_file_paths.Length != down_file_info_list.Length)
		{
			errno = -10001;
			return -1;
		}
		int num = 0;
		for (int i = 0; i < temp_file_paths.Length; i++)
		{
			if (down_file_info_list[i] == null)
			{
				errno = -10001;
				num = -1;
				break;
			}
			string filePath = down_file_info_list[i].GetFilePath(root_prog_folder);
			bool flag = down_file_info_list[i].IsReadOnly();
			if (MakeFolderIfNotExist(filePath) != 0)
			{
				errno = -6;
				return -1;
			}
			if (FileUtil.CopyOverwriteIfExist(temp_file_paths[i], filePath, out var err_no) != 0)
			{
				switch (err_no)
				{
				case -1:
					errno = -1;
					break;
				case -2:
					errno = -3;
					break;
				case -5:
					errno = -2;
					break;
				case -3:
					errno = -4;
					break;
				case -4:
					errno = -5;
					break;
				default:
					errno = -10001;
					break;
				}
				num = -1;
				break;
			}
			if (flag)
			{
				FileUtil.SetReadOnly(filePath, is_read_only: true);
			}
		}
		if (num != 0)
		{
			return num;
		}
		if (SetCurrVersion(version_code) != 0)
		{
			num = -1;
		}
		else
		{
			for (int i = 0; i < temp_file_paths.Length; i++)
			{
				FileUtil.RemoveFile(temp_file_paths[i]);
			}
		}
		return num;
	}

	private int MakeFolderIfNotExist(string dest_path)
	{
		int num = 0;
		string[] array = new string[20];
		int num2 = 0;
		try
		{
			string text = dest_path;
			int num3 = 0;
			while (text != null && text != root_prog_folder && num3 < 20)
			{
				text = Path.GetDirectoryName(text);
				if (text != root_prog_folder && text != null)
				{
					array[num3] = text;
					num3++;
				}
			}
			num2 = num3;
		}
		catch (Exception)
		{
			num = -1;
		}
		if (num != 0)
		{
			return -1;
		}
		bool flag = true;
		for (int num4 = num2 - 1; num4 >= 0; num4--)
		{
			if (FileUtil.CheckDirExist(array[num4], out var exist) == 0 && !exist && FileUtil.CreateDir(array[num4]) != 0)
			{
				flag = false;
				break;
			}
		}
		return (!flag) ? (-1) : 0;
	}

	private int SetCurrVersion(string version_code)
	{
		if (version_folder == null)
		{
			errno = -10001;
			return -1;
		}
		int result = 0;
		if (FileUtil.SaveFileContent(version_folder + "\\version.inf", version_code, out var err_no) != 0)
		{
			switch (err_no)
			{
			case -1:
				errno = -1;
				break;
			case -2:
				errno = -3;
				break;
			case -5:
				errno = -2;
				break;
			default:
				errno = -10001;
				break;
			}
			result = -1;
		}
		return result;
	}
}
