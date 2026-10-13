using System;
using System.IO;
using Foundation.Base;
using Foundation.Utility;

namespace FileTransfer;

internal class BlockBuffer : BaseObject
{
	private const char FILE_BLOCK_DELIMITER = '_';

	private const string BLOCK_FOLDER = "BLOCK";

	private const string FILE_FOLDER = "FILE";

	private const string BLOCK_FILE_EXT = "blk";

	private const string FULL_FILE_EXT = "prg";

	private const string BLOCK_DOWNLOADING_EXT = "dwl";

	private const int MAX_FILES_NUM = 2000;

	private const int MAX_BLOCKS_NUM = 1000;

	private string version_code;

	private string temp_root_folder;

	private string temp_block_folder;

	private string temp_file_folder;

	public BlockBuffer(string temp_root_folder, string version_code)
	{
		this.temp_root_folder = temp_root_folder;
		this.version_code = version_code;
	}

	public int Open()
	{
		if (version_code == null || version_code == "" || temp_root_folder == null || temp_root_folder == "")
		{
			return -1;
		}
		temp_file_folder = temp_root_folder + "\\" + version_code + "\\FILE";
		temp_block_folder = temp_root_folder + "\\" + version_code + "\\BLOCK";
		int result = 0;
		if (FileUtil.CreateDirIfNoExist(temp_block_folder) != 0 || FileUtil.CreateDirIfNoExist(temp_file_folder) != 0)
		{
			result = -1;
		}
		return result;
	}

	public void ClearBuffer()
	{
		ClearFileBuffer();
		ClearBlockBuffer();
	}

	public void ClearFileBuffer()
	{
		string s_pattern = "*.prg";
		string[] fileListInDir = FileUtil.GetFileListInDir(temp_file_folder, s_pattern, 10000);
		if (fileListInDir != null)
		{
			for (int i = 0; i < fileListInDir.Length; i++)
			{
				FileUtil.RemoveFile(fileListInDir[i]);
			}
		}
	}

	public void ClearBlockBuffer()
	{
		string s_pattern = "*.blk";
		string[] fileListInDir = FileUtil.GetFileListInDir(temp_block_folder, s_pattern, 10000);
		if (fileListInDir != null)
		{
			for (int i = 0; i < fileListInDir.Length; i++)
			{
				FileUtil.RemoveFile(fileListInDir[i]);
			}
		}
	}

	public int MakeFileFromBlocks(DownloadFileInfo file_info, int num_blocks)
	{
		string[] array = new string[num_blocks];
		int fileId = file_info.GetFileId();
		for (int i = 0; i < array.Length; i++)
		{
			array[i] = GenerateBlockTempFPath(fileId, i);
		}
		string text = GenerateTempProgFilePath(fileId);
		FileUtil.RemoveFile(text);
		int num;
		if (FileUtil.MergeFiles(array, text, out err_msg) != 0)
		{
			num = -1;
		}
		else
		{
			for (int i = 0; i < array.Length; i++)
			{
				FileUtil.RemoveFile(array[i]);
			}
			num = 0;
		}
		if (num != 0)
		{
			return -1;
		}
		return num;
	}

	public void CheckExistingDownloaded(out int downloading_file_id, out int[] file_id_list, out int[] block_index_list)
	{
		block_index_list = null;
		CheckExistingFileIds(out file_id_list);
		CheckExistingBlockIndex(out downloading_file_id, out block_index_list);
	}

	private void CheckExistingBlockIndex(out int downloading_file_ind, out int[] blk_ind_list)
	{
		string s_pattern = "*.blk";
		downloading_file_ind = -1;
		blk_ind_list = null;
		if (FindBufferedFiles(s_pattern, temp_block_folder, out var fpath_list) == 0 && fpath_list != null && FindBlockIndexes(fpath_list, out blk_ind_list, out var file_id_list) == 0 && blk_ind_list != null && file_id_list != null)
		{
			if (CheckIndexdList(blk_ind_list, file_id_list, out var downloading_file_id) == 0)
			{
				downloading_file_ind = downloading_file_id;
				return;
			}
			blk_ind_list = null;
			downloading_file_ind = -1;
		}
	}

	private void CheckExistingFileIds(out int[] id_list)
	{
		string s_pattern = "*.prg";
		id_list = null;
		if (FindBufferedFiles(s_pattern, temp_file_folder, out var fpath_list) == 0 && fpath_list != null && FindFileInds(fpath_list, out var id_list2) == 0 && id_list2 != null)
		{
			ArrayUtil.Sort(id_list2);
			id_list = ArrayUtil.Copy(id_list2);
		}
	}

	private int FindBlockIndexes(string[] fpath_list, out int[] blk_ind_list, out int[] file_id_list)
	{
		if (fpath_list == null)
		{
			blk_ind_list = null;
			file_id_list = null;
			return -1;
		}
		int[] array = new int[fpath_list.Length];
		int[] array2 = new int[fpath_list.Length];
		int num = 0;
		int result = 0;
		for (int i = 0; i < fpath_list.Length; i++)
		{
			string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fpath_list[i]);
			string[] array3 = fileNameWithoutExtension.Split(new char[1] { '_' });
			if (array3.Length == 2)
			{
				int num2 = -1;
				int num3 = -1;
				bool flag = true;
				try
				{
					num2 = Convert.ToInt32(array3[0]);
					num3 = Convert.ToInt32(array3[1]);
				}
				catch
				{
					flag = false;
				}
				if (flag && num3 >= 0)
				{
					array[num] = num3;
					array2[num] = num2;
					num++;
				}
			}
		}
		if (num <= 0)
		{
			blk_ind_list = null;
			file_id_list = null;
		}
		else
		{
			blk_ind_list = new int[num];
			file_id_list = new int[num];
			for (int i = 0; i < num; i++)
			{
				blk_ind_list[i] = array[i];
				file_id_list[i] = array2[i];
			}
		}
		return result;
	}

	private int FindFileInds(string[] fpath_list, out int[] id_list)
	{
		if (fpath_list == null)
		{
			id_list = null;
			return -1;
		}
		int[] array = new int[fpath_list.Length];
		int num = 0;
		int result = 0;
		for (int i = 0; i < fpath_list.Length; i++)
		{
			string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fpath_list[i]);
			bool flag = true;
			int num2 = -1;
			try
			{
				num2 = Convert.ToInt32(fileNameWithoutExtension);
			}
			catch
			{
				flag = false;
			}
			if (flag && num2 >= 0)
			{
				array[num] = num2;
				num++;
			}
		}
		if (num <= 0)
		{
			id_list = null;
		}
		else
		{
			id_list = new int[num];
			for (int i = 0; i < num; i++)
			{
				id_list[i] = array[i];
			}
		}
		return result;
	}

	private int FindBufferedFiles(string s_pattern, string folder, out string[] fpath_list)
	{
		DirectoryInfo directoryInfo = new DirectoryInfo(folder);
		int num = 0;
		int num2 = 0;
		string[] array = new string[2000];
		try
		{
			FileInfo[] files = directoryInfo.GetFiles(s_pattern);
			foreach (FileInfo fileInfo in files)
			{
				array[num2] = fileInfo.FullName;
				if (num2 >= 1999)
				{
					break;
				}
				num2++;
			}
		}
		catch
		{
			num = -1;
		}
		if (num != 0)
		{
			fpath_list = null;
			return num;
		}
		if (num2 == 0)
		{
			fpath_list = null;
		}
		else
		{
			fpath_list = new string[num2];
			for (int j = 0; j < num2; j++)
			{
				fpath_list[j] = array[j];
			}
		}
		return num;
	}

	private int CheckIndexdList(int[] blk_ind_list, int[] file_id_list, out int downloading_file_id)
	{
		bool flag = true;
		downloading_file_id = -1;
		int num = file_id_list[0];
		for (int i = 1; i < file_id_list.Length; i++)
		{
			if (file_id_list[i] != num)
			{
				flag = false;
				break;
			}
		}
		if (!flag)
		{
			return -1;
		}
		ArrayUtil.Sort(blk_ind_list);
		bool flag2 = true;
		for (int i = 0; i < blk_ind_list.Length; i++)
		{
			if (blk_ind_list[i] != i)
			{
				flag2 = false;
				break;
			}
		}
		int result = 0;
		if (!flag2)
		{
			result = -1;
		}
		else
		{
			downloading_file_id = file_id_list[0];
		}
		return result;
	}

	private int CheckIdList(int[] id_list, out int lastest_id)
	{
		if (id_list == null)
		{
			lastest_id = -1;
			return -1;
		}
		ArrayUtil.Sort(id_list);
		bool flag = true;
		int result = 0;
		if (!flag)
		{
			result = -1;
		}
		lastest_id = id_list[^1];
		return result;
	}

	public int SaveBlockIntoFile(byte[] bytes, int file_id, int block_ind)
	{
		string text = GenerateBlockDownloadingFPath(file_id, block_ind);
		string text2 = GenerateBlockTempFPath(file_id, block_ind);
		int result = -1;
		FileUtil.RemoveFile(text);
		if (FileUtil.SaveFileContent(text, bytes) != 0)
		{
			return result;
		}
		FileUtil.RemoveFile(text2);
		if (FileUtil.MoveFile(text, text2) != 0)
		{
			FileUtil.RemoveFile(text2);
			FileUtil.RemoveFile(text);
		}
		else
		{
			result = 0;
		}
		return result;
	}

	public string[] BufFilePaths(int[] file_id_list)
	{
		if (file_id_list == null)
		{
			return null;
		}
		string[] array = new string[file_id_list.Length];
		for (int i = 0; i < file_id_list.Length; i++)
		{
			array[i] = GenerateTempProgFilePath(file_id_list[i]);
		}
		return array;
	}

	private string GenerateTempProgFilePath(int file_id)
	{
		string text = file_id + ".prg";
		return temp_file_folder + "\\" + text;
	}

	private string GenerateBlockDownloadingFPath(int file_id, int block_ind)
	{
		string text = file_id.ToString() + '_' + block_ind.ToString() + ".dwl";
		return temp_block_folder + "\\" + text;
	}

	private string GenerateBlockTempFPath(int file_ind, int block_ind)
	{
		string text = file_ind.ToString() + '_' + block_ind.ToString() + ".blk";
		return temp_block_folder + "\\" + text;
	}
}
