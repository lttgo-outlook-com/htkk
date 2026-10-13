using System;
using System.IO;
using Foundation.Base;

namespace Foundation.Utility;

internal class FileUtil : BaseObject
{
	public const int EN_ACCESS_DENIED = -1;

	public const int EN_FILE_NOT_FOUND = -2;

	public const int EN_DIR_NOT_FOUND = -3;

	public const int EN_DRIVE_NOT_FOUND = -4;

	public const int EN_PATH_TOO_LONG = -5;

	public static int GetLastWriteTime(string f_path, out DateTime date_time)
	{
		int result = 0;
		try
		{
			date_time = File.GetLastWriteTime(f_path);
		}
		catch
		{
			date_time = DateTime.Now;
			result = -1;
		}
		return result;
	}

	public static int CreateDir(string dir_path)
	{
		int result = 0;
		try
		{
			Directory.CreateDirectory(dir_path);
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public static int CreateDirIfNoExist(string dir_path)
	{
		int result = 0;
		try
		{
			if (!Directory.Exists(dir_path))
			{
				Directory.CreateDirectory(dir_path);
			}
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public static int GetFileSize(string f_path, out long f_size)
	{
		int result = 0;
		try
		{
			FileInfo fileInfo = new FileInfo(f_path);
			f_size = fileInfo.Length;
		}
		catch (Exception)
		{
			f_size = -1L;
			result = -1;
		}
		return result;
	}

	public static int CheckDirExist(string dir_path, out bool exist)
	{
		int result = 0;
		try
		{
			if (Directory.Exists(dir_path))
			{
				exist = true;
			}
			else
			{
				exist = false;
			}
		}
		catch
		{
			exist = false;
			result = -1;
		}
		return result;
	}

	public static string[] GetFileListInDir(string dir_name, string s_pattern, int max_num_files)
	{
		int num = 0;
		string[] array = new string[max_num_files];
		string[] array2;
		try
		{
			DirectoryInfo directoryInfo = new DirectoryInfo(dir_name);
			FileInfo[] files = directoryInfo.GetFiles(s_pattern);
			foreach (FileInfo fileInfo in files)
			{
				string fullName = fileInfo.FullName;
				array[num] = fullName;
				num++;
				if (num >= max_num_files)
				{
					break;
				}
			}
			if (num <= 0)
			{
				array2 = null;
			}
			else
			{
				array2 = new string[num];
				for (int j = 0; j < num; j++)
				{
					array2[j] = array[j];
				}
			}
		}
		catch
		{
			array2 = null;
		}
		return array2;
	}

	public static int SetReadOnly(string f_path, bool is_read_only)
	{
		if (f_path == null)
		{
			return -1;
		}
		int result = 0;
		try
		{
			FileInfo fileInfo = new FileInfo(f_path);
			if (fileInfo.Exists)
			{
				fileInfo.IsReadOnly = is_read_only;
			}
			else
			{
				result = -1;
			}
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public static int CheckReadOnly(string f_path, out bool read_only)
	{
		int result = 0;
		read_only = false;
		if (f_path == null)
		{
			return -1;
		}
		try
		{
			FileInfo fileInfo = new FileInfo(f_path);
			if (fileInfo.Exists)
			{
				read_only = fileInfo.IsReadOnly;
			}
			else
			{
				read_only = false;
			}
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public static int CheckExist(string f_path, out bool exist)
	{
		int result = 0;
		try
		{
			if (f_path == null)
			{
				exist = false;
				result = -1;
			}
			else if (File.Exists(f_path))
			{
				exist = true;
			}
			else
			{
				exist = false;
			}
		}
		catch
		{
			exist = false;
			result = -1;
		}
		return result;
	}

	public static int MergeFiles(string[] source_file_path_list, string dest_file_path, out string err_msg)
	{
		int result = 0;
		err_msg = "";
		try
		{
			using FileStream fileStream = new FileStream(dest_file_path, FileMode.Create);
			for (int i = 0; i < source_file_path_list.Length; i++)
			{
				byte[] array = File.ReadAllBytes(source_file_path_list[i]);
				fileStream.Write(array, 0, array.Length);
			}
			fileStream.Close();
		}
		catch (Exception ex)
		{
			err_msg = ex.Message;
			result = -1;
		}
		return result;
	}

	public static int RemoveFile(string f_path)
	{
		int result = 0;
		try
		{
			File.Delete(f_path);
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public static int MoveFile(string old_path, string new_path)
	{
		int result = 0;
		try
		{
			if (old_path == null || new_path == null)
			{
				result = -1;
			}
			else
			{
				File.Move(old_path, new_path);
			}
		}
		catch
		{
			result = -1;
		}
		return result;
	}

	public static int Copy(string source, string dest)
	{
		int result = 0;
		try
		{
			if (source == null || dest == null)
			{
				result = -1;
			}
			else
			{
				File.Copy(source, dest);
			}
		}
		catch (Exception)
		{
			result = -1;
		}
		return result;
	}

	public static int CopyOverwriteIfExist(string source, string dest, out int err_no)
	{
		err_no = -10001;
		int result = 0;
		if (source == null || dest == null)
		{
			return -1;
		}
		try
		{
			if (CheckExist(dest, out var exist) != 0)
			{
				result = -1;
			}
			else if (exist)
			{
				File.Copy(source, dest, overwrite: true);
			}
			else
			{
				File.Copy(source, dest);
			}
		}
		catch (Exception ex)
		{
			if (ex is UnauthorizedAccessException)
			{
				err_no = -1;
			}
			else if (ex is FileNotFoundException)
			{
				err_no = -2;
			}
			else if (ex is DirectoryNotFoundException)
			{
				err_no = -3;
			}
			else if (ex is DriveNotFoundException)
			{
				err_no = -4;
			}
			else if (ex is PathTooLongException)
			{
				err_no = -5;
			}
			else
			{
				err_no = -10001;
			}
			result = -1;
		}
		return result;
	}

	public static int SaveFileContent(string file_path, byte[] file_content)
	{
		int result = -1;
		if (file_content == null)
		{
			return result;
		}
		try
		{
			using (FileStream fileStream = new FileStream(file_path, FileMode.CreateNew))
			{
				fileStream.Write(file_content, 0, file_content.Length);
				fileStream.Close();
			}
			result = 0;
		}
		catch
		{
			result = 0;
		}
		return result;
	}

	public static int SaveFileContent(string file_path, string file_content, out int err_no)
	{
		int result = 0;
		err_no = -10001;
		try
		{
			if (file_content == null)
			{
				result = -1;
			}
			else
			{
				using StreamWriter streamWriter = new StreamWriter(file_path);
				streamWriter.Write(file_content);
				streamWriter.Close();
			}
		}
		catch (Exception ex)
		{
			if (ex is UnauthorizedAccessException)
			{
				err_no = -1;
			}
			else if (ex is FileNotFoundException)
			{
				err_no = -2;
			}
			else
			{
				err_no = -10001;
			}
			result = -1;
		}
		return result;
	}

	public static int SaveFileContent(string file_path, string file_content)
	{
		int result = 0;
		try
		{
			if (file_content == null)
			{
				result = -1;
			}
			else
			{
				using StreamWriter streamWriter = new StreamWriter(file_path);
				streamWriter.Write(file_content);
				streamWriter.Close();
			}
		}
		catch (Exception)
		{
			result = -1;
		}
		return result;
	}

	public static int LoadFileContent(string file_path, out string file_content)
	{
		int result = 0;
		try
		{
			using StreamReader streamReader = new StreamReader(file_path);
			file_content = streamReader.ReadToEnd();
		}
		catch
		{
			file_content = null;
			result = -1;
		}
		return result;
	}

	public static int LoadFileContent(string file_path, out byte[] file_content)
	{
		int result = 0;
		try
		{
			file_content = File.ReadAllBytes(file_path);
		}
		catch
		{
			file_content = null;
			result = -1;
		}
		return result;
	}
}
