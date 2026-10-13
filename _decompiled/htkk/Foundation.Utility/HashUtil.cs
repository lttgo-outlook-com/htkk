using System.IO;
using System.Security.Cryptography;

namespace Foundation.Utility;

public class HashUtil
{
	public static byte[] ComputeHash(string file_path)
	{
		if (file_path == null)
		{
			return null;
		}
		byte[] result;
		try
		{
			using MD5 mD = MD5.Create();
			using FileStream inputStream = File.OpenRead(file_path);
			result = mD.ComputeHash(inputStream);
		}
		catch
		{
			result = null;
		}
		return result;
	}
}
