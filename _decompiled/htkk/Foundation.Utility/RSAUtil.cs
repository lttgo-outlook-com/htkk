using System;
using System.Security.Cryptography;
using Foundation.Base;

namespace Foundation.Utility;

internal class RSAUtil : BaseObject
{
	public const int ERR_NO_PARAMS_NULL = -1;

	public const int ERR_N = -2;

	public static int GenerateKeys(out string public_key, out string private_key)
	{
		int num = 0;
		private_key = null;
		public_key = null;
		try
		{
			RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
			string text = rSACryptoServiceProvider.ToXmlString(includePrivateParameters: false);
			string text2 = rSACryptoServiceProvider.ToXmlString(includePrivateParameters: true);
		}
		catch
		{
			private_key = null;
			public_key = null;
			num = -1;
		}
		if (num != 0)
		{
			num = -1;
		}
		if (public_key == null || private_key == null)
		{
			return -1;
		}
		return 0;
	}

	public static byte[] Encrypt(byte[] origin_data, string public_key)
	{
		byte[] array = null;
		try
		{
			RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
			rSACryptoServiceProvider.FromXmlString(public_key);
			return rSACryptoServiceProvider.Encrypt(origin_data, fOAEP: true);
		}
		catch
		{
			return null;
		}
	}

	public static byte[] Decrypt(byte[] encrypted_data, string private_key)
	{
		byte[] result = null;
		try
		{
			RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
			rSACryptoServiceProvider.FromXmlString(private_key);
			result = rSACryptoServiceProvider.Decrypt(encrypted_data, fOAEP: true);
		}
		catch
		{
		}
		return result;
	}

	public static byte[] Sign(byte[] origin_data, string private_key)
	{
		byte[] array = null;
		try
		{
			RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
			rSACryptoServiceProvider.FromXmlString(private_key);
			return rSACryptoServiceProvider.SignData(origin_data, CryptoConfig.MapNameToOID("SHA512"));
		}
		catch
		{
			return null;
		}
	}

	public static byte[] SignFile(string file_path, string private_key)
	{
		if (file_path == null || private_key == null)
		{
			return null;
		}
		byte[] file_content = null;
		byte[] array = null;
		try
		{
			array = ((FileUtil.LoadFileContent(file_path, out file_content) != 0) ? null : Sign(file_content, private_key));
		}
		catch
		{
			array = null;
		}
		return array;
	}

	public static bool VerifySignature(byte[] origin_data, byte[] signature, string public_key)
	{
		bool result = false;
		using (RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider())
		{
			try
			{
				rSACryptoServiceProvider.FromXmlString(public_key);
				result = rSACryptoServiceProvider.VerifyData(origin_data, CryptoConfig.MapNameToOID("SHA512"), signature);
			}
			catch (Exception)
			{
				result = false;
			}
		}
		return result;
	}

	public static bool VerifyFileSignature(string file_path, byte[] signature, string public_key)
	{
		if (file_path == null || public_key == null || signature == null)
		{
			return false;
		}
		byte[] file_content = null;
		bool flag = false;
		try
		{
			flag = FileUtil.LoadFileContent(file_path, out file_content) == 0 && VerifySignature(file_content, signature, public_key);
		}
		catch
		{
			flag = false;
		}
		return flag;
	}
}
