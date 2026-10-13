using System;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace TAX.HTKK.Utilities.Core;

public class HTKKSecurity
{
	public static string Encrypt(string source)
	{
		try
		{
			byte[] bytes = Encoding.UTF8.GetBytes(source);
			MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
			byte[] key = mD5CryptoServiceProvider.ComputeHash(Encoding.UTF8.GetBytes("HTKK CMS-SOFT"));
			mD5CryptoServiceProvider.Clear();
			TripleDESCryptoServiceProvider tripleDESCryptoServiceProvider = new TripleDESCryptoServiceProvider();
			tripleDESCryptoServiceProvider.Key = key;
			tripleDESCryptoServiceProvider.Mode = CipherMode.ECB;
			tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7;
			ICryptoTransform cryptoTransform = tripleDESCryptoServiceProvider.CreateEncryptor();
			byte[] array = cryptoTransform.TransformFinalBlock(bytes, 0, bytes.Length);
			tripleDESCryptoServiceProvider.Clear();
			return Convert.ToBase64String(array, 0, array.Length);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string Decrypt(string source)
	{
		try
		{
			byte[] array = Convert.FromBase64String(source);
			MD5CryptoServiceProvider mD5CryptoServiceProvider = new MD5CryptoServiceProvider();
			byte[] key = mD5CryptoServiceProvider.ComputeHash(Encoding.UTF8.GetBytes("HTKK CMS-SOFT"));
			mD5CryptoServiceProvider.Clear();
			TripleDESCryptoServiceProvider tripleDESCryptoServiceProvider = new TripleDESCryptoServiceProvider();
			tripleDESCryptoServiceProvider.Key = key;
			tripleDESCryptoServiceProvider.Mode = CipherMode.ECB;
			tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7;
			ICryptoTransform cryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor();
			byte[] bytes = cryptoTransform.TransformFinalBlock(array, 0, array.Length);
			tripleDESCryptoServiceProvider.Clear();
			return Encoding.UTF8.GetString(bytes);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string ConvertToBase64(string strSource)
	{
		try
		{
			byte[] bytes = Encoding.Unicode.GetBytes(strSource);
			return Convert.ToBase64String(bytes);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string ConvertFromBase64(string strSource)
	{
		try
		{
			byte[] bytes = Convert.FromBase64String(strSource);
			return Encoding.Unicode.GetString(bytes);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string MaHoaDuLieu(string textMaHoa, string pathCert)
	{
		try
		{
			string text = "";
			X509Certificate2 x509Certificate = new X509Certificate2(pathCert);
			RSACryptoServiceProvider rSACryptoServiceProvider = (RSACryptoServiceProvider)x509Certificate.PublicKey.Key;
			string publicKey = rSACryptoServiceProvider.ToXmlString(includePrivateParameters: false);
			return Encrypt(textMaHoa, publicKey);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	private static byte[] Encryption(byte[] Data, RSAParameters RSAKey, bool DoOAEPPadding)
	{
		try
		{
			byte[] result;
			using (RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider())
			{
				rSACryptoServiceProvider.ImportParameters(RSAKey);
				result = rSACryptoServiceProvider.Encrypt(Data, DoOAEPPadding);
			}
			return result;
		}
		catch (CryptographicException ex)
		{
			Console.WriteLine(ex.Message);
			return null;
		}
	}

	private static byte[] Decryption(byte[] Data, RSAParameters RSAKey, bool DoOAEPPadding)
	{
		try
		{
			byte[] result;
			using (RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider())
			{
				rSACryptoServiceProvider.ImportParameters(RSAKey);
				result = rSACryptoServiceProvider.Decrypt(Data, DoOAEPPadding);
			}
			return result;
		}
		catch (CryptographicException ex)
		{
			Console.WriteLine(ex.ToString());
			return null;
		}
	}

	public static string GetPublicKey(string Certificate)
	{
		try
		{
			X509Certificate x509Certificate = X509Certificate.CreateFromCertFile(Certificate);
			return x509Certificate.GetPublicKeyString();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string Decrypt(string data, string privateKey)
	{
		byte[] rgb = Convert.FromBase64String(data);
		RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
		rSACryptoServiceProvider.FromXmlString(privateKey);
		byte[] bytes = rSACryptoServiceProvider.Decrypt(rgb, fOAEP: false);
		return Encoding.UTF8.GetString(bytes);
	}

	public static string Encrypt(string data, string publicKey)
	{
		RSACryptoServiceProvider rSACryptoServiceProvider = new RSACryptoServiceProvider();
		rSACryptoServiceProvider.FromXmlString(publicKey);
		byte[] bytes = Encoding.UTF8.GetBytes(data);
		byte[] inArray = rSACryptoServiceProvider.Encrypt(bytes, fOAEP: false).ToArray();
		return Convert.ToBase64String(inArray);
	}
}
