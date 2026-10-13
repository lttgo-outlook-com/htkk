using System;
using System.Management;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace Foundation;

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
		catch (Exception)
		{
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
		catch (Exception)
		{
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
			Log.LogEvent(MethodBase.GetCurrentMethod().Name, ex);
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
			Log.LogEvent(MethodBase.GetCurrentMethod().Name, ex);
			return string.Empty;
		}
	}

	public static string GetProcessorId()
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		ManagementClass val = new ManagementClass("win32_processor");
		ManagementObjectCollection instances = val.GetInstances();
		string result = string.Empty;
		ManagementObjectEnumerator enumerator = instances.GetEnumerator();
		try
		{
			if (enumerator.MoveNext())
			{
				ManagementObject val2 = (ManagementObject)enumerator.Current;
				result = ((ManagementBaseObject)val2).Properties["processorID"].Value.ToString();
			}
		}
		finally
		{
			((IDisposable)enumerator)?.Dispose();
		}
		return result;
	}

	public static string GetHDDSerialNo()
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		ManagementClass val = new ManagementClass("Win32_LogicalDisk");
		ManagementObjectCollection instances = val.GetInstances();
		string text = "";
		ManagementObjectEnumerator enumerator = instances.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ManagementObject val2 = (ManagementObject)enumerator.Current;
				text += Convert.ToString(((ManagementBaseObject)val2)["VolumeSerialNumber"]);
			}
		}
		finally
		{
			((IDisposable)enumerator)?.Dispose();
		}
		return text;
	}

	public static string GetMACAddress()
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		//IL_002b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Expected O, but got Unknown
		ManagementClass val = new ManagementClass("Win32_NetworkAdapterConfiguration");
		ManagementObjectCollection instances = val.GetInstances();
		string text = string.Empty;
		ManagementObjectEnumerator enumerator = instances.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ManagementObject val2 = (ManagementObject)enumerator.Current;
				if (text == string.Empty && (bool)((ManagementBaseObject)val2)["IPEnabled"])
				{
					text = ((ManagementBaseObject)val2)["MacAddress"].ToString();
				}
				val2.Dispose();
			}
		}
		finally
		{
			((IDisposable)enumerator)?.Dispose();
		}
		return text.Replace(":", "");
	}

	public static string GetBoardMaker()
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Expected O, but got Unknown
		ManagementObjectSearcher val = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
		ManagementObjectEnumerator enumerator = val.Get().GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ManagementObject val2 = (ManagementObject)enumerator.Current;
				try
				{
					return ((ManagementBaseObject)val2).GetPropertyValue("Manufacturer").ToString();
				}
				catch
				{
				}
			}
		}
		finally
		{
			((IDisposable)enumerator)?.Dispose();
		}
		return "Board Maker: Unknown";
	}

	public static string GetBoardProductId()
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Expected O, but got Unknown
		ManagementObjectSearcher val = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
		ManagementObjectEnumerator enumerator = val.Get().GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ManagementObject val2 = (ManagementObject)enumerator.Current;
				try
				{
					return ((ManagementBaseObject)val2).GetPropertyValue("Product").ToString();
				}
				catch
				{
				}
			}
		}
		finally
		{
			((IDisposable)enumerator)?.Dispose();
		}
		return "Product: Unknown";
	}
}
