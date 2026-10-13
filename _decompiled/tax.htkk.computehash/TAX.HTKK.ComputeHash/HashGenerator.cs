using System;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Security.Cryptography;
using System.Text;

namespace TAX.HTKK.ComputeHash;

public class HashGenerator
{
	private static readonly object locker = new object();

	public static string GenerateKey(object sourceObject, string encryptKey)
	{
		if (sourceObject == null)
		{
			throw new ArgumentNullException("Null as parameter is not allowed");
		}
		try
		{
			string objectToSerialize = Encrypt.EncryptText(sourceObject.ToString(), encryptKey);
			return ComputeHash(ObjectToByteArray(objectToSerialize));
		}
		catch (AmbiguousMatchException ex)
		{
			throw new ApplicationException("Could not definitly decide if object is serializable. Message:" + ex.Message);
		}
	}

	private static byte[] ObjectToByteArray(object objectToSerialize)
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryFormatter binaryFormatter = new BinaryFormatter();
		try
		{
			lock (locker)
			{
				binaryFormatter.Serialize(memoryStream, objectToSerialize);
			}
			return memoryStream.ToArray();
		}
		catch (SerializationException ex)
		{
			Console.WriteLine("Error occured during serialization. Message: " + ex.Message);
			return null;
		}
		finally
		{
			memoryStream.Close();
		}
	}

	private static string ComputeHash(byte[] objectAsBytes)
	{
		MD5 mD = new MD5CryptoServiceProvider();
		try
		{
			byte[] array = mD.ComputeHash(objectAsBytes);
			StringBuilder stringBuilder = new StringBuilder();
			for (int i = 0; i < array.Length; i++)
			{
				stringBuilder.Append(array[i].ToString("X2"));
			}
			return stringBuilder.ToString();
		}
		catch (ArgumentNullException ex)
		{
			Console.WriteLine("Hash has not been generated." + ex.Message);
			return null;
		}
	}

	public static bool VerifyHash(string input, string hash, string encryptKey)
	{
		string x = GenerateKey(input, encryptKey);
		StringComparer ordinalIgnoreCase = StringComparer.OrdinalIgnoreCase;
		if (0 == ordinalIgnoreCase.Compare(x, hash))
		{
			return true;
		}
		return false;
	}
}
