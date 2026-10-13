using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace HTKK.GW.TEST;

public class Cryptography
{
	public class RijndaelEnhanced
	{
		private static string DEFAULT_HASH_ALGORITHM = "SHA1";

		private static int DEFAULT_KEY_SIZE = 256;

		private static int MAX_ALLOWED_SALT_LEN = 255;

		private static int MIN_ALLOWED_SALT_LEN = 4;

		private static int DEFAULT_MIN_SALT_LEN = MIN_ALLOWED_SALT_LEN;

		private static int DEFAULT_MAX_SALT_LEN = 8;

		private ICryptoTransform decryptor = null;

		private ICryptoTransform encryptor = null;

		private int maxSaltLen = -1;

		private int minSaltLen = -1;

		public RijndaelEnhanced(string passPhrase)
			: this(passPhrase, null)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector)
			: this(passPhrase, initVector, -1)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector, int minSaltLen)
			: this(passPhrase, initVector, minSaltLen, -1)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector, int minSaltLen, int maxSaltLen)
			: this(passPhrase, initVector, minSaltLen, maxSaltLen, -1)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector, int minSaltLen, int maxSaltLen, int keySize)
			: this(passPhrase, initVector, minSaltLen, maxSaltLen, keySize, null)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector, int minSaltLen, int maxSaltLen, int keySize, string hashAlgorithm)
			: this(passPhrase, initVector, minSaltLen, maxSaltLen, keySize, hashAlgorithm, null)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector, int minSaltLen, int maxSaltLen, int keySize, string hashAlgorithm, string saltValue)
			: this(passPhrase, initVector, minSaltLen, maxSaltLen, keySize, hashAlgorithm, saltValue, 1)
		{
		}

		public RijndaelEnhanced(string passPhrase, string initVector, int minSaltLen, int maxSaltLen, int keySize, string hashAlgorithm, string saltValue, int passwordIterations)
		{
			if (minSaltLen < MIN_ALLOWED_SALT_LEN)
			{
				this.minSaltLen = DEFAULT_MIN_SALT_LEN;
			}
			else
			{
				this.minSaltLen = minSaltLen;
			}
			if (maxSaltLen < 0 || maxSaltLen > MAX_ALLOWED_SALT_LEN)
			{
				this.maxSaltLen = DEFAULT_MAX_SALT_LEN;
			}
			else
			{
				this.maxSaltLen = maxSaltLen;
			}
			if (keySize <= 0)
			{
				keySize = DEFAULT_KEY_SIZE;
			}
			hashAlgorithm = ((hashAlgorithm != null) ? hashAlgorithm.ToUpper().Replace("-", "") : DEFAULT_HASH_ALGORITHM);
			byte[] array = null;
			byte[] array2 = null;
			array = ((initVector != null) ? Encoding.ASCII.GetBytes(initVector) : new byte[0]);
			byte[] bytes = new PasswordDeriveBytes(passPhrase, (saltValue != null) ? Encoding.ASCII.GetBytes(saltValue) : new byte[0], hashAlgorithm, passwordIterations).GetBytes(keySize / 8);
			RijndaelManaged rijndaelManaged = new RijndaelManaged();
			if (array.Length == 0)
			{
				rijndaelManaged.Mode = CipherMode.ECB;
			}
			else
			{
				rijndaelManaged.Mode = CipherMode.CBC;
			}
			encryptor = rijndaelManaged.CreateEncryptor(bytes, array);
			decryptor = rijndaelManaged.CreateDecryptor(bytes, array);
		}

		public string Encrypt(string plainText)
		{
			return Encrypt(Encoding.UTF8.GetBytes(plainText));
		}

		public string Encrypt(byte[] plainTextBytes)
		{
			return Convert.ToBase64String(EncryptToBytes(plainTextBytes));
		}

		public byte[] EncryptToBytes(string plainText)
		{
			return EncryptToBytes(Encoding.UTF8.GetBytes(plainText));
		}

		public byte[] EncryptToBytes(byte[] plainTextBytes)
		{
			byte[] array = AddSalt(plainTextBytes);
			MemoryStream memoryStream = new MemoryStream();
			lock (this)
			{
				CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write);
				cryptoStream.Write(array, 0, array.Length);
				cryptoStream.FlushFinalBlock();
				byte[] result = memoryStream.ToArray();
				memoryStream.Close();
				cryptoStream.Close();
				return result;
			}
		}

		public string Decrypt(string cipherText)
		{
			return Decrypt(Convert.FromBase64String(cipherText));
		}

		public string Decrypt(byte[] cipherTextBytes)
		{
			return Encoding.UTF8.GetString(DecryptToBytes(cipherTextBytes));
		}

		public byte[] DecryptToBytes(string cipherText)
		{
			return DecryptToBytes(Convert.FromBase64String(cipherText));
		}

		public byte[] DecryptToBytes(byte[] cipherTextBytes)
		{
			byte[] array = null;
			byte[] array2 = null;
			int num = 0;
			int num2 = 0;
			MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
			array = new byte[cipherTextBytes.Length];
			lock (this)
			{
				CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
				num = cryptoStream.Read(array, 0, array.Length);
				memoryStream.Close();
				cryptoStream.Close();
			}
			if (maxSaltLen > 0 && maxSaltLen >= minSaltLen)
			{
				num2 = (array[0] & 3) | (array[1] & 0xC) | (array[2] & 0x30) | (array[3] & 0xC0);
			}
			array2 = new byte[num - num2];
			Array.Copy(array, num2, array2, 0, num - num2);
			return array2;
		}

		private byte[] AddSalt(byte[] plainTextBytes)
		{
			if (maxSaltLen == 0 || maxSaltLen < minSaltLen)
			{
				return plainTextBytes;
			}
			byte[] array = GenerateSalt();
			byte[] array2 = new byte[plainTextBytes.Length + array.Length];
			Array.Copy(array, array2, array.Length);
			Array.Copy(plainTextBytes, 0, array2, array.Length, plainTextBytes.Length);
			return array2;
		}

		private byte[] GenerateSalt()
		{
			int num = 0;
			num = ((minSaltLen != maxSaltLen) ? GenerateRandomNumber(minSaltLen, maxSaltLen) : minSaltLen);
			byte[] array = new byte[num];
			RNGCryptoServiceProvider rNGCryptoServiceProvider = new RNGCryptoServiceProvider();
			rNGCryptoServiceProvider.GetNonZeroBytes(array);
			array[0] = (byte)((array[0] & 0xFC) | (num & 3));
			array[1] = (byte)((array[1] & 0xF3) | (num & 0xC));
			array[2] = (byte)((array[2] & 0xCF) | (num & 0x30));
			array[3] = (byte)((array[3] & 0x3F) | (num & 0xC0));
			return array;
		}

		private int GenerateRandomNumber(int minValue, int maxValue)
		{
			byte[] array = new byte[4];
			RNGCryptoServiceProvider rNGCryptoServiceProvider = new RNGCryptoServiceProvider();
			rNGCryptoServiceProvider.GetBytes(array);
			int seed = ((array[0] & 0x7F) << 24) | (array[1] << 16) | (array[2] << 8) | array[3];
			Random random = new Random(seed);
			return random.Next(minValue, maxValue + 1);
		}
	}
}
