using System;
using System.Security.Cryptography.X509Certificates;

namespace CA;

internal class CA_FPT
{
	private static string LableOfMST = "MST:";

	public static X509Certificate2 GetCertificate()
	{
		X509Store x509Store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
		x509Store.Open(OpenFlags.ReadOnly);
		X509Certificate2Collection certificates = x509Store.Certificates;
		X509Certificate2 result = null;
		X509Certificate2Collection x509Certificate2Collection = X509Certificate2UI.SelectFromCollection(certificates, "Danh sách chứng thư số", "Hãy chọn chứng thư số của bạn", (X509SelectionFlag)0);
		if (x509Certificate2Collection.Count > 0)
		{
			X509Certificate2Enumerator enumerator = x509Certificate2Collection.GetEnumerator();
			enumerator.MoveNext();
			result = enumerator.Current;
		}
		x509Store.Close();
		return result;
	}

	public static bool CheckValidDate(X509Certificate2 cer, DateTime dtCheck)
	{
		try
		{
			return (cer.NotBefore <= dtCheck) & (dtCheck <= cer.NotAfter);
		}
		catch (Exception ex)
		{
			throw ex;
		}
	}

	public static string GetMaSoThueOfCA(X509Certificate2 cer)
	{
		string result = "";
		string name = cer.SubjectName.Name;
		string[] array = name.Split(new char[1] { ',' });
		string text = "";
		for (int i = 0; i < array.Length; i++)
		{
			if (array[i].IndexOf(LableOfMST) > -1)
			{
				text = array[i];
				int num = text.IndexOf(LableOfMST);
				result = text.Substring(num + LableOfMST.Length);
				break;
			}
		}
		return result;
	}
}
