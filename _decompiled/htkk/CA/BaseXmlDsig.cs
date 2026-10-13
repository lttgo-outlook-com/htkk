using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;

namespace CA;

public abstract class BaseXmlDsig
{
	public abstract string Name { get; }

	public abstract string SignXml(XmlDocument Document, X509Certificate2 cer, string tagContaint);

	public bool VerifyXml(string SignedXmlDocumentString)
	{
		byte[] bytes = Encoding.UTF8.GetBytes(SignedXmlDocumentString);
		using MemoryStream signedXmlDocumentStream = new MemoryStream(bytes);
		return VerifyXmlFromStream(signedXmlDocumentStream);
	}

	public bool VerifyXmlFromStream(Stream SignedXmlDocumentStream)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0018: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected O, but got Unknown
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Expected O, but got Unknown
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_005d: Expected O, but got Unknown
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_0092: Unknown result type (might be due to invalid IL or missing references)
		XmlDocument val = new XmlDocument();
		val.PreserveWhitespace = true;
		val.Load(SignedXmlDocumentStream);
		SignedXml val2 = new SignedXml(val);
		XmlNode val3 = val.GetElementsByTagName("Signature")[0];
		val2.LoadXml((XmlElement)val3);
		X509Certificate2 x509Certificate = null;
		foreach (KeyInfoClause item in val2.KeyInfo)
		{
			KeyInfoClause val4 = item;
			if (val4 is KeyInfoX509Data && ((KeyInfoX509Data)val4).Certificates.Count > 0)
			{
				x509Certificate = (X509Certificate2)((KeyInfoX509Data)val4).Certificates[0];
			}
		}
		return val2.CheckSignature(x509Certificate, true);
	}
}
