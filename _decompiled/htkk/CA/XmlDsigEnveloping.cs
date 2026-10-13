using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;

namespace CA;

public class XmlDsigEnveloping : BaseXmlDsig
{
	private bool c14 { get; set; }

	public override string Name => string.Format("XmlDsigEnveloping {0} c14", c14 ? "with" : "without");

	public XmlDsigEnveloping(bool c14)
	{
		this.c14 = c14;
	}

	public override string SignXml(XmlDocument Document, X509Certificate2 cer, string tagContaint)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Expected O, but got Unknown
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0033: Expected O, but got Unknown
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_006b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Expected O, but got Unknown
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Expected O, but got Unknown
		//IL_00a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		SignedXml val = new SignedXml(Document);
		val.SigningKey = cer.PrivateKey;
		KeyInfo val2 = new KeyInfo();
		RSACryptoServiceProvider rSACryptoServiceProvider = (RSACryptoServiceProvider)cer.PublicKey.Key;
		RSAKeyValue val3 = new RSAKeyValue((RSA)rSACryptoServiceProvider);
		val2.AddClause((KeyInfoClause)(object)val3);
		KeyInfoX509Data val4 = new KeyInfoX509Data((X509Certificate)cer);
		string name = cer.SubjectName.Name;
		val4.AddSubjectName(name);
		val2.AddClause((KeyInfoClause)(object)val4);
		val.KeyInfo = val2;
		DataObject val5 = new DataObject();
		val5.Id = "NODE_TO_SIGN";
		Reference val6 = new Reference();
		val6.Uri = "#_NODE_TO_SIGN";
		if (c14)
		{
			XmlDsigEnvelopedSignatureTransform val7 = new XmlDsigEnvelopedSignatureTransform(true);
			val6.AddTransform((Transform)(object)val7);
		}
		val.AddReference(val6);
		val.ComputeSignature();
		XmlElement xml = val.GetXml();
		return string.IsNullOrEmpty(tagContaint) ? ((XmlNode)xml).OuterXml : string.Format("<{0}>{1}</{0}>", tagContaint, ((XmlNode)xml).OuterXml);
	}
}
