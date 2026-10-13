using System;
using System.Xml;
using System.Xml.Linq;

namespace TAX.HTKK.Utilities.Core;

public static class XMLHelper
{
	public static XmlDocument ToXmlDocument(this XDocument xDocument)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		XmlReader val2 = ((XNode)xDocument).CreateReader();
		try
		{
			val.Load(val2);
		}
		finally
		{
			((IDisposable)val2)?.Dispose();
		}
		return val;
	}

	public static XDocument ToXDocument(this XmlDocument xmlDocument)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		XmlNodeReader val = new XmlNodeReader((XmlNode)(object)xmlDocument);
		try
		{
			((XmlReader)val).MoveToContent();
			return XDocument.Load((XmlReader)(object)val);
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
	}
}
