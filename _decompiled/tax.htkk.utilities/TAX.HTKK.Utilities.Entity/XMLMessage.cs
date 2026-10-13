using System.Xml;

namespace TAX.HTKK.Utilities.Entity;

public class XMLMessage
{
	private string msg;

	public string PathFile { get; set; }

	public XmlDocument XDocument { get; set; }

	public XMLMessage(string pathFile)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Expected O, but got Unknown
		PathFile = pathFile;
		XDocument = new XmlDocument();
		XDocument.Load(PathFile);
	}

	public string GetMessage(string id)
	{
		XmlNodeList elementsByTagName = XDocument.GetElementsByTagName("Infor");
		for (int i = 0; i < elementsByTagName.Count; i++)
		{
			if (id == ((XmlNode)elementsByTagName[i].Attributes["ID"]).Value)
			{
				msg = ((XmlNode)elementsByTagName[i].Attributes["Msg"]).Value;
			}
		}
		return msg;
	}
}
