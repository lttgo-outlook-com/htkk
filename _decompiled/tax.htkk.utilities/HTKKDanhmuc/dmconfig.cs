using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace HTKKDanhmuc;

[Serializable]
public class dmconfig
{
	[XmlAttribute("version")]
	public string version { get; set; }

	public List<dmcf> dmcfs { get; set; }

	public string ToXML()
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Expected O, but got Unknown
		using StringWriter stringWriter = new StringWriter();
		XmlSerializer val = new XmlSerializer(GetType());
		val.Serialize((TextWriter)stringWriter, (object)this);
		return stringWriter.ToString();
	}
}
