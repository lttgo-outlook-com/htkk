using System;
using System.Xml.Serialization;

namespace HTKKDanhmuc;

[Serializable]
public class cell
{
	[XmlAttribute("col")]
	public string col { get; set; }

	[XmlAttribute("field")]
	public string field { get; set; }

	[XmlAttribute("ctltype")]
	public string ctltype { get; set; }
}
