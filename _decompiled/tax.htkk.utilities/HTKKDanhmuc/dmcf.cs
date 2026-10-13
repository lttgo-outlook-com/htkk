using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace HTKKDanhmuc;

[Serializable]
public class dmcf
{
	[XmlAttribute("type")]
	public string type { get; set; }

	[XmlAttribute("checkbox")]
	public string checkbox { get; set; }

	public string start { get; set; }

	public string active { get; set; }

	public List<cell> cells { get; set; }
}
