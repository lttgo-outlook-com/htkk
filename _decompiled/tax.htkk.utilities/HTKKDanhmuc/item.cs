using System;
using System.Xml.Serialization;

namespace HTKKDanhmuc;

[Serializable]
public class item
{
	[XmlAttribute("id")]
	public string id { get; set; }

	public string check { get; set; }

	public string bold { get; set; }

	public string ma { get; set; }

	public string macha { get; set; }

	public string ten { get; set; }

	public string hltu { get; set; }

	public string hlden { get; set; }

	public string thuesuat { get; set; }

	public string dvt { get; set; }

	public string mamuc { get; set; }

	public string matieumuc { get; set; }

	public string giatoithieu { get; set; }

	public string giatoida { get; set; }

	public string cancu { get; set; }

	public string ngayhlvbpq { get; set; }

	public string cqtgia { get; set; }

	public string cqtsoqd { get; set; }

	public string cqtngayqd { get; set; }

	public string cqthltu { get; set; }

	public string cqthlden { get; set; }

	public bool isCqt { get; set; }

	public string level { get; set; }

	public string mucthu { get; set; }

	public string manhom { get; set; }

	public string tentinh { get; set; }

	public string ghichu { get; set; }
}
