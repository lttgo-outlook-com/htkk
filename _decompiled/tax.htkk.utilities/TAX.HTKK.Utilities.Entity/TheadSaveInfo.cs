using System;
using System.Collections.Generic;
using System.Xml;

namespace TAX.HTKK.Utilities.Entity;

public class TheadSaveInfo
{
	public DateTime ThreadCreated { get; set; }

	public int ThreadId { get; set; }

	public int RowStart { get; set; }

	public int RowEnd { get; set; }

	public XmlDocument XmlChild { get; set; }

	public Dictionary<int, string> DictData { get; set; }

	public TheadSaveInfo()
	{
		ThreadCreated = DateTime.Now;
		ThreadId = 0;
		RowStart = 0;
		RowEnd = 0;
		DictData = new Dictionary<int, string>();
	}
}
