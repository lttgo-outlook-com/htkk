using System;
using System.Collections.Generic;

namespace HTKK;

public class HTKKCatalogues
{
	public string InterfaceTemplate { get; set; }

	public string DataFile { get; set; }

	public DateTime StartDate { get; set; }

	public List<HTKKCataloguesSheet> catalogSheets { get; set; }
}
