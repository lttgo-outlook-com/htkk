using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCBigDataInfo
{
	public List<HTKKCMCBigTable> listTable { get; set; }

	public bool showFooter { get; set; }

	public List<HTKKCMCCellInfo> staticCells { get; set; }
}
