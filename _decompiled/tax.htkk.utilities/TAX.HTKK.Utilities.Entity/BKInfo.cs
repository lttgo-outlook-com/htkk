using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class BKInfo
{
	public Dictionary<string, string> DicSetData { get; set; }

	public List<CellAttributeInfo> ListCellAttribute { get; set; }

	public BKInfo()
	{
		DicSetData = new Dictionary<string, string>();
		ListCellAttribute = new List<CellAttributeInfo>();
	}
}
