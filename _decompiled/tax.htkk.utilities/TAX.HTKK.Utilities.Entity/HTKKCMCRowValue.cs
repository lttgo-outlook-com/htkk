using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCRowValue
{
	public List<string> RowValue;

	public string RowDisplay { get; set; }

	public HTKKCMCRowValue()
	{
		RowValue = new List<string>();
		RowDisplay = "1";
	}

	public HTKKCMCRowValue(HTKKCMCRowValue rowValue)
	{
		RowValue = new List<string>(rowValue.RowValue);
		RowDisplay = rowValue.RowDisplay;
	}
}
