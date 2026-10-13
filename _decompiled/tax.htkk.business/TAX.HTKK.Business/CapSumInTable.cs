using System.Collections.Generic;

namespace TAX.HTKK.Business;

public class CapSumInTable
{
	public string Cap { get; set; }

	public int RowSum { get; set; }

	public List<int> RowsChilden { get; set; }

	public string Value { get; set; }

	public string ColsSum { get; set; }
}
