namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCTableInfoABC
{
	public int TableLength { get; set; }

	public HTKKCMCTableInfoABC(HTKKCMCTableInfo table)
	{
		int num = table.TableLength + table.RowsValues.Count;
		TableLength = table.TableLength;
	}
}
