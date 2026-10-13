namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCCellBasicInfo
{
	public int ColIndex { get; set; }

	public int RowIndex { get; set; }

	public string Value { get; set; }

	public HTKKCMCCellBasicInfo()
	{
		ColIndex = 0;
		RowIndex = 0;
		Value = string.Empty;
	}

	public HTKKCMCCellBasicInfo(HTKKCMCCellBasicInfo cellBasicInfo)
	{
		ColIndex = cellBasicInfo.ColIndex;
		RowIndex = cellBasicInfo.RowIndex;
		Value = cellBasicInfo.Value;
	}
}
