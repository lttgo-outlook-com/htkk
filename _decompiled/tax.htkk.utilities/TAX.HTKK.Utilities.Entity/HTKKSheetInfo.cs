namespace TAX.HTKK.Utilities.Entity;

public class HTKKSheetInfo
{
	public string SheetID { get; set; }

	public string SheetName { get; set; }

	public string Caption { get; set; }

	public string DataFile { get; set; }

	public string Active { get; set; }

	public short SheetIndex { get; set; }

	public string DynamicData { get; set; }

	public string ShowSearchButton { get; set; }

	public string ButtonCaption { get; set; }

	public BKInfo BKInfoDetail { get; set; }

	public string RowBreakPage { get; set; }

	public string TotalRowFooter { get; set; }

	public HTKKSheetInfo()
	{
		SheetID = string.Empty;
		SheetName = string.Empty;
		Caption = string.Empty;
		DataFile = string.Empty;
		Active = string.Empty;
		DynamicData = string.Empty;
		SheetIndex = 0;
		RowBreakPage = string.Empty;
		ShowSearchButton = string.Empty;
		ButtonCaption = string.Empty;
		TotalRowFooter = string.Empty;
	}
}
