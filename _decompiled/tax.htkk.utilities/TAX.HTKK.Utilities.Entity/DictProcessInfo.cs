using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class DictProcessInfo
{
	public int Status { get; set; }

	public HTKKSheetInfo SheetInfo { get; set; }

	public Dictionary<string, List<DictInfo>> DictData { get; set; }

	public DictProcessInfo()
	{
		Status = 0;
		SheetInfo = new HTKKSheetInfo();
		DictData = new Dictionary<string, List<DictInfo>>();
	}
}
