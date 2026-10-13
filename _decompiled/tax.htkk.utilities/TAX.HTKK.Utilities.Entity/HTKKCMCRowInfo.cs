using System;
using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCRowInfo
{
	public List<HTKKCMCTableCellInfo> Cells { get; set; }

	public string RowDefaultDisplay { get; set; }

	public string RowPath { get; set; }

	public HTKKCMCRowInfo()
	{
		Cells = new List<HTKKCMCTableCellInfo>();
		RowDefaultDisplay = "1";
		RowPath = "";
	}

	public HTKKCMCRowInfo(HTKKCMCRowInfo rowInfo)
	{
		Cells = new List<HTKKCMCTableCellInfo>();
		List<HTKKCMCTableCellInfo> cells = rowInfo.Cells;
		Action<HTKKCMCTableCellInfo> action = delegate(HTKKCMCTableCellInfo item)
		{
			Cells.Add(new HTKKCMCTableCellInfo(item));
		};
		cells.ForEach(action);
		RowDefaultDisplay = rowInfo.RowDefaultDisplay;
		RowPath = rowInfo.RowPath;
	}
}
