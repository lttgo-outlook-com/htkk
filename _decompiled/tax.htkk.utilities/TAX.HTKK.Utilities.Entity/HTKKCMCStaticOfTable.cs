using System;
using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCStaticOfTable
{
	public List<HTKKCMCCellInfo> StaticCells { get; set; }

	public int Length { get; set; }

	public HTKKCMCStaticOfTable()
	{
		StaticCells = new List<HTKKCMCCellInfo>();
		Length = 0;
	}

	public HTKKCMCStaticOfTable(HTKKCMCStaticOfTable sTable)
	{
		StaticCells = new List<HTKKCMCCellInfo>();
		List<HTKKCMCCellInfo> staticCells = sTable.StaticCells;
		Action<HTKKCMCCellInfo> action = delegate(HTKKCMCCellInfo item)
		{
			StaticCells.Add(new HTKKCMCCellInfo(item));
		};
		staticCells.ForEach(action);
		Length = sTable.Length;
	}
}
