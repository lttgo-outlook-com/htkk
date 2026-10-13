using FarPoint.Win.Spread.CellType;

namespace TAX.HTKK.Utilities.Core;

public class CellTypeKHBS
{
	public int sheetIndex { get; set; }

	public int row { get; set; }

	public int col { get; set; }

	public ICellType iCellType { get; set; }

	public object value { get; set; }
}
