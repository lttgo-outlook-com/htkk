using System.Drawing;
using System.Windows.Forms;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;

namespace HTKK.Class;

public class TaxComboBoxCellType : ComboBoxCellType
{
	public override object IsReservedLocation(Graphics g, int x, int y, Rectangle rc, Appearance appearance, object value, float zoomFactor)
	{
		return null;
	}

	public override Cursor GetReservedCursor(object o)
	{
		if (o != null)
		{
			return null;
		}
		return null;
	}
}
