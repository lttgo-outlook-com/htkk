namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCCellInfo
{
	public string Controltype { get; set; }

	public string HelpContextID { get; set; }

	public string StatusID { get; set; }

	public string Path { get; set; }

	public string Value { get; set; }

	public string DefaultValue { get; set; }

	public string PropertyInfo { get; set; }

	public string Formula { get; set; }

	public HTKKCMCValidate Validate { get; set; }

	public HTKKCMCCellBasicInfo CellTemplate { get; set; }

	public HTKKCMCCellBasicInfo CellReport { get; set; }

	public HTKKCMCCellBasicInfo CellSelectedValue { get; set; }

	public HTKKCMCCellBasicInfo ParentCell { get; set; }

	public HTKKCMCCellBasicInfo ChildCell { get; set; }

	public string Encode { get; set; }

	public HTKKCMCCellInfo()
	{
		Controltype = string.Empty;
		HelpContextID = string.Empty;
		StatusID = string.Empty;
		Path = string.Empty;
		Value = string.Empty;
		DefaultValue = string.Empty;
		PropertyInfo = string.Empty;
		Formula = string.Empty;
		CellTemplate = new HTKKCMCCellBasicInfo();
		CellReport = new HTKKCMCCellBasicInfo();
		CellSelectedValue = new HTKKCMCCellBasicInfo();
		ChildCell = new HTKKCMCCellBasicInfo();
		ParentCell = new HTKKCMCCellBasicInfo();
		Validate = new HTKKCMCValidate();
		Encode = string.Empty;
	}

	public void Increment(int distance)
	{
		CellTemplate.RowIndex += distance;
		if (CellReport != null)
		{
			CellReport.RowIndex += distance;
		}
		if (ChildCell != null)
		{
			ChildCell.RowIndex += distance;
		}
		if (ParentCell != null)
		{
			ParentCell.RowIndex += distance;
		}
		if (CellSelectedValue != null)
		{
			CellSelectedValue.RowIndex += distance;
		}
	}

	public HTKKCMCCellInfo(HTKKCMCCellInfo cellInfo)
	{
		Controltype = cellInfo.Controltype;
		HelpContextID = cellInfo.HelpContextID;
		StatusID = cellInfo.StatusID;
		Path = cellInfo.Path;
		Value = cellInfo.Value;
		DefaultValue = ((cellInfo.DefaultValue != null) ? cellInfo.DefaultValue : null);
		PropertyInfo = cellInfo.PropertyInfo;
		CellTemplate = new HTKKCMCCellBasicInfo(cellInfo.CellTemplate);
		Formula = cellInfo.Formula;
		if (cellInfo.CellReport != null)
		{
			CellReport = new HTKKCMCCellBasicInfo(cellInfo.CellReport);
		}
		else
		{
			CellReport = null;
		}
		if (cellInfo.CellSelectedValue != null)
		{
			CellSelectedValue = new HTKKCMCCellBasicInfo(cellInfo.CellSelectedValue);
		}
		else
		{
			CellSelectedValue = null;
		}
		if (cellInfo.ChildCell != null)
		{
			ChildCell = new HTKKCMCCellBasicInfo(cellInfo.ChildCell);
		}
		else
		{
			ChildCell = null;
		}
		if (cellInfo.ParentCell != null)
		{
			ParentCell = new HTKKCMCCellBasicInfo(cellInfo.ParentCell);
		}
		else
		{
			ParentCell = null;
		}
		if (cellInfo.Validate != null)
		{
			Validate = new HTKKCMCValidate(cellInfo.Validate);
		}
		else
		{
			Validate = null;
		}
		Encode = cellInfo.Encode;
	}
}
