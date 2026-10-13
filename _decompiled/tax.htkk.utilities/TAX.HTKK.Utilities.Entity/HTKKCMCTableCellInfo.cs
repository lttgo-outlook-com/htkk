namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCTableCellInfo
{
	public string Controltype { get; set; }

	public string HelpContextID { get; set; }

	public string ParentValue { get; set; }

	public string ChildValue { get; set; }

	public string StatusID { get; set; }

	public string Path { get; set; }

	public string DefaultValue { get; set; }

	public HTKKCMCValidate Validate { get; set; }

	public string Formula { get; set; }

	public HTKKCMCCellBasicInfo CellTemplate { get; set; }

	public HTKKCMCCellBasicInfo CellReport { get; set; }

	public HTKKCMCCellBasicInfo CellBangKe { get; set; }

	public string indexColumnSelectedValue { get; set; }

	public string Encode { get; set; }

	public HTKKCMCTableCellInfo()
	{
		Controltype = string.Empty;
		HelpContextID = string.Empty;
		ParentValue = string.Empty;
		ChildValue = string.Empty;
		StatusID = string.Empty;
		Path = string.Empty;
		Formula = string.Empty;
		DefaultValue = string.Empty;
		indexColumnSelectedValue = string.Empty;
		CellTemplate = new HTKKCMCCellBasicInfo();
		CellReport = new HTKKCMCCellBasicInfo();
		CellBangKe = new HTKKCMCCellBasicInfo();
		Validate = new HTKKCMCValidate();
		Encode = string.Empty;
	}

	public HTKKCMCTableCellInfo(HTKKCMCTableCellInfo tableCellInfo)
	{
		Controltype = tableCellInfo.Controltype;
		HelpContextID = tableCellInfo.HelpContextID;
		ParentValue = tableCellInfo.ParentValue;
		ChildValue = tableCellInfo.ChildValue;
		StatusID = tableCellInfo.StatusID;
		Path = tableCellInfo.Path;
		DefaultValue = tableCellInfo.DefaultValue;
		indexColumnSelectedValue = tableCellInfo.indexColumnSelectedValue;
		CellTemplate = new HTKKCMCCellBasicInfo(tableCellInfo.CellTemplate);
		Formula = tableCellInfo.Formula;
		CellReport = new HTKKCMCCellBasicInfo(tableCellInfo.CellReport);
		CellBangKe = new HTKKCMCCellBasicInfo(tableCellInfo.CellBangKe);
		Validate = new HTKKCMCValidate(tableCellInfo.Validate);
		Encode = tableCellInfo.Encode;
	}
}
