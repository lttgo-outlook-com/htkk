using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKToKhaiInfo
{
	public string SignNode { get; set; }

	public string MaToKhai { get; set; }

	public string ParentID { get; set; }

	public string Title { get; set; }

	public string MoTaBMau { get; set; }

	public string FolderName { get; set; }

	public string Version { get; set; }

	public string InterfaceTemplateExcel { get; set; }

	public string InterfaceTemplateXml { get; set; }

	public string ReportTemplates { get; set; }

	public string Caption { get; set; }

	public string InterfaceIni { get; set; }

	public string StartDate { get; set; }

	public string EndDate { get; set; }

	public Dictionary<string, string> ListConfigCells { get; set; }

	public HTKKToKhaiInfo()
	{
		SignNode = string.Empty;
		MaToKhai = string.Empty;
		ParentID = string.Empty;
		FolderName = string.Empty;
		Title = string.Empty;
		MoTaBMau = string.Empty;
		Version = string.Empty;
		InterfaceTemplateExcel = string.Empty;
		InterfaceTemplateXml = string.Empty;
		ReportTemplates = string.Empty;
		Caption = string.Empty;
		InterfaceIni = string.Empty;
		StartDate = string.Empty;
		EndDate = string.Empty;
		ListConfigCells = new Dictionary<string, string>();
	}
}
