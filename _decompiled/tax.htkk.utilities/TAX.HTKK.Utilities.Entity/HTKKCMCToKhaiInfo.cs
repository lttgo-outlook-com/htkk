using System.Collections.Generic;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCToKhaiInfo
{
	public string TenToKhai { get; set; }

	public string MaFile { get; set; }

	public string MaToKhai { get; set; }

	public string VersionXML { get; set; }

	public string VersionKBSXML { get; set; }

	public string VersionXMLOld { get; set; }

	public string SignNode { get; set; }

	public string MenuID { get; set; }

	public string ParentID { get; set; }

	public string MoTaBieuMau { get; set; }

	public string FolderName { get; set; }

	public string Version { get; set; }

	public string InterfaceTemplateExcel { get; set; }

	public string InterfaceTemplateValidate { get; set; }

	public string InterfaceTemplateXml { get; set; }

	public string InterfaceTemplateKBSXml { get; set; }

	public string XsdFileName { get; set; }

	public string InterfaceIni { get; set; }

	public string ReportTemplates { get; set; }

	public string Caption { get; set; }

	public string StartDate { get; set; }

	public string EndDate { get; set; }

	public string FileNameBS { get; set; }

	public string DataFileNameSaveExtend { get; set; }

	public string FileNameExport { get; set; }

	public string UseDate { get; set; }

	public string KieuKHBS { get; set; }

	public string NopTheoTinhCQT { get; set; }

	public string MaQuyetToan { get; set; }

	public string TenQuyetToan { get; set; }

	public string TaiDuLieuLon { get; set; }

	public string TaiDuLieuRatLon { get; set; }

	public List<HTKKCMCSheetInfo> SheetInfos { get; set; }

	public Dictionary<object, object> dicObject { get; set; }

	public HTKKCMCToKhaiInfo()
	{
		TenToKhai = string.Empty;
		MaFile = string.Empty;
		MaToKhai = string.Empty;
		VersionXML = string.Empty;
		VersionXMLOld = string.Empty;
		SignNode = string.Empty;
		MenuID = string.Empty;
		ParentID = string.Empty;
		FolderName = string.Empty;
		MoTaBieuMau = string.Empty;
		Version = string.Empty;
		InterfaceTemplateExcel = string.Empty;
		InterfaceTemplateValidate = string.Empty;
		InterfaceTemplateXml = string.Empty;
		XsdFileName = string.Empty;
		ReportTemplates = string.Empty;
		Caption = string.Empty;
		InterfaceIni = string.Empty;
		StartDate = string.Empty;
		EndDate = string.Empty;
		DataFileNameSaveExtend = string.Empty;
		FileNameExport = string.Empty;
		FileNameBS = string.Empty;
		UseDate = string.Empty;
		KieuKHBS = string.Empty;
		NopTheoTinhCQT = string.Empty;
		MaQuyetToan = string.Empty;
		TenQuyetToan = string.Empty;
		SheetInfos = new List<HTKKCMCSheetInfo>();
		dicObject = new Dictionary<object, object>();
		TaiDuLieuLon = string.Empty;
		TaiDuLieuRatLon = string.Empty;
	}

	public void RemoveAllSheet()
	{
		for (int num = SheetInfos.Count - 1; num >= 0; num--)
		{
			SheetInfos.RemoveAt(num);
		}
	}
}
