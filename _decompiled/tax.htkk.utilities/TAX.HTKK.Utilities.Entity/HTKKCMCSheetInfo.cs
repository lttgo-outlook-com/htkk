using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCSheetInfo
{
	public SheetView sViewBK;

	public string SheetID { get; set; }

	public string SheetName { get; set; }

	public string OldSheetName { get; set; }

	public string Caption { get; set; }

	public string DataFile { get; set; }

	public string Active { get; set; }

	public short SheetIndex { get; set; }

	public string XMLnode { get; set; }

	public int CoDieuChinhSoLieu { get; set; }

	public bool IsReadDataFile { get; set; }

	public string ShowSearchButton { get; set; }

	public string ShowTaiBangKeButton { get; set; }

	public string strPrintBarcode { get; set; }

	public bool isValidData { get; set; }

	public bool CoDuLieu { get; set; }

	public List<HTKKCMCCellInfo> StaticCells { get; set; }

	public List<HTKKCMCTableInfo> Tables { get; set; }

	public Dictionary<string, string> dicTexttip { get; set; }

	public int ActiveRowIndex { get; set; }

	public int ActiveColIndex { get; set; }

	public string TaiDuLieuLon { get; set; }

	public string TaiDuLieuRatLon { get; set; }

	public Dictionary<string, object> dicVar { get; set; }

	public bool IsEmptySheet { get; set; }

	public HTKKCMCXML XMLTCT { get; set; }

	public string RowTTDLT { get; set; }

	public string RowBreakPage { get; set; }

	public string FileName { get; set; }

	public HTKKCMCSheetInfo()
	{
		SheetID = string.Empty;
		SheetName = string.Empty;
		Caption = string.Empty;
		DataFile = string.Empty;
		Active = string.Empty;
		SheetIndex = 0;
		ActiveColIndex = 0;
		ActiveRowIndex = 0;
		ShowSearchButton = null;
		ShowTaiBangKeButton = null;
		strPrintBarcode = string.Empty;
		StaticCells = new List<HTKKCMCCellInfo>();
		Tables = new List<HTKKCMCTableInfo>();
		XMLnode = string.Empty;
		CoDieuChinhSoLieu = 0;
		IsReadDataFile = true;
		IsEmptySheet = false;
		RowTTDLT = string.Empty;
		RowBreakPage = string.Empty;
		FileName = string.Empty;
		isValidData = true;
		CoDuLieu = false;
		TaiDuLieuLon = string.Empty;
		TaiDuLieuRatLon = string.Empty;
		dicVar = new Dictionary<string, object>();
	}

	public HTKKCMCSheetInfo(HTKKCMCSheetInfo sInfo)
	{
		SheetID = sInfo.SheetID;
		SheetName = sInfo.SheetName;
		Caption = sInfo.Caption;
		DataFile = sInfo.DataFile;
		Active = sInfo.Active;
		SheetIndex = sInfo.SheetIndex;
		ActiveColIndex = sInfo.ActiveColIndex;
		ActiveRowIndex = sInfo.ActiveRowIndex;
		ShowSearchButton = sInfo.ShowSearchButton;
		ShowTaiBangKeButton = sInfo.ShowTaiBangKeButton;
		strPrintBarcode = sInfo.strPrintBarcode;
		StaticCells = sInfo.CloneStaticCells();
		Tables = new List<HTKKCMCTableInfo>();
		List<HTKKCMCTableInfo> tables = sInfo.Tables;
		Action<HTKKCMCTableInfo> action = delegate(HTKKCMCTableInfo item)
		{
			Tables.Add(new HTKKCMCTableInfo(item, 2));
		};
		tables.ForEach(action);
		XMLnode = sInfo.XMLnode;
		CoDieuChinhSoLieu = sInfo.CoDieuChinhSoLieu;
		IsEmptySheet = sInfo.IsEmptySheet;
		RowTTDLT = sInfo.RowTTDLT;
		RowBreakPage = sInfo.RowBreakPage;
		FileName = sInfo.FileName;
		isValidData = sInfo.isValidData;
		CoDuLieu = sInfo.CoDuLieu;
		TaiDuLieuLon = sInfo.TaiDuLieuLon;
		TaiDuLieuRatLon = sInfo.TaiDuLieuRatLon;
		dicVar = sInfo.dicVar;
	}

	public HTKKCMCCellBasicInfo GetCellLocation(string location)
	{
		HTKKCMCCellBasicInfo hTKKCMCCellBasicInfo = new HTKKCMCCellBasicInfo();
		try
		{
			if (location.Length > 0)
			{
				string[] array = location.Split(new char[1] { '_' });
				if (array.Length == 2 && !string.IsNullOrEmpty(array[1]) && !string.IsNullOrEmpty(array[0]))
				{
					hTKKCMCCellBasicInfo.RowIndex = int.Parse(array[1]);
					hTKKCMCCellBasicInfo.ColIndex = HTKKFile.ColumnLettersToIndex(array[0]);
				}
				else
				{
					hTKKCMCCellBasicInfo = null;
				}
			}
			else
			{
				hTKKCMCCellBasicInfo = null;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			hTKKCMCCellBasicInfo = null;
		}
		return hTKKCMCCellBasicInfo;
	}

	public List<HTKKCMCCellInfo> CloneStaticCells()
	{
		List<HTKKCMCCellInfo> reValue = new List<HTKKCMCCellInfo>();
		StaticCells.ForEach(delegate(HTKKCMCCellInfo item)
		{
			reValue.Add(new HTKKCMCCellInfo(item));
		});
		return reValue;
	}

	public List<HTKKCMCCellBasicInfo> GetListCellLocation(string location)
	{
		List<HTKKCMCCellBasicInfo> list = new List<HTKKCMCCellBasicInfo>();
		HTKKCMCCellBasicInfo hTKKCMCCellBasicInfo = new HTKKCMCCellBasicInfo();
		try
		{
			string[] array = Regex.Split(location, "###");
			string[] array2 = array;
			foreach (string location2 in array2)
			{
				hTKKCMCCellBasicInfo = GetCellLocation(location2);
				list.Add(hTKKCMCCellBasicInfo);
			}
			return list;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			list = null;
		}
		return list;
	}

	public HTKKCMCCellInfo GetStaticCellInfo(XmlNode xmlCellNode)
	{
		HTKKCMCCellInfo hTKKCMCCellInfo = new HTKKCMCCellInfo();
		try
		{
			hTKKCMCCellInfo = new HTKKCMCCellInfo();
			hTKKCMCCellInfo.CellTemplate = ((xmlCellNode.Attributes["CellID"] == null) ? null : GetCellLocation(((XmlNode)xmlCellNode.Attributes["CellID"]).Value));
			hTKKCMCCellInfo.CellReport = ((xmlCellNode.Attributes["CellID2"] == null) ? null : GetCellLocation(((XmlNode)xmlCellNode.Attributes["CellID2"]).Value));
			hTKKCMCCellInfo.Value = ((xmlCellNode.Attributes["Value"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Value"]).Value);
			hTKKCMCCellInfo.DefaultValue = ((xmlCellNode.Attributes["DefaultValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["DefaultValue"]).Value);
			hTKKCMCCellInfo.Controltype = ((xmlCellNode.Attributes["Controltype"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Controltype"]).Value);
			hTKKCMCCellInfo.HelpContextID = ((xmlCellNode.Attributes["HelpContextID"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["HelpContextID"]).Value);
			hTKKCMCCellInfo.Path = ((xmlCellNode.Attributes["Path"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Path"]).Value);
			hTKKCMCCellInfo.StatusID = ((xmlCellNode.Attributes["StatusID"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["StatusID"]).Value);
			hTKKCMCCellInfo.PropertyInfo = ((xmlCellNode.Attributes["PropertyInfo"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["PropertyInfo"]).Value);
			hTKKCMCCellInfo.ChildCell = ((xmlCellNode.Attributes["ChildCell"] == null) ? null : GetCellLocation(((XmlNode)xmlCellNode.Attributes["ChildCell"]).Value));
			hTKKCMCCellInfo.ParentCell = ((xmlCellNode.Attributes["ParentCell"] == null) ? null : GetCellLocation(((XmlNode)xmlCellNode.Attributes["ParentCell"]).Value));
			hTKKCMCCellInfo.CellSelectedValue = ((xmlCellNode.Attributes["SelectedValue"] == null) ? null : GetCellLocation(((XmlNode)xmlCellNode.Attributes["SelectedValue"]).Value));
			if (xmlCellNode.Attributes["MinValue"] != null)
			{
				hTKKCMCCellInfo.Validate.MinValue = ((xmlCellNode.Attributes["MinValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MinValue"]).Value);
			}
			if (xmlCellNode.Attributes["MaxValue"] != null)
			{
				hTKKCMCCellInfo.Validate.MaxValue = ((xmlCellNode.Attributes["MaxValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MaxValue"]).Value);
			}
			if (xmlCellNode.Attributes["MaxLen"] != null)
			{
				hTKKCMCCellInfo.Validate.MaxLen = ((xmlCellNode.Attributes["MaxLen"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MaxLen"]).Value);
			}
			if (xmlCellNode.Attributes["LonHon"] != null)
			{
				hTKKCMCCellInfo.Validate.LonHon = ((xmlCellNode.Attributes["LonHon"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["LonHon"]).Value);
			}
			if (xmlCellNode.Attributes["MessageID"] != null)
			{
				hTKKCMCCellInfo.Validate.MessageID = ((xmlCellNode.Attributes["MessageID"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MessageID"]).Value);
			}
			hTKKCMCCellInfo.Encode = ((xmlCellNode.Attributes["Encode"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Encode"]).Value);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			hTKKCMCCellInfo = null;
		}
		return hTKKCMCCellInfo;
	}

	public HTKKCMCTableCellInfo GetRowInfo(XmlNode xmlCellNode)
	{
		HTKKCMCTableCellInfo hTKKCMCTableCellInfo = new HTKKCMCTableCellInfo();
		try
		{
			hTKKCMCTableCellInfo.CellTemplate.ColIndex = ((xmlCellNode.Attributes["CellID"] == null) ? (-1) : HTKKFile.ColumnLettersToIndex(((XmlNode)xmlCellNode.Attributes["CellID"]).Value));
			hTKKCMCTableCellInfo.CellReport.ColIndex = ((xmlCellNode.Attributes["CellID2"] == null) ? (-1) : HTKKFile.ColumnLettersToIndex(((XmlNode)xmlCellNode.Attributes["CellID2"]).Value));
			hTKKCMCTableCellInfo.CellBangKe.ColIndex = ((xmlCellNode.Attributes["CellID3"] == null) ? (-1) : HTKKFile.ColumnLettersToIndex(((XmlNode)xmlCellNode.Attributes["CellID3"]).Value));
			hTKKCMCTableCellInfo.Controltype = ((xmlCellNode.Attributes["Controltype"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Controltype"]).Value);
			hTKKCMCTableCellInfo.HelpContextID = ((xmlCellNode.Attributes["HelpContextID"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["HelpContextID"]).Value);
			hTKKCMCTableCellInfo.ParentValue = ((xmlCellNode.Attributes["ParentValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["ParentValue"]).Value);
			hTKKCMCTableCellInfo.ChildValue = ((xmlCellNode.Attributes["ChildValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["ChildValue"]).Value);
			hTKKCMCTableCellInfo.indexColumnSelectedValue = ((xmlCellNode.Attributes["SelectedValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["SelectedValue"]).Value);
			hTKKCMCTableCellInfo.Path = ((xmlCellNode.Attributes["Path"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Path"]).Value);
			hTKKCMCTableCellInfo.DefaultValue = ((xmlCellNode.Attributes["DefaultValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["DefaultValue"]).Value);
			hTKKCMCTableCellInfo.StatusID = ((xmlCellNode.Attributes["StatusID"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["StatusID"]).Value);
			hTKKCMCTableCellInfo.Validate.MinValue = ((xmlCellNode.Attributes["MinValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MinValue"]).Value);
			hTKKCMCTableCellInfo.Validate.MaxValue = ((xmlCellNode.Attributes["MaxValue"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MaxValue"]).Value);
			hTKKCMCTableCellInfo.Validate.MaxLen = ((xmlCellNode.Attributes["MaxLen"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MaxLen"]).Value);
			hTKKCMCTableCellInfo.Validate.NhoHon = ((xmlCellNode.Attributes["NhoHon"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["NhoHon"]).Value);
			hTKKCMCTableCellInfo.Validate.LonHon = ((xmlCellNode.Attributes["LonHon"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["LonHon"]).Value);
			hTKKCMCTableCellInfo.Validate.MessageID = ((xmlCellNode.Attributes["MessageID"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["MessageID"]).Value);
			hTKKCMCTableCellInfo.Encode = ((xmlCellNode.Attributes["Encode"] == null) ? null : ((XmlNode)xmlCellNode.Attributes["Encode"]).Value);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			hTKKCMCTableCellInfo = null;
		}
		return hTKKCMCTableCellInfo;
	}

	public HTKKCMCDataLocation GetTableLocation(XmlNode elemLocations)
	{
		HTKKCMCDataLocation hTKKCMCDataLocation = new HTKKCMCDataLocation();
		try
		{
			XmlNode val = elemLocations.SelectSingleNode("DataLocationTemplateRowStart");
			hTKKCMCDataLocation.DataLocationTemplateRowStart = int.Parse(val.InnerText);
			val = elemLocations.SelectSingleNode("DataLocationReportRowStart");
			hTKKCMCDataLocation.DataLocationReportRowStart = int.Parse(val.InnerText);
			val = elemLocations.SelectSingleNode("DataLocationTaiBangKeRowStart");
			hTKKCMCDataLocation.DataLocationTaiBangKeRowStart = int.Parse(val.InnerText);
			val = elemLocations.SelectSingleNode("TableTemplateRowStart");
			hTKKCMCDataLocation.TableTemplateRowStart = int.Parse(val.InnerText);
			val = elemLocations.SelectSingleNode("TableReportRowStart");
			hTKKCMCDataLocation.TableReportRowStart = int.Parse(val.InnerText);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			hTKKCMCDataLocation = null;
		}
		return hTKKCMCDataLocation;
	}

	public HTKKCMCTableInfo GetTable(XmlNode SectionNode)
	{
		//IL_0223: Unknown result type (might be due to invalid IL or missing references)
		//IL_022a: Expected O, but got Unknown
		HTKKCMCTableInfo hTKKCMCTableInfo = new HTKKCMCTableInfo();
		try
		{
			XmlNode elemLocations = SectionNode.SelectSingleNode("Location");
			hTKKCMCTableInfo.Location = GetTableLocation(elemLocations);
			XmlNode val = SectionNode.SelectSingleNode("Header");
			if (val != null)
			{
				HTKKCMCStaticOfTable hTKKCMCStaticOfTable = new HTKKCMCStaticOfTable();
				hTKKCMCStaticOfTable.Length = ((val.Attributes["Length"] != null) ? int.Parse(((XmlNode)val.Attributes["Length"]).Value) : 0);
				for (int i = 0; i < val.ChildNodes.Count; i++)
				{
					if (val.ChildNodes[i].Name == "Cell")
					{
						hTKKCMCStaticOfTable.StaticCells.Add(GetStaticCellInfo(val.ChildNodes[i]));
					}
				}
				hTKKCMCTableInfo.Header = hTKKCMCStaticOfTable;
			}
			XmlNode val2 = SectionNode.SelectSingleNode("RowInfo");
			if (val2 != null)
			{
				HTKKCMCRowInfo hTKKCMCRowInfo = new HTKKCMCRowInfo();
				hTKKCMCRowInfo.RowDefaultDisplay = ((val2.Attributes["RowDefaultDisplay"] == null) ? "1" : ((XmlNode)val2.Attributes["RowDefaultDisplay"]).Value);
				hTKKCMCRowInfo.RowPath = ((val2.Attributes["RowPath"] == null) ? null : ((XmlNode)val2.Attributes["RowPath"]).Value);
				for (int i = 0; i < val2.ChildNodes.Count; i++)
				{
					if (val2.ChildNodes[i].Name == "Cell")
					{
						HTKKCMCTableCellInfo rowInfo = GetRowInfo(val2.ChildNodes[i]);
						hTKKCMCRowInfo.Cells.Add(rowInfo);
					}
				}
				hTKKCMCTableInfo.RowsInfo = hTKKCMCRowInfo;
			}
			XmlNodeList val3 = SectionNode.SelectNodes("RowData");
			if (val3 != null)
			{
				foreach (XmlNode item in val3)
				{
					XmlNode xmlRowDataNode = item;
					hTKKCMCTableInfo.RowsValues.Add(GetDataFromRowDataNode(xmlRowDataNode));
				}
			}
			XmlNode val4 = SectionNode.SelectSingleNode("Footer");
			if (val4 != null)
			{
				HTKKCMCStaticOfTable hTKKCMCStaticOfTable2 = new HTKKCMCStaticOfTable();
				hTKKCMCStaticOfTable2.Length = ((val4.Attributes["Length"] != null) ? int.Parse(((XmlNode)val4.Attributes["Length"]).Value) : 0);
				for (int i = 0; i < val4.ChildNodes.Count; i++)
				{
					if (val4.ChildNodes[i].Name == "Cell")
					{
						hTKKCMCStaticOfTable2.StaticCells.Add(GetStaticCellInfo(val4.ChildNodes[i]));
					}
				}
				hTKKCMCTableInfo.Footer = hTKKCMCStaticOfTable2;
			}
			hTKKCMCTableInfo.TableID = ((XmlNode)SectionNode.Attributes["id"]).Value;
			hTKKCMCTableInfo.TablePath = ((XmlNode)SectionNode.Attributes["TablePath"]).Value;
			hTKKCMCTableInfo.iColumnSTT = ((SectionNode.Attributes["iColumnSTT"] == null) ? (-1) : int.Parse(((XmlNode)SectionNode.Attributes["iColumnSTT"]).Value));
			hTKKCMCTableInfo.iColumnSTTReport = ((SectionNode.Attributes["iColumnSTTReport"] == null) ? (-1) : int.Parse(((XmlNode)SectionNode.Attributes["iColumnSTTReport"]).Value));
			hTKKCMCTableInfo.CellTableSTT = ((SectionNode.Attributes["TableSTTLocation"] == null) ? null : GetCellLocation(((XmlNode)SectionNode.Attributes["TableSTTLocation"]).Value));
			if (hTKKCMCTableInfo.CellTableSTT != null)
			{
				hTKKCMCTableInfo.CellTableSTT.Value = ((SectionNode.Attributes["TableSTTValue"] == null) ? "" : ((XmlNode)SectionNode.Attributes["TableSTTValue"]).Value);
			}
			hTKKCMCTableInfo.CellTableReportSTT = ((SectionNode.Attributes["TableReportSTTLocation"] == null) ? null : GetCellLocation(((XmlNode)SectionNode.Attributes["TableReportSTTLocation"]).Value));
			if (hTKKCMCTableInfo.CellTableReportSTT != null)
			{
				hTKKCMCTableInfo.CellTableReportSTT.Value = hTKKCMCTableInfo.CellTableSTT.Value;
			}
			hTKKCMCTableInfo.TableName = ((SectionNode.Attributes["TableName"] == null) ? null : ((XmlNode)SectionNode.Attributes["TableName"]).Value);
			hTKKCMCTableInfo.TableLength = ((SectionNode.Attributes["TableLength"] != null) ? int.Parse(((XmlNode)SectionNode.Attributes["TableLength"]).Value) : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			hTKKCMCTableInfo = null;
		}
		return hTKKCMCTableInfo;
	}

	public bool GetDetailSheetInfo(string infaceTemplatePath)
	{
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Expected O, but got Unknown
		//IL_00ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Expected O, but got Unknown
		bool result = false;
		try
		{
			if (!File.Exists(infaceTemplatePath))
			{
				return result;
			}
			HTKKCMCXML hTKKCMCXML = new HTKKCMCXML(infaceTemplatePath);
			XmlNodeList val = ((((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections") != null) ? ((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections").ChildNodes : null);
			if (val != null && val.Count > 0)
			{
				foreach (XmlNode item in val)
				{
					XmlNode val2 = item;
					if (!(val2.Name == "Section"))
					{
						continue;
					}
					if (((XmlNode)val2.Attributes["Dynamic"]).Value == "0")
					{
						XmlNodeList childNodes = val2.SelectSingleNode("Cells").ChildNodes;
						foreach (XmlNode item2 in childNodes)
						{
							XmlNode xmlCellNode = item2;
							StaticCells.Add(GetStaticCellInfo(xmlCellNode));
						}
					}
					else
					{
						Tables.Add(GetTable(val2));
					}
				}
			}
			result = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public bool GetStaticSheetInfoNoTable(string infaceTemplatePath)
	{
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Expected O, but got Unknown
		//IL_00ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Expected O, but got Unknown
		bool result = false;
		try
		{
			if (!File.Exists(infaceTemplatePath))
			{
				return result;
			}
			HTKKCMCXML hTKKCMCXML = new HTKKCMCXML(infaceTemplatePath);
			XmlNodeList val = ((((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections") != null) ? ((XmlNode)hTKKCMCXML.XDocument).SelectSingleNode("/Sections").ChildNodes : null);
			if (val != null && val.Count > 0)
			{
				foreach (XmlNode item in val)
				{
					XmlNode val2 = item;
					if (!(val2.Name == "Section") || !(((XmlNode)val2.Attributes["Dynamic"]).Value == "0"))
					{
						continue;
					}
					XmlNodeList childNodes = val2.SelectSingleNode("Cells").ChildNodes;
					foreach (XmlNode item2 in childNodes)
					{
						XmlNode xmlCellNode = item2;
						StaticCells.Add(GetStaticCellInfo(xmlCellNode));
					}
				}
			}
			result = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = false;
		}
		return result;
	}

	public HTKKCMCTableCellInfo GetTableCellInfo(int iRow, int iCol)
	{
		HTKKCMCTableCellInfo result = new HTKKCMCTableCellInfo();
		try
		{
			for (int i = 0; i < Tables.Count; i++)
			{
				if (!((iRow >= Tables[i].Location.DataLocationTemplateRowStart) & (iRow <= Tables[i].Location.DataLocationTemplateRowStart + Tables[i].RowsValues.Count)))
				{
					continue;
				}
				for (int j = 0; j < Tables[i].RowsInfo.Cells.Count; j++)
				{
					if (Tables[i].RowsInfo.Cells[j].CellTemplate.ColIndex == iCol)
					{
						return Tables[i].RowsInfo.Cells[j];
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return result;
	}

	public int GetTableIndex(int iRow, int iCol)
	{
		int result = -1;
		try
		{
			for (int i = 0; i < Tables.Count; i++)
			{
				if (Tables[i].Location.TableTemplateRowStart > -1)
				{
					if ((iRow >= Tables[i].Location.TableTemplateRowStart) & (iRow < Tables[i].Location.TableTemplateRowStart + Tables[i].RowsValues.Count + Tables[i].TableLength))
					{
						result = i;
						break;
					}
					if ((iRow >= Tables[i].Location.DataLocationTemplateRowStart) & (iRow < Tables[i].Location.DataLocationTemplateRowStart + Tables[i].RowsValues.Count))
					{
						result = i;
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
		return result;
	}

	public HTKKCMCTableInfo GetTable(string tableID, string tableName)
	{
		HTKKCMCTableInfo result = null;
		try
		{
			for (int i = 0; i < Tables.Count; i++)
			{
				if (tableID == Tables[i].TableID && tableName == Tables[i].TableName)
				{
					result = Tables[i];
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return result;
	}

	public List<string> GetDataFromStringRowDataValues(string strRowDataValues, string strSeparate)
	{
		List<string> list = new List<string>();
		try
		{
			string[] array = Regex.Split(strRowDataValues, strSeparate);
			string[] array2 = array;
			foreach (string item in array2)
			{
				list.Add(item);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return list;
	}

	public HTKKCMCRowValue GetDataFromRowDataNode(XmlNode xmlRowDataNode)
	{
		HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
		try
		{
			hTKKCMCRowValue.RowDisplay = ((xmlRowDataNode.Attributes["RowDisplay"] == null) ? "1" : ((XmlNode)xmlRowDataNode.Attributes["RowDisplay"]).Value);
			hTKKCMCRowValue.RowValue = GetDataFromStringRowDataValues(xmlRowDataNode.InnerText, "###");
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return hTKKCMCRowValue;
	}

	public void IncrementStaticCell(long incrementPosition, int Distance)
	{
		try
		{
			foreach (HTKKCMCCellInfo staticCell in StaticCells)
			{
				if (staticCell.CellTemplate.RowIndex > incrementPosition)
				{
					staticCell.Increment(Distance);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
