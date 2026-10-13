using System;
using System.Reflection;
using System.Xml;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCXML
{
	public string PathFile { get; set; }

	public XmlDocument XDocument { get; set; }

	public HTKKCMCXML(string pathFile)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_001b: Expected O, but got Unknown
		PathFile = pathFile;
		XDocument = new XmlDocument();
		XDocument.Load(PathFile);
	}

	public void BindDataStaticTableToXMLNode(SheetView sView, XmlNode xmlStaticTableNode, HTKKCMCStaticOfTable staticOfTable, int intDistanctTable)
	{
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected O, but got Unknown
		if (xmlStaticTableNode == null)
		{
			return;
		}
		XmlNodeList childNodes = xmlStaticTableNode.ChildNodes;
		foreach (XmlNode item in childNodes)
		{
			XmlNode xmlCelNode = item;
			BindDataToStaticOfTableNode(sView, staticOfTable, xmlCelNode, intDistanctTable);
		}
	}

	public void BindDataToStaticOfTableNode(SheetView sView, HTKKCMCStaticOfTable staticOfTable, XmlNode xmlCelNode, int intDistanctTable)
	{
		try
		{
			string value = ((XmlNode)xmlCelNode.Attributes["CellID"]).Value;
			string value2 = ((XmlNode)xmlCelNode.Attributes["CellID2"]).Value;
			string text = "";
			string text2 = "";
			string text3 = "";
			if (xmlCelNode.Attributes["SelectedValue"] != null)
			{
				text = ((XmlNode)xmlCelNode.Attributes["SelectedValue"]).Value;
			}
			if (xmlCelNode.Attributes["ChildCell"] != null)
			{
				text2 = ((XmlNode)xmlCelNode.Attributes["ChildCell"]).Value;
			}
			if (xmlCelNode.Attributes["ParentCell"] != null)
			{
				text3 = ((XmlNode)xmlCelNode.Attributes["ParentCell"]).Value;
			}
			if (value.Length <= 0)
			{
				return;
			}
			string[] array = value.Split(new char[1] { '_' });
			string[] array2 = value2.Split(new char[1] { '_' });
			string[] array3 = null;
			string[] array4 = null;
			string[] array5 = null;
			if (text != null && text != "")
			{
				array3 = text.Split(new char[1] { '_' });
			}
			if (text2 != null && text2 != "")
			{
				array4 = text2.Split(new char[1] { '_' });
			}
			if (text3 != null && text3 != "")
			{
				array5 = text3.Split(new char[1] { '_' });
			}
			if (array.Length != 2)
			{
				return;
			}
			int num = -1;
			int num2 = -1;
			num2 = int.Parse(array[1]);
			num = HTKKFile.ColumnLettersToIndex(array[0]);
			if (num < 0 || num2 < 0)
			{
				return;
			}
			for (int i = 0; i < staticOfTable.StaticCells.Count; i++)
			{
				if (num2 + intDistanctTable == staticOfTable.StaticCells[i].CellTemplate.RowIndex && num == staticOfTable.StaticCells[i].CellTemplate.ColIndex)
				{
					if (sView.Cells[num2 + intDistanctTable - 1, num - 1].Value != null)
					{
						((XmlNode)xmlCelNode.Attributes["Value"]).Value = sView.Cells[num2 + intDistanctTable - 1, num - 1].Value.ToString();
					}
					else
					{
						((XmlNode)xmlCelNode.Attributes["Value"]).Value = staticOfTable.StaticCells[i].DefaultValue;
					}
					((XmlNode)xmlCelNode.Attributes["CellID"]).Value = array[0] + "_" + staticOfTable.StaticCells[i].CellTemplate.RowIndex;
					((XmlNode)xmlCelNode.Attributes["CellID2"]).Value = array2[0] + "_" + staticOfTable.StaticCells[i].CellReport.RowIndex;
					if (array3 != null && array3.Length == 2)
					{
						((XmlNode)xmlCelNode.Attributes["SelectedValue"]).Value = array3[0] + "_" + staticOfTable.StaticCells[i].CellSelectedValue.RowIndex;
					}
					if (array4 != null && array4.Length == 2)
					{
						((XmlNode)xmlCelNode.Attributes["ChildCell"]).Value = array4[0] + "_" + staticOfTable.StaticCells[i].ChildCell.RowIndex;
					}
					if (array5 != null && array5.Length == 2)
					{
						((XmlNode)xmlCelNode.Attributes["ParentCell"]).Value = array5[0] + "_" + staticOfTable.StaticCells[i].ParentCell.RowIndex;
					}
					break;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public string GetValue(SheetView sView, string index)
	{
		string result = string.Empty;
		try
		{
			if (index.Length > 0)
			{
				string[] array = index.Split(new char[1] { '_' });
				if (array.Length == 2)
				{
					int num = -1;
					int num2 = -1;
					num2 = int.Parse(array[1]) - 1;
					num = HTKKFile.ColumnLettersToIndex(array[0]) - 1;
					result = ((num < 0 || num2 < 0) ? null : sView.Cells[num2, num].Value.ToString());
				}
				else
				{
					result = null;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			result = null;
		}
		return result;
	}

	public void BindDataRowToRowNode(SheetView sView, XmlNode xmlRowDataNode, HTKKCMCTableInfo tableInfo, int iRowDistanct)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		XmlAttribute val2 = val.CreateAttribute("RowDisplay");
		((XmlNode)val2).Value = tableInfo.RowsValues[iRowDistanct].RowDisplay;
		((XmlNamedNodeMap)xmlRowDataNode.Attributes).SetNamedItem((XmlNode)(object)val2);
		xmlRowDataNode.InnerText = tableInfo.GetRowData(sView, iRowDistanct);
	}

	public void BindDataTableToTableNode(SheetView sView, XmlNode xmlTableNode, HTKKCMCTableInfo tableInfo, HTKKCMCTableInfo rootTableInfo)
	{
		int num = tableInfo.Location.DataLocationTemplateRowStart - rootTableInfo.Location.DataLocationTemplateRowStart;
		((XmlNode)xmlTableNode.Attributes["MaxRows"]).Value = tableInfo.RowsValues.Count.ToString();
		((XmlNode)xmlTableNode.Attributes["id"]).Value = tableInfo.TableID.ToString();
		if (tableInfo.CellTableSTT != null && xmlTableNode.Attributes["TableSTTValue"] != null)
		{
			((XmlNode)xmlTableNode.Attributes["TableSTTValue"]).Value = tableInfo.CellTableSTT.Value.ToString();
		}
		if (tableInfo.CellTableReportSTT != null && xmlTableNode.Attributes["TableReportSTTLocation"] != null)
		{
			((XmlNode)xmlTableNode.Attributes["TableReportSTTLocation"]).Value = HTKKFile.ColumnIndexToLetters(tableInfo.CellTableReportSTT.ColIndex) + "_" + tableInfo.CellTableReportSTT.RowIndex;
		}
		if (tableInfo.CellTableSTT != null && xmlTableNode.Attributes["TableSTTLocation"] != null)
		{
			((XmlNode)xmlTableNode.Attributes["TableSTTLocation"]).Value = HTKKFile.ColumnIndexToLetters(tableInfo.CellTableSTT.ColIndex) + "_" + tableInfo.CellTableSTT.RowIndex;
		}
		XmlNode val = xmlTableNode.SelectSingleNode("Location");
		XmlNode val2 = val.SelectSingleNode("TableTemplateRowStart");
		val2.InnerText = tableInfo.Location.TableTemplateRowStart.ToString();
		val2 = val.SelectSingleNode("TableReportRowStart");
		val2.InnerText = tableInfo.Location.TableReportRowStart.ToString();
		val2 = val.SelectSingleNode("DataLocationTemplateRowStart");
		val2.InnerText = tableInfo.Location.DataLocationTemplateRowStart.ToString();
		val2 = val.SelectSingleNode("DataLocationReportRowStart");
		val2.InnerText = tableInfo.Location.DataLocationReportRowStart.ToString();
		val2 = val.SelectSingleNode("DataLocationTaiBangKeRowStart");
		val2.InnerText = tableInfo.Location.DataLocationTaiBangKeRowStart.ToString();
		BindDataStaticTableToXMLNode(sView, xmlTableNode.SelectSingleNode("Header"), tableInfo.Header, num);
		XmlNode xmlRowDataNode = xmlTableNode.SelectSingleNode("RowData");
		if (xmlRowDataNode != null)
		{
			if (tableInfo.RowsValues.Count > 0)
			{
				BindDataRowToRowNode(sView, xmlRowDataNode, tableInfo, 0);
			}
			for (int i = 1; i < tableInfo.RowsValues.Count; i++)
			{
				InsertRowDataNode(sView, ref xmlRowDataNode, tableInfo, i);
			}
		}
		BindDataStaticTableToXMLNode(sView, xmlTableNode.SelectSingleNode("Footer"), tableInfo.Footer, num + tableInfo.RowsValues.Count - 1);
	}

	public void InsertTableNode(SheetView sView, ref XmlNode xmlTableNode, XmlNode xmlTableTemplateNode, HTKKCMCTableInfo tableInfo, HTKKCMCTableInfo rootTableInfo)
	{
		try
		{
			XmlNode val = xmlTableTemplateNode.Clone();
			BindDataTableToTableNode(sView, val, tableInfo, rootTableInfo);
			xmlTableNode.ParentNode.InsertAfter(val, xmlTableNode);
			xmlTableNode = val;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void InsertRowDataNode(SheetView sView, ref XmlNode xmlRowDataNode, HTKKCMCTableInfo tableInfo, int iRowDistanct)
	{
		try
		{
			XmlNode val = xmlRowDataNode.Clone();
			BindDataRowToRowNode(sView, val, tableInfo, iRowDistanct);
			xmlRowDataNode.ParentNode.InsertAfter(val, xmlRowDataNode);
			xmlRowDataNode = val;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}
}
