using System;
using System.Collections.Generic;
using System.Reflection;
using FarPoint.Win.Spread;
using TAX.HTKK.Utilities.Core;

namespace TAX.HTKK.Utilities.Entity;

public class HTKKCMCTableInfo
{
	public string TableID { get; set; }

	public string TableName { get; set; }

	public string TablePath { get; set; }

	public int iColumnSTT { get; set; }

	public int iColumnSTTReport { get; set; }

	public HTKKCMCCellBasicInfo CellTableSTT { get; set; }

	public HTKKCMCCellBasicInfo CellTableReportSTT { get; set; }

	public HTKKCMCStaticOfTable Header { get; set; }

	public HTKKCMCDataLocation Location { get; set; }

	public HTKKCMCRowInfo RowsInfo { get; set; }

	public List<HTKKCMCRowValue> RowsValues { get; set; }

	public HTKKCMCStaticOfTable Footer { get; set; }

	public int TableLength { get; set; }

	public HTKKCMCTableInfo()
	{
		TableID = "0";
		TableName = string.Empty;
		TablePath = string.Empty;
		iColumnSTT = -1;
		iColumnSTTReport = -1;
		CellTableSTT = new HTKKCMCCellBasicInfo();
		CellTableReportSTT = new HTKKCMCCellBasicInfo();
		Location = new HTKKCMCDataLocation();
		RowsInfo = new HTKKCMCRowInfo();
		RowsValues = new List<HTKKCMCRowValue>();
		Header = new HTKKCMCStaticOfTable();
		Footer = new HTKKCMCStaticOfTable();
		TableLength = 0;
	}

	public HTKKCMCTableInfo(HTKKCMCTableInfo table, int check)
	{
		int num = table.TableLength + table.RowsValues.Count;
		if (check != 1)
		{
			TableID = table.TableID;
		}
		TableName = table.TableName;
		TablePath = table.TablePath;
		iColumnSTT = table.iColumnSTT;
		iColumnSTTReport = table.iColumnSTTReport;
		if (table.CellTableSTT != null)
		{
			CellTableSTT = new HTKKCMCCellBasicInfo(table.CellTableSTT);
		}
		else
		{
			CellTableSTT = null;
		}
		if (table.CellTableReportSTT != null)
		{
			CellTableReportSTT = new HTKKCMCCellBasicInfo(table.CellTableReportSTT);
		}
		else
		{
			CellTableReportSTT = null;
		}
		Location = new HTKKCMCDataLocation(table.Location);
		RowsInfo = new HTKKCMCRowInfo(table.RowsInfo);
		RowsValues = new List<HTKKCMCRowValue>();
		if (check == 1)
		{
			if (table.RowsInfo.Cells.Count != 0)
			{
				RowsValues.Add(table.getDefaultRowValue());
			}
		}
		else if (table.RowsValues.Count > 0)
		{
			RowsValues = table.RowsValues;
		}
		Header = new HTKKCMCStaticOfTable(table.Header);
		Footer = new HTKKCMCStaticOfTable(table.Footer);
		TableLength = table.TableLength;
	}

	public HTKKCMCTableInfo Clone()
	{
		HTKKCMCTableInfo reTable = new HTKKCMCTableInfo();
		reTable.TableID = TableID;
		reTable.TableName = TableName;
		reTable.TablePath = TablePath;
		reTable.iColumnSTT = iColumnSTT;
		reTable.iColumnSTTReport = iColumnSTTReport;
		if (CellTableSTT != null)
		{
			reTable.CellTableSTT = new HTKKCMCCellBasicInfo(CellTableSTT);
		}
		else
		{
			reTable.CellTableSTT = null;
		}
		if (CellTableReportSTT != null)
		{
			reTable.CellTableReportSTT = new HTKKCMCCellBasicInfo(CellTableReportSTT);
		}
		else
		{
			reTable.CellTableReportSTT = null;
		}
		reTable.Location = new HTKKCMCDataLocation(Location);
		reTable.RowsInfo = new HTKKCMCRowInfo(RowsInfo);
		reTable.RowsValues = new List<HTKKCMCRowValue>();
		RowsValues.ForEach(delegate(HTKKCMCRowValue item)
		{
			reTable.RowsValues.Add(new HTKKCMCRowValue(item));
		});
		reTable.Header = new HTKKCMCStaticOfTable(Header);
		reTable.Footer = new HTKKCMCStaticOfTable(Footer);
		reTable.TableLength = TableLength;
		return reTable;
	}

	public HTKKCMCTableInfo CloneWithoutData()
	{
		HTKKCMCTableInfo hTKKCMCTableInfo = new HTKKCMCTableInfo();
		hTKKCMCTableInfo.TableID = TableID;
		hTKKCMCTableInfo.TableName = TableName;
		hTKKCMCTableInfo.TablePath = TablePath;
		hTKKCMCTableInfo.iColumnSTT = iColumnSTT;
		hTKKCMCTableInfo.iColumnSTTReport = iColumnSTTReport;
		if (CellTableSTT != null)
		{
			hTKKCMCTableInfo.CellTableSTT = new HTKKCMCCellBasicInfo(CellTableSTT);
		}
		else
		{
			hTKKCMCTableInfo.CellTableSTT = null;
		}
		if (CellTableReportSTT != null)
		{
			hTKKCMCTableInfo.CellTableReportSTT = new HTKKCMCCellBasicInfo(CellTableReportSTT);
		}
		else
		{
			hTKKCMCTableInfo.CellTableReportSTT = null;
		}
		hTKKCMCTableInfo.Location = new HTKKCMCDataLocation(Location);
		hTKKCMCTableInfo.RowsInfo = new HTKKCMCRowInfo(RowsInfo);
		hTKKCMCTableInfo.RowsValues = new List<HTKKCMCRowValue>();
		hTKKCMCTableInfo.Header = new HTKKCMCStaticOfTable(Header);
		hTKKCMCTableInfo.Footer = new HTKKCMCStaticOfTable(Footer);
		hTKKCMCTableInfo.TableLength = TableLength;
		return hTKKCMCTableInfo;
	}

	public void Increment(int distanct)
	{
		if (CellTableSTT != null)
		{
			CellTableSTT.RowIndex += distanct;
		}
		if (CellTableReportSTT != null)
		{
			CellTableReportSTT.RowIndex += distanct;
		}
		Location.DataLocationTemplateRowStart += distanct;
		Location.DataLocationReportRowStart += distanct;
		Location.TableTemplateRowStart += distanct;
		Location.TableReportRowStart += distanct;
		foreach (HTKKCMCCellInfo staticCell in Header.StaticCells)
		{
			if (staticCell.CellReport != null)
			{
				staticCell.CellReport.RowIndex += distanct;
			}
			if (staticCell.CellSelectedValue != null)
			{
				staticCell.CellSelectedValue.RowIndex += distanct;
			}
			if (staticCell.CellTemplate != null)
			{
				staticCell.CellTemplate.RowIndex += distanct;
			}
			if (staticCell.ChildCell != null)
			{
				staticCell.ChildCell.RowIndex += distanct;
			}
			if (staticCell.ParentCell != null)
			{
				staticCell.ParentCell.RowIndex += distanct;
			}
		}
	}

	public void IncrementFooter(int distanct)
	{
		foreach (HTKKCMCCellInfo staticCell in Footer.StaticCells)
		{
			if (staticCell.CellReport != null)
			{
				staticCell.CellReport.RowIndex += distanct;
			}
			if (staticCell.CellSelectedValue != null)
			{
				staticCell.CellSelectedValue.RowIndex += distanct;
			}
			if (staticCell.CellTemplate != null)
			{
				staticCell.CellTemplate.RowIndex += distanct;
			}
			if (staticCell.ChildCell != null)
			{
				staticCell.ChildCell.RowIndex += distanct;
			}
			if (staticCell.ParentCell != null)
			{
				staticCell.ParentCell.RowIndex += distanct;
			}
		}
	}

	public void IncrementFooter(long incrementPosition, int distanct)
	{
		foreach (HTKKCMCCellInfo staticCell in Footer.StaticCells)
		{
			if (staticCell.CellTemplate.RowIndex > incrementPosition)
			{
				if (staticCell.CellReport != null)
				{
					staticCell.CellReport.RowIndex += distanct;
				}
				if (staticCell.CellSelectedValue != null)
				{
					staticCell.CellSelectedValue.RowIndex += distanct;
				}
				if (staticCell.CellTemplate != null)
				{
					staticCell.CellTemplate.RowIndex += distanct;
				}
				if (staticCell.ChildCell != null)
				{
					staticCell.ChildCell.RowIndex += distanct;
				}
				if (staticCell.ParentCell != null)
				{
					staticCell.ParentCell.RowIndex += distanct;
				}
			}
		}
	}

	public bool addNewDefaultRow(int iRow)
	{
		try
		{
			HTKKCMCRowValue defaultRowValue = getDefaultRowValue();
			defaultRowValue.RowDisplay = "1";
			RowsValues.Insert(iRow, defaultRowValue);
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool IsEqualDefaultValue(HTKKCMCRowValue intemRowValue)
	{
		bool result = true;
		try
		{
			HTKKCMCRowValue defaultRowValue = getDefaultRowValue();
			for (int i = 0; i < RowsInfo.Cells.Count; i++)
			{
				if (intemRowValue.RowValue[i] != defaultRowValue.RowValue[i])
				{
					result = false;
					break;
				}
			}
			return result;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public HTKKCMCRowValue getDefaultRowValue()
	{
		try
		{
			HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
			for (int i = 0; i < RowsInfo.Cells.Count; i++)
			{
				hTKKCMCRowValue.RowValue.Add(RowsInfo.Cells[i].DefaultValue);
			}
			hTKKCMCRowValue.RowDisplay = RowsInfo.RowDefaultDisplay;
			return hTKKCMCRowValue;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public string GetRowDataDefault()
	{
		string text = string.Empty;
		try
		{
			for (int i = 0; i < RowsInfo.Cells.Count; i++)
			{
				text = text + RowsInfo.Cells[i].DefaultValue + "###";
			}
			if (RowsInfo.Cells.Count > 1)
			{
				text = text.Substring(0, text.Length - 3);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			text = null;
		}
		return text;
	}

	public string GetRowData(SheetView sView, int iRowDistanct)
	{
		string text = string.Empty;
		try
		{
			for (int i = 0; i < RowsInfo.Cells.Count; i++)
			{
				text = ((!(sView.Cells[Location.DataLocationTemplateRowStart + iRowDistanct - 1, RowsInfo.Cells[i].CellTemplate.ColIndex - 1].CellType.ToString() == "DateTimeCellType")) ? ((!(RowsInfo.Cells[i].Controltype == "104")) ? ((sView.Cells[Location.DataLocationTemplateRowStart + iRowDistanct - 1, RowsInfo.Cells[i].CellTemplate.ColIndex - 1].Value == null) ? (text + "###") : string.Concat(text, sView.Cells[Location.DataLocationTemplateRowStart + iRowDistanct - 1, RowsInfo.Cells[i].CellTemplate.ColIndex - 1].Value, "###")) : (text + sView.Cells[Location.DataLocationTemplateRowStart + iRowDistanct - 1, RowsInfo.Cells[i].CellTemplate.ColIndex - 1].Text + "###")) : (text + sView.Cells[Location.DataLocationTemplateRowStart + iRowDistanct - 1, RowsInfo.Cells[i].CellTemplate.ColIndex - 1].Text + "###"));
			}
			if (RowsInfo.Cells.Count > 1)
			{
				text = text.Substring(0, text.Length - 3);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			text = null;
		}
		return text;
	}
}
