using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using FarPoint.Win.Spread.Model;
using HTKKDanhmuc;
using HTKKSerializer;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmRegistrationCatalogues : Form
{
	private int[] rowCheck = null;

	private int rowCount = 180;

	private int totalRow;

	private string strOldKeyWord = "";

	private int index;

	private bool irowCheck;

	private Row ar;

	private int LoaiDM = 0;

	private int startRowTimKiem = 3;

	private int activeSheet = 0;

	private bool isDataChanged = false;

	private string excelPath = "../InterfaceTemplates/excel/Catalogue_072016.xls";

	private string[] filePath = new string[2] { "", "" };

	private string[] filePathSave = new string[2] { "", "" };

	private bool isDmNew = false;

	private dmconfig dmconfig;

	private string filePathCf = "";

	private string filePathSaveCf = "";

	private List<danhmuc> danhmucs = new List<danhmuc>();

	private List<string> filePathData = new List<string>();

	private static XmlNode _xmlNode = null;

	private XmlNodeList elemList;

	public static Dictionary<object, object> catalog = new Dictionary<object, object>();

	private BackgroundWorker bw = new BackgroundWorker();

	private int countCbbHL = 0;

	private List<KeyValuePair<string, string>> dataCboTinh = null;

	private string maTinh = "";

	private string hieuLuc = "";

	private IContainer components = null;

	private Button btnGhi;

	private Button btnDong;

	private Panel pnWinSpread;

	private FpSpread winFpSpread;

	private SheetView winFpSpread_Sheet1;

	private ComboBox cbbHieuLuc;

	private Label label1;

	private Label label2;

	private TextBox txtKeyword;

	private Button btnSearch;

	private Label lblTinh;

	private ComboBox cbbTinh;

	private void btnSearch_Click(object sender, EventArgs e)
	{
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_0229: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Unknown result type (might be due to invalid IL or missing references)
		if (strOldKeyWord != ((Control)txtKeyword).Text)
		{
			startRowTimKiem = 3;
			index = 0;
		}
		strOldKeyWord = ((Control)txtKeyword).Text;
		if (HTKKFormater.isNullString(((Control)txtKeyword).Text))
		{
			frmMessage frmMessage = new frmMessage("0160", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		int num = 0;
		int activeSheetIndex = winFpSpread.ActiveSheetIndex;
		num = winFpSpread.ActiveSheet.NonEmptyRowCount - 3;
		totalRow = startRowTimKiem + num;
		int foundRowIndex = 0;
		int foundColumnIndex = 0;
		string text = "";
		winFpSpread.Sheets[activeSheetIndex].ClearSelection();
		int num2 = 5;
		if (WinSettings.MenuID == "100_14")
		{
			num2 = 6;
		}
		text = winFpSpread.Search(activeSheetIndex, ((Control)txtKeyword).Text.Trim(), caseSensitive: false, exactMatch: false, alternateSearch: false, useWildcards: false, index, num2, index + totalRow, num2, ref foundRowIndex, ref foundColumnIndex);
		if (foundRowIndex > -1)
		{
			ar = winFpSpread.Sheets[activeSheetIndex].Rows[foundRowIndex];
			index = foundRowIndex + 1;
			winFpSpread.Sheets[activeSheetIndex].AddSelection(foundRowIndex, 1, 1, 7);
			winFpSpread.SetViewportTopRow(0, foundRowIndex);
		}
		else if (index > 3)
		{
			frmMessage frmMessage2 = new frmMessage("0459", "5", "2", "Thông báo");
			((Form)frmMessage2).ShowDialog();
			if (frmMessage2.Yes)
			{
				winFpSpread.Sheets[activeSheetIndex].ClearSelection();
				index = 3;
				btnSearch_Click(sender, e);
			}
		}
		else
		{
			frmMessage frmMessage2 = new frmMessage("0160", "3", "", "Thông báo");
			((Form)frmMessage2).ShowDialog();
			index = 3;
		}
	}

	public frmRegistrationCatalogues(XmlNode xmlMenuNode)
	{
		InitializeComponent();
		_xmlNode = xmlMenuNode;
		if (WinSettings.MenuID == "100_9")
		{
			LoaiDM = 1;
		}
		else if (WinSettings.MenuID == "100_10")
		{
			LoaiDM = 2;
		}
		else if (WinSettings.MenuID == "100_13")
		{
			((Control)this).Text = "Đăng ký danh mục phí BVMT";
			LoaiDM = 1;
			((Control)cbbHieuLuc).Visible = true;
		}
		else if (WinSettings.MenuID == "100_14")
		{
			((Control)this).Text = "Bảng giá tính thuế của địa phương";
			LoaiDM = 1;
			((Control)btnGhi).Visible = false;
			((Control)lblTinh).Visible = true;
			((Control)cbbTinh).Visible = true;
			Point location = ((Control)label1).Location;
			Point location2 = ((Control)cbbHieuLuc).Location;
			((Control)label1).Location = new Point(((Control)label2).Location.X - 160, ((Control)label2).Location.Y);
			((Control)cbbHieuLuc).Location = new Point(((Control)txtKeyword).Location.X - 150, ((Control)txtKeyword).Location.Y);
			((Control)label2).Location = new Point(((Control)lblTinh).Location.X - 120, ((Control)lblTinh).Location.Y);
			((Control)txtKeyword).Location = new Point(((Control)cbbTinh).Location.X - 130, ((Control)cbbTinh).Location.Y);
			((Control)btnSearch).Location = new Point(((Control)btnSearch).Location.X + 250, ((Control)btnSearch).Location.Y);
			((Control)lblTinh).Location = location;
			((Control)cbbTinh).Location = location2;
			Dictionary<string, string> dictionary = new Dictionary<string, string>(HTKKWinDictionaryDiaDanh.dicTinhTP);
			dataCboTinh = dictionary.ToList();
			List<string> list = new List<string>();
			list.Add("Chọn tỉnh");
			list.AddRange(dictionary.Keys.ToList());
			cbbTinh.DataSource = list;
		}
		else
		{
			LoaiDM = 1;
			((Control)cbbHieuLuc).Visible = true;
		}
		SetDataFile();
		ReLoadForm();
	}

	private void ReLoadForm()
	{
		winFpSpread.SuspendLayout();
		activeSheet = winFpSpread.ActiveSheetIndex;
		WinFnProcess.ConvertCatalogue_BVMT_CF();
		for (int i = 0; i < winFpSpread.Sheets.Count; i++)
		{
			winFpSpread.Sheets[i].AutoCalculation = false;
		}
		winFpSpread.OpenExcel(excelPath);
		winFpSpread.ActiveSheetIndex = 0;
		for (int j = 0; j < filePathSave.Length; j++)
		{
			WinSpreadProcess.WinSetupSpreadSheet(winFpSpread.Sheets[j]);
			WinSpreadProcess.WinSetupSpread(winFpSpread);
			WinSpreadProcess.setSpread_WH(winFpSpread.Sheets[j], new HTKKCMCSheetInfo());
			winFpSpread.Sheets[winFpSpread.Sheets.Count - 1].Visible = false;
			winFpSpread.Sheets[winFpSpread.Sheets.Count - 2].Visible = false;
		}
		winFpSpread.TabStripInsertTab = false;
		CheckBoxType(winFpSpread);
		LoadData();
		if (winFpSpread.Sheets.Count > 3 && activeSheet > 0)
		{
			winFpSpread.ActiveSheetIndex = activeSheet;
		}
		else
		{
			activeSheet = 0;
		}
		winFpSpread.ResumeLayout();
		winFpSpread.ActiveSheet.ClearSelection();
		index = 3;
	}

	private void SetDataFile()
	{
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		//IL_0089: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected O, but got Unknown
		try
		{
			string text = "";
			text = ((!string.IsNullOrEmpty(maTinh)) ? maTinh : WinSettings.HTKKHeaderInfo.maCQTCapCuc.Substring(0, 3));
			XmlDocument val = new XmlDocument();
			if (WinSettings.MenuID == "100_14")
			{
				XmlDocument val2 = new XmlDocument();
				val2.Load("HieuLucBangGia.xml");
				XmlNodeList val3 = ((XmlNode)val2).SelectNodes("/Root/Map");
				foreach (XmlNode item in val3)
				{
					XmlNode val4 = item;
					if (((XmlNode)val4.Attributes["ID"]).Value == text)
					{
						val.LoadXml(val4.OuterXml);
						break;
					}
				}
			}
			else
			{
				val.LoadXml(_xmlNode.OuterXml);
			}
			elemList = val.GetElementsByTagName("Sheet");
			string value = "";
			catalog = new Dictionary<object, object>();
			isDmNew = false;
			filePathData.Clear();
			for (int i = 0; i < elemList.Count; i++)
			{
				string text2 = ((elemList[i].ParentNode.Attributes["StartDate"] != null) ? ((XmlNode)elemList[i].ParentNode.Attributes["StartDate"]).Value : "");
				if (WinSettings.MenuID == "100_13" && DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture) < DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture))
				{
					continue;
				}
				text2 = HTKKFormater.FormatStringToDateNumberBegin(text2).ToString();
				HTKKCatalogues hTKKCatalogues = getHTKKCatalogues(elemList[i].ParentNode);
				HTKKFile.AddOrUpdate(catalog, text2, hTKKCatalogues);
				if (!(DateTime.Now >= DateTime.ParseExact(((XmlNode)elemList[i].ParentNode.Attributes["StartDate"]).Value, "dd/MM/yyyy", CultureInfo.CurrentCulture)))
				{
					continue;
				}
				filePath[ConvertHelpers.ConvertToInt(((XmlNode)elemList[i].Attributes["ID"]).Value) - 1] = ((elemList[i].Attributes["InterfaceIni"] != null) ? ((XmlNode)elemList[i].Attributes["InterfaceIni"]).Value : "");
				filePathSave[ConvertHelpers.ConvertToInt(((XmlNode)elemList[i].Attributes["ID"]).Value) - 1] = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + ((elemList[i].Attributes["DataFile"] != null) ? ((XmlNode)elemList[i].Attributes["DataFile"]).Value : "") + ".xml";
				excelPath = ((elemList[i].ParentNode.Attributes["InterfaceTemplate"] != null) ? ((XmlNode)elemList[i].ParentNode.Attributes["InterfaceTemplate"]).Value : "");
				value = ((elemList[i].ParentNode.Attributes["StartDate"] != null) ? ((XmlNode)elemList[i].ParentNode.Attributes["StartDate"]).Value : "");
				value = HTKKFormater.FormatStringToDateNumberBegin(value).ToString();
				if (!(WinSettings.MenuID == "100_10") && DateTime.Now >= DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture) && DateTime.ParseExact(((XmlNode)elemList[i].ParentNode.Attributes["StartDate"]).Value, "dd/MM/yyyy", CultureInfo.CurrentCulture) >= DateTime.ParseExact("01/10/2019", "dd/MM/yyyy", CultureInfo.CurrentCulture))
				{
					filePathSaveCf = (("../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + elemList[i].ParentNode.Attributes["DataFile"] != null) ? ((XmlNode)elemList[i].ParentNode.Attributes["DataFile"]).Value : "");
					filePathCf = (("../InterfaceIni/" + WinSettings.HTKKHeaderInfo.mst + "/" + elemList[i].ParentNode.Attributes["DataFile"] != null) ? ((XmlNode)elemList[i].ParentNode.Attributes["DataFile"]).Value : "");
					string text3 = ((elemList[i].Attributes["InterfaceIni"] != null) ? ((XmlNode)elemList[i].Attributes["InterfaceIni"]).Value : "");
					if (WinSettings.MenuID == "100_14")
					{
						text3 = text3.Replace(".xml", "_" + text + "_" + ((XmlNode)elemList[i].ParentNode.Attributes["StartDate"]).Value.Replace("/", "") + ".xml");
						filePathCf = (("../InterfaceIni/" + WinSettings.HTKKHeaderInfo.mst + "/" + elemList[i].ParentNode.Attributes["DataFile"] != null) ? ("GIA_TINH/" + ((XmlNode)elemList[i].ParentNode.Attributes["DataFile"]).Value) : "");
					}
					filePathData.Add(text3);
					isDmNew = true;
				}
			}
			LoadDropHieuLuc();
			int num = 0;
			for (int j = 0; j < cbbHieuLuc.Items.Count; j++)
			{
				string text4 = ((KeyValuePair<string, string>)cbbHieuLuc.Items[j]).Key.ToString();
				if (text4.Equals(value))
				{
					num = j;
				}
			}
			if (((ListControl)cbbHieuLuc).SelectedIndex != num)
			{
				((ListControl)cbbHieuLuc).SelectedIndex = num;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LoadData()
	{
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Expected O, but got Unknown
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003b: Expected O, but got Unknown
		winFpSpread.ActiveSheet.AutoCalculation = false;
		winFpSpread.SuspendLayout();
		try
		{
			XmlDocument val = new XmlDocument();
			bool value = true;
			int row = 0;
			Font font = new Font("Tahoma", 8f, (FontStyle)1);
			TextCellType cellType = new TextCellType();
			if (isDmNew)
			{
				danhmucs.Clear();
				for (int i = 0; i < filePathData.Count(); i++)
				{
					Serializer serializer = new Serializer();
					string path = filePathSaveCf;
					if (!File.Exists(path))
					{
						path = filePathCf;
					}
					if (File.Exists(path))
					{
						string input = File.ReadAllText(path);
						dmconfig = serializer.Deserialize<dmconfig>(input);
					}
					string text = filePathData[i];
					danhmuc danhmuc = null;
					if (File.Exists(filePathData[i]))
					{
						string input = File.ReadAllText(filePathData[i]);
						danhmuc = serializer.Deserialize<danhmuc>(input);
					}
					if (danhmuc == null || danhmuc.items.Count() <= 0 || dmconfig == null || dmconfig.dmcfs.Count() <= 0)
					{
						continue;
					}
					danhmucs.Add(danhmuc);
					List<item> items = danhmuc.items;
					dmcf dmcf = null;
					foreach (dmcf dmcf2 in dmconfig.dmcfs)
					{
						if (!HTKKFormater.isNullString(dmcf2.type) && dmcf2.type.Equals(danhmuc.type))
						{
							dmcf = dmcf2;
							break;
						}
					}
					if (dmcf == null || dmcf.cells == null || dmcf.cells.Count() == 0)
					{
						continue;
					}
					int num = ConvertHelpers.ConvertToInt(dmcf.start);
					if (!HTKKFormater.isNullString(dmcf.checkbox))
					{
						winFpSpread.Sheets[i].Cells[num - 1, ColumnLettersToIndex(dmcf.checkbox) - 1].CellType = new CheckBoxCellType();
						winFpSpread.Sheets[i].Cells[num - 1, ColumnLettersToIndex(dmcf.checkbox) - 1].Locked = false;
					}
					WinSpreadProcess.InsertTableRows(winFpSpread.Sheets[i], num, -1, num, -1, danhmuc.items.Count() - 1, -1);
					if (HTKKFormater.isNullString(dmcf.active))
					{
						SetValue(winFpSpread.Sheets[i], 2, ColumnLettersToIndex(dmcf.checkbox), value: true);
					}
					for (int j = 0; j < items.Count(); j++)
					{
						item item = items[j];
						bool flag = false;
						bool flag2 = false;
						if (!HTKKFormater.isNullString(item.bold) && Convert.ToBoolean(item.bold))
						{
							flag2 = true;
						}
						if (!HTKKFormater.isNullString(item.check) && Convert.ToBoolean(item.check))
						{
							flag = true;
						}
						if (!HTKKFormater.isNullString(dmcf.checkbox))
						{
							if (!flag && WinSettings.MenuID != "100_7")
							{
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(dmcf.checkbox) - 1].CellType = cellType;
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(dmcf.checkbox) - 1].Locked = true;
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(dmcf.checkbox) - 1].Text = "";
							}
							else if (!HTKKFormater.isNullString(dmcf.active))
							{
								if (dmcf.active.IndexOf("~" + item.id + "~") > -1)
								{
									SetValue(winFpSpread.Sheets[i], num + j, ColumnLettersToIndex(dmcf.checkbox), value: true);
								}
							}
							else
							{
								SetValue(winFpSpread.Sheets[i], num + j, ColumnLettersToIndex(dmcf.checkbox), value);
							}
						}
						string text2 = "";
						foreach (cell cell in dmcf.cells)
						{
							string text3 = "";
							PropertyInfo[] properties = typeof(item).GetProperties();
							foreach (PropertyInfo propertyInfo in properties)
							{
								if (cell.field.Equals(propertyInfo.Name))
								{
									if (propertyInfo.GetValue(item, null) != null)
									{
										text3 = propertyInfo.GetValue(item, null).ToString();
									}
									break;
								}
							}
							if (!HTKKFormater.isNullString(text3) && text3.Length > text2.Length)
							{
								text2 = text3;
							}
							if (!HTKKFormater.isNullString(cell.ctltype) && cell.ctltype.Equals("16"))
							{
								NumberCellType numberCellType = new NumberCellType();
								numberCellType.ShowSeparator = true;
								numberCellType.NegativeFormat = NegativeFormat.Parentheses;
								numberCellType.DecimalPlaces = 0;
								numberCellType.DecimalSeparator = ",";
								numberCellType.Separator = ".";
								if (dmcf.type == "DANH_MUC_TTDB")
								{
									numberCellType.DecimalPlaces = 1;
								}
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(cell.col) - 1].CellType = numberCellType;
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(cell.col) - 1].Value = text3;
								if (dmcf.type == "DANH_MUC_TTDB")
								{
									winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(cell.col) - 1].Value = text3.Replace(".", ",");
								}
							}
							else
							{
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(cell.col) - 1].Value = text3;
							}
							if (flag2)
							{
								winFpSpread.Sheets[i].Cells[num + j - 1, ColumnLettersToIndex(cell.col) - 1].Font = font;
							}
						}
						if (text2.Length > 30)
						{
							Row row2 = winFpSpread.Sheets[i].Rows[num + j - 1];
							float preferredHeight = row2.GetPreferredHeight();
							float num2 = preferredHeight / row2.Height;
							preferredHeight += num2 * 5f;
							row2.Height = preferredHeight;
						}
					}
				}
			}
			else
			{
				for (int i = 0; i < winFpSpread.ActiveSheet.NonEmptyRowCount; i++)
				{
					if (winFpSpread.ActiveSheet.Cells[i, 1].CellType.ToString() == "CheckBoxCellType")
					{
						row = i;
						break;
					}
				}
				for (int i = 0; i < filePath.Length; i++)
				{
					if (string.IsNullOrEmpty(filePath[i]))
					{
						continue;
					}
					if (File.Exists(filePathSave[i]))
					{
						if (!(WinSettings.MenuID == "100_10"))
						{
							convertByTashk2427(filePathSave[i], filePath[i]);
						}
						val.Load(filePathSave[i]);
					}
					else
					{
						val.Load(filePath[i]);
					}
					winFpSpread.Sheets[i].Cells[row, ColumnLettersToIndex("B") - 1].Value = true;
					XmlNodeList elementsByTagName = val.GetElementsByTagName("Cell");
					for (int l = 0; l < elementsByTagName.Count; l++)
					{
						string value2 = ((XmlNode)elementsByTagName[l].Attributes["Value"]).Value;
						string text4 = ((XmlNode)elementsByTagName[l].Attributes["CellID"]).Value.ToString();
						switch (value2.Substring(value2.Length - 1, 1))
						{
						case "1":
							value = true;
							break;
						case "0":
							value = false;
							winFpSpread.Sheets[i].Cells[row, ColumnLettersToIndex("B") - 1].Value = false;
							break;
						case "2":
							continue;
						}
						SetValue(winFpSpread.Sheets[i], text4, value);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		winFpSpread.ResumeLayout();
	}

	private void convertByTashk2427(string pathData, string pathTmp)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Expected O, but got Unknown
		//IL_02fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0301: Expected O, but got Unknown
		try
		{
			XmlDocument val = new XmlDocument();
			val.Load(pathData);
			XmlNodeList elementsByTagName = val.GetElementsByTagName("Sections");
			if (elementsByTagName == null)
			{
				return;
			}
			XmlNode val2 = elementsByTagName[0];
			string value = ((XmlNode)val2.Attributes["Version"]).Value;
			if (!HTKKFormater.isNullString(value) && !(value.ToNumber() < 418m))
			{
				return;
			}
			((XmlNode)val2.Attributes["Version"]).Value = HTKKConfig.VERSION.Replace(".", "");
			XmlDocument val3 = new XmlDocument();
			val3.Load(pathTmp);
			XmlNodeList elementsByTagName2 = val3.GetElementsByTagName("Cell");
			for (int i = 0; i < elementsByTagName2.Count; i++)
			{
				XmlNode val4 = elementsByTagName2[i];
				string value2 = ((XmlNode)val4.Attributes["CellID"]).Value;
				if (value2.IndexOf("_") == -1)
				{
					continue;
				}
				string[] array = value2.Split(new char[1] { '_' });
				XmlNodeList val5 = null;
				if (array[1].ToNumber() >= 33m)
				{
					val5 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@CellID='" + array[0] + "_" + (array[1].ToNumber() - 2m) + "']");
				}
				else if (array[1].ToNumber() <= 29m)
				{
					val5 = ((XmlNode)val).SelectNodes("/Sections/Section/Cells/Cell[@CellID='" + value2 + "']");
				}
				else if (array[1].ToNumber() == 30m || array[1].ToNumber() == 31m)
				{
					((XmlNode)val4.Attributes["Value"]).Value = ((XmlNode)val4.Attributes["Value"]).Value.Replace("###0###1", "###0###0");
				}
				if (val5 == null)
				{
					continue;
				}
				XmlNode val6 = val5[0];
				if (val6 != null)
				{
					string value3 = ((XmlNode)val6.Attributes["Value"]).Value;
					if (!HTKKFormater.isNullString(value3))
					{
						((XmlNode)val4.Attributes["Value"]).Value = value3;
					}
				}
			}
			XmlTextWriter val7 = new XmlTextWriter(pathData, Encoding.UTF8);
			try
			{
				val7.Formatting = (Formatting)1;
				val3.Save((XmlWriter)(object)val7);
			}
			finally
			{
				((IDisposable)val7)?.Dispose();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static int ColumnLettersToIndex(string column)
	{
		int num = 0;
		int length = column.Length;
		for (int i = 0; i < length; i++)
		{
			num += (column[length - i - 1] - 65 + 1) * (int)Math.Pow(26.0, i);
		}
		return num;
	}

	public bool NotCheckType(int[] rowNotCheck, int irow)
	{
		for (int i = 0; i < rowNotCheck.Length; i++)
		{
			if (irow == rowNotCheck[i])
			{
				return true;
			}
		}
		return false;
	}

	public void CheckBoxType(FpSpread winFpSpread)
	{
		try
		{
			CheckBoxCellType cellType = new CheckBoxCellType();
			for (int i = 0; i < winFpSpread.Sheets.Count - 2; i++)
			{
				for (int j = 2; j < winFpSpread.Sheets[i].NonEmptyRowCount; j++)
				{
					if (winFpSpread.Sheets[i].Cells[j - 1, ColumnLettersToIndex("B") - 1].Text == "chk")
					{
						winFpSpread.Sheets[i].Cells[j - 1, ColumnLettersToIndex("B") - 1].CellType = cellType;
					}
					if (((XmlNode)_xmlNode.Attributes["ID"]).Value == "100_7" && j > 3)
					{
						winFpSpread.Sheets[i].Cells[j - 1, ColumnLettersToIndex("H") - 1].CellType = new NumberCellType
						{
							DecimalSeparator = ","
						};
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void Form1_Load(object sender, EventArgs e)
	{
	}

	private void checkedParent(int indexRow)
	{
		try
		{
			string text = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[indexRow, ColumnLettersToIndex("E") - 1].Text;
			if (HTKKFormater.isNullString(text) || danhmucs == null || danhmucs.Count() == 0 || dmconfig == null || dmconfig.dmcfs == null || dmconfig.dmcfs.Count() == 0)
			{
				return;
			}
			int num = 1;
			for (int i = num + 1; i < indexRow; i++)
			{
				string text2 = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("D") - 1].Text;
				if (!(text2 == text))
				{
					continue;
				}
				winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Value = false;
				for (int j = num + 1; j < winFpSpread.ActiveSheet.NonEmptyRowCount; j++)
				{
					string text3 = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[j, ColumnLettersToIndex("D") - 1].Text;
					if (!HTKKFormater.isNullString(text3) && text3.Length > 1 && text3.Length > text2.Length && text3.StartsWith(text2))
					{
						string text4 = text3.Substring(text2.Length);
						if (HTKKFormater.IsNumber(text4) && text4.ToNumber() > 0m && ConvertHelpers.ConvertToBool(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[j, ColumnLettersToIndex("B") - 1].Value))
						{
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Value = true;
							break;
						}
					}
				}
				checkedParent(i);
				break;
			}
			for (int k = 0; k < danhmucs.Count(); k++)
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_ButtonClicked(object sender, EditorNotifyEventArgs e)
	{
		try
		{
			isDataChanged = true;
			int num = 0;
			if (isDmNew)
			{
				num = 1;
				if (e.Column == ColumnLettersToIndex("B") - 1 && e.Row == 1)
				{
					bool flag = ConvertHelpers.ConvertToBool(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num, ColumnLettersToIndex("B") - 1].Value);
					for (int i = num + 1; i < winFpSpread.ActiveSheet.NonEmptyRowCount; i++)
					{
						if (WinSettings.MenuID == "100_7")
						{
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Value = flag;
						}
						else if (!winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Locked)
						{
							winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Value = flag;
						}
					}
					return;
				}
				if (WinSettings.MenuID == "100_7")
				{
					bool flag = ConvertHelpers.ConvertToBool(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[e.Row, ColumnLettersToIndex("B") - 1].Value);
					string text = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[e.Row, ColumnLettersToIndex("D") - 1].Text;
					if (!HTKKFormater.isNullString(text))
					{
						for (int i = num + 1; i < winFpSpread.ActiveSheet.NonEmptyRowCount; i++)
						{
							string text2 = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("D") - 1].Text;
							string text3 = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("E") - 1].Text;
							if (text2.IndexOf("II") > -1)
							{
								int num2 = 1;
							}
							if (!HTKKFormater.isNullString(text2) && text2.Length > 1 && text2.Length >= text.Length && !HTKKFormater.isNullString(text3) && !text.Equals("II1") && (text3.Equals(text) || text2.StartsWith(text)))
							{
								string text4 = text2.Substring(text.Length);
								if (!HTKKFormater.isNullString(text4) && HTKKFormater.IsNumber(text4) && (text4.ToNumber() > 0m || text3.Equals(text)))
								{
									winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Value = flag;
								}
							}
						}
					}
					checkedParent(e.Row);
					return;
				}
				bool flag2 = true;
				for (int i = 2; i < winFpSpread.ActiveSheet.NonEmptyRowCount; i++)
				{
					if (winFpSpread.ActiveSheet.Cells[i, 1].CellType != null && winFpSpread.ActiveSheet.Cells[i, 1].CellType.ToString() == "CheckBoxCellType" && !Convert.ToBoolean(winFpSpread.ActiveSheet.Cells[i, 1].Value))
					{
						flag2 = false;
						break;
					}
				}
				winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[1, ColumnLettersToIndex("B") - 1].Value = flag2;
				return;
			}
			for (int i = 0; i < winFpSpread.ActiveSheet.NonEmptyRowCount; i++)
			{
				if (winFpSpread.ActiveSheet.Cells[i, 1].CellType.ToString() == "CheckBoxCellType")
				{
					num = i;
					break;
				}
			}
			if (e.Row != num)
			{
				SheetView sheetView = winFpSpread.Sheets[winFpSpread.ActiveSheetIndex];
				if (e.Column != ColumnLettersToIndex("B") - 1)
				{
					return;
				}
				for (int i = num + 1; i < rowCount; i++)
				{
					if (winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].CellType.ToString() == "CheckBoxCellType" && !ConvertHelpers.ConvertToBool(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i, ColumnLettersToIndex("B") - 1].Value))
					{
						winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num, ColumnLettersToIndex("B") - 1].Value = false;
						return;
					}
				}
				winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num, ColumnLettersToIndex("B") - 1].Value = true;
				return;
			}
			irowCheck = ConvertHelpers.ConvertToBool(winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[num, ColumnLettersToIndex("B") - 1].Value);
			rowCount = winFpSpread.ActiveSheet.NonEmptyRowCount;
			if (e.Row != num || e.Column != ColumnLettersToIndex("B") - 1)
			{
				return;
			}
			for (int i = num + 1; i < rowCount; i++)
			{
				if (winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i - 1, ColumnLettersToIndex("B") - 1].CellType.ToStringOrDefault() == "CheckBoxCellType")
				{
					winFpSpread.Sheets[winFpSpread.ActiveSheetIndex].Cells[i - 1, ColumnLettersToIndex("B") - 1].Value = irowCheck;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static string GetValue(SheetView sView, string index)
	{
		string result = string.Empty;
		try
		{
			if (index.Length > 0)
			{
				int[] array = GeneralFunctions.ArrayCellId(index);
				result = sView.Cells[array[1], array[0] - 1].Text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static string GetValue(SheetView sView, int row, int coll)
	{
		string result = string.Empty;
		try
		{
			if (row > 0 && coll > 0)
			{
				result = sView.Cells[row - 1, coll - 1].Text;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return result;
	}

	public static void SetValue(SheetView sView, string index, bool value)
	{
		try
		{
			if (index.Length > 0)
			{
				int[] array = GeneralFunctions.ArrayCellId(index);
				sView.Cells[array[1], array[0] - 1].CellType = new CheckBoxCellType();
				sView.Cells[array[1], array[0] - 1].Value = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public static void SetValue(SheetView sView, int row, int coll, bool value)
	{
		try
		{
			if (row > 0 && coll > 0)
			{
				sView.Cells[row - 1, coll - 1].CellType = new CheckBoxCellType();
				sView.Cells[row - 1, coll - 1].Value = value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnGhi_Click(object sender, EventArgs e)
	{
		//IL_02d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_02db: Expected O, but got Unknown
		//IL_04bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_0464: Unknown result type (might be due to invalid IL or missing references)
		//IL_046b: Expected O, but got Unknown
		//IL_026f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0276: Expected O, but got Unknown
		//IL_028b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0292: Expected O, but got Unknown
		int num = 0;
		try
		{
			if (isDmNew)
			{
				if (danhmucs == null || danhmucs.Count() == 0 || dmconfig == null || dmconfig.dmcfs == null || dmconfig.dmcfs.Count() == 0)
				{
					return;
				}
				foreach (dmcf dmcf in dmconfig.dmcfs)
				{
					string text = "~";
					if (!HTKKFormater.isNullString(dmcf.type))
					{
						for (int i = 0; i < danhmucs.Count(); i++)
						{
							danhmuc danhmuc = danhmucs[i];
							if (danhmuc == null || HTKKFormater.isNullString(danhmuc.type) || danhmuc.items == null || danhmuc.items.Count() == 0 || HTKKFormater.isNullString(dmcf.checkbox) || !dmcf.type.Equals(danhmuc.type))
							{
								continue;
							}
							int num2 = ConvertHelpers.ConvertToInt(dmcf.start);
							for (int j = 0; j < danhmuc.items.Count(); j++)
							{
								string value = GetValue(winFpSpread.Sheets[i], num2 + j, ColumnLettersToIndex(dmcf.checkbox));
								if (!string.IsNullOrEmpty(value) && danhmuc.items[j] != null && !HTKKFormater.isNullString(danhmuc.items[j].id) && Convert.ToBoolean(value))
								{
									text = text + danhmuc.items[j].id + "~";
								}
							}
							break;
						}
					}
					if (!HTKKFormater.isNullString(text.Replace("~", "")))
					{
						dmcf.active = text;
					}
				}
				string text2 = dmconfig.ToXML();
				text2 = text2.Replace("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"", "");
				XmlDocument val = new XmlDocument();
				val.LoadXml(text2);
				XmlTextWriter val2 = new XmlTextWriter(filePathSaveCf, Encoding.UTF8);
				val2.Formatting = (Formatting)1;
				val.Save((XmlWriter)(object)val2);
			}
			else
			{
				int num3 = 0;
				for (int k = 0; k < filePathSave.Length; k++)
				{
					if (string.IsNullOrEmpty(filePath[k]))
					{
						continue;
					}
					XmlDocument val = new XmlDocument();
					val.Load(filePath[k]);
					XmlNodeList elementsByTagName = val.GetElementsByTagName("Sections");
					if (elementsByTagName != null)
					{
						XmlNode val3 = elementsByTagName[0];
						((XmlNode)val3.Attributes["Version"]).Value = HTKKConfig.VERSION.Replace(".", "");
					}
					elementsByTagName = val.GetElementsByTagName("Cell");
					for (int l = 0; l < elementsByTagName.Count; l++)
					{
						num = l;
						string text3 = ((XmlNode)elementsByTagName[l].Attributes["CellID"]).Value.ToString();
						string value = GetValue(winFpSpread.Sheets[k], text3);
						if (!string.IsNullOrEmpty(value))
						{
							num3 = ((!(value == "chk")) ? (Convert.ToBoolean(value) ? 1 : 0) : 0);
							string value2 = ((XmlNode)elementsByTagName[l].Attributes["Value"]).Value;
							((XmlNode)elementsByTagName[l].Attributes["Value"]).Value = value2.Substring(0, value2.Length - 1) + num3;
						}
					}
					XmlTextWriter val2 = new XmlTextWriter(filePathSave[k], Encoding.UTF8);
					val2.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val2);
				}
			}
			isDataChanged = false;
			frmMessage frmMessage = new frmMessage("0002", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnDong_Click(object sender, EventArgs e)
	{
		((Form)this).Close();
	}

	private HTKKCataloguesSheet getCatalogSheet(XmlNode xmlNodeSheet)
	{
		try
		{
			HTKKCataloguesSheet hTKKCataloguesSheet = new HTKKCataloguesSheet();
			hTKKCataloguesSheet.ID = ((xmlNodeSheet.Attributes["ID"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["ID"]).Value : "");
			hTKKCataloguesSheet.SheetName = ((xmlNodeSheet.Attributes["SheetName"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["SheetName"]).Value : "");
			hTKKCataloguesSheet.Caption = ((xmlNodeSheet.Attributes["Caption"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["Caption"]).Value : "");
			hTKKCataloguesSheet.Active = ((xmlNodeSheet.Attributes["Active"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["Active"]).Value : "");
			hTKKCataloguesSheet.TemplateFolder = ((xmlNodeSheet.Attributes["TemplateFolder"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["TemplateFolder"]).Value : "");
			hTKKCataloguesSheet.Folder = ((xmlNodeSheet.Attributes["Folder"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["Folder"]).Value : "");
			hTKKCataloguesSheet.DataFile = ((xmlNodeSheet.Attributes["DataFile"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["DataFile"]).Value : "");
			hTKKCataloguesSheet.InterfaceIni = ((xmlNodeSheet.Attributes["InterfaceIni"] != null) ? ((XmlNode)xmlNodeSheet.Attributes["InterfaceIni"]).Value : "");
			return hTKKCataloguesSheet;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private HTKKCatalogues getHTKKCatalogues(XmlNode xmlMenuValidity)
	{
		//IL_00f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fc: Expected O, but got Unknown
		if (xmlMenuValidity == null)
		{
			return null;
		}
		try
		{
			HTKKCatalogues hTKKCatalogues = new HTKKCatalogues();
			string input = ((xmlMenuValidity.Attributes["StartDate"] != null) ? ((XmlNode)xmlMenuValidity.Attributes["StartDate"]).Value : "");
			hTKKCatalogues.StartDate = HTKKFormater.ConvertToDateTime(input, out var _);
			hTKKCatalogues.InterfaceTemplate = ((xmlMenuValidity.Attributes["InterfaceTemplate"] != null) ? ((XmlNode)xmlMenuValidity.Attributes["InterfaceTemplate"]).Value : "");
			hTKKCatalogues.DataFile = ((xmlMenuValidity.Attributes["DataFile"] != null) ? ((XmlNode)xmlMenuValidity.Attributes["DataFile"]).Value : "");
			List<HTKKCataloguesSheet> list = new List<HTKKCataloguesSheet>();
			if (xmlMenuValidity.HasChildNodes)
			{
				foreach (XmlNode childNode in xmlMenuValidity.ChildNodes)
				{
					XmlNode xmlNodeSheet = childNode;
					HTKKCataloguesSheet catalogSheet = getCatalogSheet(xmlNodeSheet);
					list.Add(catalogSheet);
				}
			}
			hTKKCatalogues.catalogSheets = list;
			return hTKKCatalogues;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private void LoadDropHieuLuc()
	{
		//IL_0114: Unknown result type (might be due to invalid IL or missing references)
		//IL_011e: Expected O, but got Unknown
		try
		{
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			Dictionary<object, object> dictionary2 = catalog;
			if (dictionary2.Count <= 0)
			{
				return;
			}
			for (int num = dictionary2.Count - 1; num >= 0; num--)
			{
				KeyValuePair<object, object> keyValuePair = dictionary2.ElementAt(num);
				if (dictionary2.Count == 1)
				{
					dictionary.Add(ConvertHelpers.ConvertToString(keyValuePair.Key), " ");
				}
				else if (ConvertHelpers.ConvertToInt(((HTKKCatalogues)keyValuePair.Value).StartDate.ToString("yyyy")) < 2007)
				{
					dictionary.Add(ConvertHelpers.ConvertToString(keyValuePair.Key), "Danh mục có hiệu lực trước năm 2007");
				}
				else
				{
					dictionary.Add(ConvertHelpers.ConvertToString(keyValuePair.Key), "Danh mục có hiệu lực từ " + ((HTKKCatalogues)keyValuePair.Value).StartDate.ToString("dd/MM/yyyy"));
				}
			}
			cbbHieuLuc.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)cbbHieuLuc).DisplayMember = "Value";
			((ListControl)cbbHieuLuc).ValueMember = "Key";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cbbHieuLuc_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string text = (hieuLuc = ((KeyValuePair<string, string>)cbbHieuLuc.SelectedItem).Key.ToString());
			if (ConvertHelpers.ConvertToInt(text) >= ConvertHelpers.ConvertToInt(ConvertHelpers.ConvertDateTimeToStr(ConvertHelpers.ConvertToDateTime("01/10/2019", "dd/MM/yyyy"), "yyyyMMdd")))
			{
				filePathData.Clear();
				if (WinSettings.MenuID == "100_10")
				{
					isDmNew = false;
				}
				else
				{
					isDmNew = true;
				}
			}
			else
			{
				isDmNew = false;
			}
			catalog.TryGetValue(text, out var value);
			HTKKCatalogues hTKKCatalogues = (HTKKCatalogues)value;
			if (hTKKCatalogues == null)
			{
				return;
			}
			filePathSaveCf = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + hTKKCatalogues.DataFile;
			if (WinSettings.MenuID == "100_14")
			{
				filePathCf = "../InterfaceIni/GIA_TINH/" + hTKKCatalogues.DataFile;
			}
			else
			{
				filePathCf = "../InterfaceIni/" + hTKKCatalogues.DataFile;
			}
			for (int i = 0; i < hTKKCatalogues.catalogSheets.Count; i++)
			{
				filePath[i] = hTKKCatalogues.catalogSheets[i].InterfaceIni;
				filePathSave[i] = "../DataFiles/" + WinSettings.HTKKHeaderInfo.mst + "/" + hTKKCatalogues.catalogSheets[i].DataFile + ".xml";
				if (isDmNew)
				{
					string text2 = filePath[i];
					if (WinSettings.MenuID == "100_14")
					{
						text2 = ((!string.IsNullOrEmpty(maTinh)) ? text2.Replace(".xml", "_" + maTinh + "_" + ConvertHelpers.ConvertDateTimeToStr(hTKKCatalogues.StartDate, "ddMMyyyy") + ".xml") : text2.Replace(".xml", "_" + WinSettings.HTKKHeaderInfo.maCQTCapCuc.Substring(0, 3) + "_" + ConvertHelpers.ConvertDateTimeToStr(hTKKCatalogues.StartDate, "ddMMyyyy") + ".xml"));
					}
					filePathData.Add(text2);
				}
			}
			excelPath = hTKKCatalogues.InterfaceTemplate;
			if (countCbbHL > 0)
			{
				if (File.Exists("ProgressIndicatorTest.exe"))
				{
					Process.Start("ProgressIndicatorTest.exe");
				}
				ReLoadForm();
			}
			Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
			foreach (Process process in processesByName)
			{
				process.Kill();
			}
			((ContainerControl)this).ActiveControl = (Control)(object)winFpSpread;
			((Control)winFpSpread).Focus();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmRegistrationCatalogues_Resize(object sender, EventArgs e)
	{
		FormProcess.SetConfigForProgcessBar((Control)(object)this);
	}

	private void cbbHieuLuc_Click(object sender, EventArgs e)
	{
		countCbbHL++;
	}

	private void frmRegistrationCatalogues_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
			WinSettings.ToKhaiInfo = new HTKKCMCToKhaiInfo();
			WinSettings.TTPeriodInfo = new HTKKPeriodInfo();
			FormProcess.ReLoadfrmSystems();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_SheetTabClick(object sender, SheetTabClickEventArgs e)
	{
		index = 3;
	}

	private void frmRegistrationCatalogues_FormClosing(object sender, FormClosingEventArgs e)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		if (isDataChanged)
		{
			frmMessage frmMessage = new frmMessage("0013", "6", "2", "Cảnh báo");
			((Form)frmMessage).ShowDialog();
			if (frmMessage.Yes)
			{
				btnGhi_Click(sender, (EventArgs)(object)e);
				isDataChanged = false;
			}
			else if (frmMessage.Cancel)
			{
				((CancelEventArgs)(object)e).Cancel = true;
				((Control)frmMessage).Hide();
			}
			else
			{
				isDataChanged = false;
			}
		}
	}

	private void txtKeyword_KeyDown(object sender, KeyEventArgs e)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Invalid comparison between Unknown and I4
		try
		{
			if ((int)e.KeyCode == 13)
			{
				btnSearch_Click(sender, (EventArgs)(object)e);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			throw ex;
		}
	}

	private void cbbTinh_SelectedIndexChanged(object sender, EventArgs e)
	{
		try
		{
			string tenTinh = ((Control)cbbTinh).Text.Trim();
			if (tenTinh == "Chọn tỉnh")
			{
				maTinh = "";
			}
			else
			{
				maTinh = HTKKWinDictionaryDiaDanh.dicTinhTP.FirstOrDefault((KeyValuePair<string, string> x) => x.Key == tenTinh).Value;
			}
			if (((ListControl)cbbTinh).SelectedIndex > 0)
			{
				countCbbHL++;
				SetDataFile();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		((Form)this).Dispose(disposing);
	}

	private void InitializeComponent()
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Expected O, but got Unknown
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Expected O, but got Unknown
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_005e: Expected O, but got Unknown
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0069: Expected O, but got Unknown
		//IL_006a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_008a: Expected O, but got Unknown
		//IL_008b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_0aaa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab4: Expected O, but got Unknown
		//IL_0d2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d37: Expected O, but got Unknown
		//IL_0d68: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d72: Expected O, but got Unknown
		//IL_0d7b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d85: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmRegistrationCatalogues));
		btnGhi = new Button();
		btnDong = new Button();
		pnWinSpread = new Panel();
		winFpSpread = new FpSpread();
		winFpSpread_Sheet1 = new SheetView();
		cbbHieuLuc = new ComboBox();
		label1 = new Label();
		label2 = new Label();
		txtKeyword = new TextBox();
		btnSearch = new Button();
		lblTinh = new Label();
		cbbTinh = new ComboBox();
		((Control)pnWinSpread).SuspendLayout();
		((ISupportInitialize)winFpSpread).BeginInit();
		((ISupportInitialize)winFpSpread_Sheet1).BeginInit();
		((Control)this).SuspendLayout();
		((Control)btnGhi).Anchor = (AnchorStyles)10;
		((Control)btnGhi).Location = new Point(976, 573);
		((Control)btnGhi).Name = "btnGhi";
		((Control)btnGhi).Size = new Size(75, 23);
		((Control)btnGhi).TabIndex = 1;
		((Control)btnGhi).Text = "&Ghi";
		((ButtonBase)btnGhi).UseVisualStyleBackColor = true;
		((Control)btnGhi).Click += btnGhi_Click;
		((Control)btnDong).Anchor = (AnchorStyles)10;
		((Control)btnDong).Location = new Point(1067, 573);
		((Control)btnDong).Name = "btnDong";
		((Control)btnDong).Size = new Size(75, 23);
		((Control)btnDong).TabIndex = 1;
		((Control)btnDong).Text = "Đó&ng";
		((ButtonBase)btnDong).UseVisualStyleBackColor = true;
		((Control)btnDong).Click += btnDong_Click;
		((Control)pnWinSpread).Anchor = (AnchorStyles)15;
		((Control)pnWinSpread).BackgroundImageLayout = (ImageLayout)3;
		pnWinSpread.BorderStyle = (BorderStyle)2;
		((Control)pnWinSpread).Controls.Add((Control)(object)winFpSpread);
		((Control)pnWinSpread).Location = new Point(2, 41);
		((Control)pnWinSpread).Name = "pnWinSpread";
		((Control)pnWinSpread).Size = new Size(1155, 521);
		((Control)pnWinSpread).TabIndex = 27;
		((Control)winFpSpread).AccessibleDescription = "winFpSpread, Sheet1, Row 0, Column 0, ";
		((Control)winFpSpread).Anchor = (AnchorStyles)15;
		((Control)winFpSpread).BackColor = SystemColors.Control;
		((Control)winFpSpread).BackgroundImageLayout = (ImageLayout)0;
		((Control)winFpSpread).Location = new Point(0, -2);
		winFpSpread.Name = "winFpSpread";
		winFpSpread.Sheets.AddRange(new SheetView[1] { winFpSpread_Sheet1 });
		((Control)winFpSpread).Size = new Size(1140, 521);
		((Control)winFpSpread).TabIndex = 25;
		winFpSpread.TextTipPolicy = TextTipPolicy.Floating;
		winFpSpread.SheetTabClick += winFpSpread_SheetTabClick;
		winFpSpread.ButtonClicked += winFpSpread_ButtonClicked;
		winFpSpread_Sheet1.Reset();
		winFpSpread_Sheet1.SheetName = "Sheet1";
		winFpSpread_Sheet1.ReferenceStyle = ReferenceStyle.R1C1;
		winFpSpread_Sheet1.Cells.Get(0, 0).BackColor = Color.White;
		winFpSpread_Sheet1.Cells.Get(0, 17).BackColor = Color.White;
		winFpSpread_Sheet1.Cells.Get(10, 17).BackColor = Color.White;
		winFpSpread_Sheet1.ColumnFooter.DefaultStyle.BackColor = SystemColors.Control;
		winFpSpread_Sheet1.ColumnFooter.DefaultStyle.ForeColor = SystemColors.Control;
		winFpSpread_Sheet1.ColumnFooter.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		winFpSpread_Sheet1.ColumnFooter.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.BackColor = SystemColors.Control;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerFooterDefaultEnhanced";
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.BackColor = Color.Empty;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.DefaultStyle.BackColor = Color.White;
		winFpSpread_Sheet1.DefaultStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.DefaultStyle.Parent = "DataAreaDefault";
		winFpSpread_Sheet1.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.BackColor = Color.Empty;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefaultEnhanced";
		winFpSpread_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.FilterBarHeaderStyle.BackColor = Color.Empty;
		winFpSpread_Sheet1.FilterBarHeaderStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.FilterBarHeaderStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.FilterBarHeaderStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.FilterBarHeaderStyle.Parent = "RowHeaderDefaultEnhanced";
		winFpSpread_Sheet1.FilterBarHeaderStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.HorizontalGridLine = new GridLine(GridLineType.Flat, Color.LightGray, SystemColors.Control, SystemColors.ControlDark);
		winFpSpread_Sheet1.LockBackColor = SystemColors.Control;
		winFpSpread_Sheet1.LockForeColor = SystemColors.Control;
		winFpSpread_Sheet1.NullBackColor = SystemColors.Control;
		winFpSpread_Sheet1.NullForeColor = SystemColors.Control;
		winFpSpread_Sheet1.RowHeader.Columns.Default.Resizable = false;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.BackColor = Color.Empty;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefaultEnhanced";
		winFpSpread_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.SelectionBackColor = SystemColors.Control;
		winFpSpread_Sheet1.SelectionForeColor = SystemColors.Control;
		winFpSpread_Sheet1.SheetCornerStyle.BackColor = Color.Empty;
		winFpSpread_Sheet1.SheetCornerStyle.ForeColor = Color.Empty;
		winFpSpread_Sheet1.SheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.SheetCornerStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.SheetCornerStyle.Parent = "CornerDefaultEnhanced";
		winFpSpread_Sheet1.SheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.ReferenceStyle = ReferenceStyle.A1;
		cbbHieuLuc.DropDownStyle = (ComboBoxStyle)2;
		((ListControl)cbbHieuLuc).FormattingEnabled = true;
		((Control)cbbHieuLuc).Location = new Point(110, 12);
		((Control)cbbHieuLuc).Name = "cbbHieuLuc";
		((Control)cbbHieuLuc).Size = new Size(301, 21);
		((Control)cbbHieuLuc).TabIndex = 28;
		cbbHieuLuc.SelectedIndexChanged += cbbHieuLuc_SelectedIndexChanged;
		((Control)cbbHieuLuc).Click += cbbHieuLuc_Click;
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(59, 16);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(49, 13);
		((Control)label1).TabIndex = 29;
		((Control)label1).Text = "Hiệu lực:";
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(443, 16);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(29, 13);
		((Control)label2).TabIndex = 30;
		((Control)label2).Text = "Tên:";
		((Control)txtKeyword).Location = new Point(490, 13);
		((Control)txtKeyword).Name = "txtKeyword";
		((Control)txtKeyword).Size = new Size(186, 20);
		((Control)txtKeyword).TabIndex = 31;
		((Control)txtKeyword).KeyDown += new KeyEventHandler(txtKeyword_KeyDown);
		((Control)btnSearch).Location = new Point(685, 12);
		((Control)btnSearch).Name = "btnSearch";
		((Control)btnSearch).Size = new Size(75, 23);
		((Control)btnSearch).TabIndex = 32;
		((Control)btnSearch).Text = "Tìm kiếm";
		((ButtonBase)btnSearch).UseVisualStyleBackColor = true;
		((Control)btnSearch).Click += btnSearch_Click;
		((Control)lblTinh).AutoSize = true;
		((Control)lblTinh).Location = new Point(806, 19);
		((Control)lblTinh).Name = "lblTinh";
		((Control)lblTinh).Size = new Size(31, 13);
		((Control)lblTinh).TabIndex = 33;
		((Control)lblTinh).Text = "Tỉnh:";
		((Control)lblTinh).Visible = false;
		cbbTinh.DropDownStyle = (ComboBoxStyle)2;
		((ListControl)cbbTinh).FormattingEnabled = true;
		((Control)cbbTinh).Location = new Point(854, 14);
		((Control)cbbTinh).Name = "cbbTinh";
		((Control)cbbTinh).Size = new Size(122, 21);
		((Control)cbbTinh).TabIndex = 34;
		((Control)cbbTinh).Visible = false;
		cbbTinh.SelectedIndexChanged += cbbTinh_SelectedIndexChanged;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(1158, 603);
		((Control)this).Controls.Add((Control)(object)cbbTinh);
		((Control)this).Controls.Add((Control)(object)lblTinh);
		((Control)this).Controls.Add((Control)(object)btnSearch);
		((Control)this).Controls.Add((Control)(object)txtKeyword);
		((Control)this).Controls.Add((Control)(object)label2);
		((Control)this).Controls.Add((Control)(object)label1);
		((Control)this).Controls.Add((Control)(object)cbbHieuLuc);
		((Control)this).Controls.Add((Control)(object)pnWinSpread);
		((Control)this).Controls.Add((Control)(object)btnDong);
		((Control)this).Controls.Add((Control)(object)btnGhi);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Control)this).Name = "frmRegistrationCatalogues";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Đăng ký biểu thuế";
		((Form)this).WindowState = (FormWindowState)2;
		((Form)this).FormClosing += new FormClosingEventHandler(frmRegistrationCatalogues_FormClosing);
		((Form)this).FormClosed += new FormClosedEventHandler(frmRegistrationCatalogues_FormClosed);
		((Form)this).Load += Form1_Load;
		((Control)this).Resize += frmRegistrationCatalogues_Resize;
		((Control)pnWinSpread).ResumeLayout(false);
		((ISupportInitialize)winFpSpread).EndInit();
		((ISupportInitialize)winFpSpread_Sheet1).EndInit();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
