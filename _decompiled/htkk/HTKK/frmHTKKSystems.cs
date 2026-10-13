using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Xml;
using FarPoint.Win;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using FarPoint.Win.Spread.Model;
using HTKK.Forms;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;

namespace HTKK;

public class frmHTKKSystems : Form
{
	private static string fileSelect = string.Empty;

	private static int indexRow = 1;

	private static int capMenu = 0;

	private Cursor curMouse = Cursor.Current;

	private CellRange cellsOld = null;

	private static int indexOld = -1;

	private static int indexCap1Last = -1;

	private ComplexBorder boder = null;

	private Color colorBG = default(Color);

	private Form frmControl = new Form();

	public string formName = string.Empty;

	public Control frmChild = new Control();

	private static int colA = HTKKFile.ColumnLettersToIndex("B");

	private static int colB = HTKKFile.ColumnLettersToIndex("C");

	private static int colC = HTKKFile.ColumnLettersToIndex("D");

	private static int colD = HTKKFile.ColumnLettersToIndex("E");

	private static int colE = HTKKFile.ColumnLettersToIndex("F");

	private static int colF = HTKKFile.ColumnLettersToIndex("G");

	private static int colG = HTKKFile.ColumnLettersToIndex("H");

	private static int colH = HTKKFile.ColumnLettersToIndex("I");

	private Image out_pluscol1 = Image.FromFile("..\\Pictures\\out_pluscol1.bmp");

	private Image out_pluscol2 = Image.FromFile("..\\Pictures\\out_pluscol2.bmp");

	private Image out_pluscol3 = Image.FromFile("..\\Pictures\\out_pluscol3.bmp");

	private Image out_oneline1 = Image.FromFile("..\\Pictures\\out_oneline1.bmp");

	private Image out_subline1 = Image.FromFile("..\\Pictures\\out_subline1.bmp");

	private Image out_subline3 = Image.FromFile("..\\Pictures\\out_subline3.bmp");

	private Image out_minuscol1 = Image.FromFile("..\\Pictures\\out_minuscol1.bmp");

	private Image out_minuscol2 = Image.FromFile("..\\Pictures\\out_minuscol2.bmp");

	private Image out_minuscol3 = Image.FromFile("..\\Pictures\\out_minuscol3.bmp");

	private IContainer components = null;

	private ColumnHeader columnHeader2;

	private ColumnHeader columnHeader5;

	private ColumnHeader columnHeaderName;

	private ColumnHeader columnHeader4;

	private ColumnHeader columnHeader3;

	private ColumnHeader columnHeader1;

	public StatusStrip statusStrip;

	public ToolStripStatusLabel toolStripStatusMST;

	public ToolStripStatusLabel toolStripStatusUserName;

	private ToolStripStatusLabel toolStripStatusLabel5;

	public ToolStripStatusLabel toolStripStatusTimeLogin;

	private ToolStripStatusLabel txtVersion;

	internal StatusStrip statusStrip1;

	private ToolStripStatusLabel toolStripStatusfooter;

	private FpSpread winSpMenu;

	private SheetView winSpMenu_Sheet1;

	public Panel panelMenu;

	public Panel panelPeriod;

	private Label lblTitle_Version_bg;

	public frmHTKKSystems()
	{
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0046: Expected O, but got Unknown
		InitializeComponent();
	}

	private void frmHTKKSystems_Load(object sender, EventArgs e)
	{
		//IL_010c: Unknown result type (might be due to invalid IL or missing references)
		colorBG = Color.FromArgb(244, 238, 202);
		ComplexBorderSide complexBorderSide = new ComplexBorderSide(ComplexBorderSideStyle.ThinLine);
		complexBorderSide = new ComplexBorderSide(ComplexBorderSideStyle.ThinLine, colorBG);
		boder = new ComplexBorder(complexBorderSide, complexBorderSide, complexBorderSide, complexBorderSide);
		((Control)this).Text = "Hỗ trợ kê khai thuế -  Phiên bản " + WinSettings.VersionApp;
		((Control)lblTitle_Version_bg).Text = "Phiên bản " + WinSettings.VersionApp;
		((Control)panelPeriod).Size = new Size(1, 1);
		((Control)panelPeriod).Location = new Point(((Control)this).Width, ((Control)this).Height);
		WinSettings.XmlMessage.LoadXml(HTKKFile.LoadFile("Message.xml"));
		((ToolStripItem)toolStripStatusMST).Text = WinSettings.HTKKHeaderInfo.mst + " : " + WinSettings.HTKKHeaderInfo.tenNNT.Replace("&", "&&");
		int num = (((Control)this).Width - ((ToolStripItem)toolStripStatusfooter).Width) / 2;
		((ToolStripItem)toolStripStatusfooter).Margin = new Padding(num, 3, 0, 2);
		curMouse = Cursor.Current;
		LoadMenu();
		WinSettings.myHTKKSystems = this;
	}

	private void LoadMenu()
	{
		//IL_015e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0164: Expected O, but got Unknown
		//IL_03b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_03bf: Expected O, but got Unknown
		winSpMenu.EditModePermanent = true;
		winSpMenu.AllowCellOverflow = false;
		GridLine horizontalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
		GridLine verticalGridLine = new GridLine(GridLineType.Flat, Color.Empty);
		winSpMenu.Sheets[0].HorizontalGridLine = horizontalGridLine;
		winSpMenu.Sheets[0].VerticalGridLine = verticalGridLine;
		winSpMenu.Sheets[0].RowHeader.Visible = false;
		winSpMenu.Sheets[0].ColumnHeader.Visible = false;
		winSpMenu.AllowEditOverflow = true;
		winSpMenu.AllowColumnMove = false;
		winSpMenu.AllowRowMove = false;
		winSpMenu.AllowDragDrop = false;
		winSpMenu.AllowSheetMove = false;
		winSpMenu.EditModeReplace = false;
		winSpMenu.RowSplitBoxPolicy = SplitBoxPolicy.Never;
		winSpMenu.HorizontalScrollBarPolicy = ScrollBarPolicy.Never;
		winSpMenu.VerticalScrollBarPolicy = ScrollBarPolicy.AsNeeded;
		winSpMenu.FocusRenderer = null;
		winSpMenu.Sheets[0].GrayAreaBackColor = colorBG;
		XmlNodeList elementsByTagName = ((XmlDocument)WinSettings.xmlDocMenu).GetElementsByTagName("Menu");
		foreach (XmlNode item in elementsByTagName)
		{
			XmlNode val = item;
			AddNode(((XmlNode)val.Attributes["ParentID"]).Value, ((XmlNode)val.Attributes["ID"]).Value, ((XmlNode)val.Attributes["Caption"]).Value, ((XmlNode)val.Attributes["Icon"]).Value, ((XmlNode)val.Attributes["PopID"]).Value);
		}
		((Control)winSpMenu).Cursor = curMouse;
		ExpandMenu(WinSettings.treeIndex);
		GeneralCellType generalCellType = new GeneralCellType();
		generalCellType.BackgroundImage = new Picture(out_pluscol1, RenderStyle.Stretch);
		generalCellType.CancelEditing();
		winSpMenu.Sheets[0].Cells[0, colA - 1].CellType = generalCellType;
		GeneralCellType generalCellType2 = new GeneralCellType();
		generalCellType2.BackgroundImage = new Picture(out_pluscol3, RenderStyle.Stretch);
		generalCellType2.CancelEditing();
		winSpMenu.Sheets[0].Cells[indexCap1Last, colA - 1].CellType = generalCellType2;
		GeneralCellType generalCellType3 = new GeneralCellType();
		generalCellType3.BackgroundImage = new Picture(out_oneline1, RenderStyle.Stretch);
		generalCellType3.CancelEditing();
		winSpMenu.Sheets[0].Cells[winSpMenu.Sheets[0].Rows.Count - 1, colA - 1].CellType = generalCellType3;
		winSpMenu.Sheets[0].Cells[winSpMenu.Sheets[0].Rows.Count - 1, colB - 1].CellType = generalCellType3;
		winSpMenu.Sheets[0].SetActiveCell(0, colH - 1);
		string text = "";
		for (int i = 0; i < 100; i++)
		{
			bool flag = true;
			foreach (XmlNode item2 in elementsByTagName)
			{
				XmlNode val = item2;
				string value = ((XmlNode)val.Attributes["ID"]).Value;
				if (value.Length == 2 && ConvertHelpers.ConvertToInt(value) == i)
				{
					flag = false;
				}
			}
			if (flag)
			{
				text = text + i + "; ";
			}
		}
		string text2 = "0";
	}

	private void AddNode(string parentNode, string nodeValue, string nodeText, string ImageURL, string PopID)
	{
		//IL_04e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_04f6: Expected O, but got Unknown
		//IL_04f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_04fb: Expected O, but got Unknown
		//IL_0ab4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ac3: Expected O, but got Unknown
		//IL_0abe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ac8: Expected O, but got Unknown
		//IL_082c: Unknown result type (might be due to invalid IL or missing references)
		//IL_083c: Expected O, but got Unknown
		//IL_0837: Unknown result type (might be due to invalid IL or missing references)
		//IL_0841: Expected O, but got Unknown
		//IL_07e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_07f6: Expected O, but got Unknown
		//IL_07f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_07fb: Expected O, but got Unknown
		try
		{
			if (WinSettings.treeIndex == -1 && nodeValue == "01")
			{
				WinSettings.treeIndex = indexRow - 1;
			}
			if (indexCap1Last == -1 && nodeValue == "102")
			{
				indexCap1Last = indexRow - 1;
			}
			winSpMenu.Sheets[0].Rows.Count = indexRow;
			winSpMenu.Sheets[0].Cells[indexRow - 1, 0].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colA - 1].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colB - 1].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colD - 1].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colE - 1].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colF - 1].BackColor = colorBG;
			winSpMenu.Sheets[0].Cells[indexRow - 1, 0].Border = boder;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colA - 1].Border = boder;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colB - 1].Border = boder;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].Border = boder;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colD - 1].Border = boder;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colE - 1].Border = boder;
			winSpMenu.Sheets[0].Cells[indexRow - 1, colF - 1].Border = boder;
			GeneralCellType generalCellType = new GeneralCellType();
			GC.Collect();
			GC.WaitForPendingFinalizers();
			if (HTKKFormater.isNullString(parentNode) && generalCellType != null)
			{
				generalCellType = new GeneralCellType();
				Image i = ((indexRow != 1) ? out_pluscol2 : out_pluscol1);
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[indexRow - 1, colA - 1].CellType = generalCellType;
				generalCellType = new GeneralCellType();
				if (ImageURL != "" && File.Exists(ImageURL))
				{
					generalCellType.BackgroundImage = new Picture(Image.FromFile(ImageURL), RenderStyle.Stretch);
					generalCellType.CancelEditing();
				}
				winSpMenu.Sheets[0].Cells[indexRow - 1, colB - 1].CellType = generalCellType;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].ColumnSpan = 2;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].Value = nodeText;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].Font = new Font(new FontFamily("Tahoma"), 8f, (FontStyle)1);
				winSpMenu.Sheets[0].Cells[indexRow - 1, colE - 1].Value = nodeValue;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colF - 1].Value = 1;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colG - 1].Value = parentNode;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colH - 1].Value = PopID;
				winSpMenu.Sheets[0].Rows[indexRow - 1].Locked = true;
				indexRow++;
			}
			else if (nodeValue.IndexOf("_") > -1 && !HTKKFormater.isNullString(parentNode))
			{
				generalCellType = new GeneralCellType();
				Image i = ((indexRow != 1) ? out_subline1 : out_subline1);
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[indexRow - 1, colA - 1].CellType = generalCellType;
				generalCellType = new GeneralCellType();
				i = out_pluscol2;
				if (indexRow == 1)
				{
					i = out_pluscol1;
				}
				if (!HTKKFormater.isNullString(PopID))
				{
					i = out_subline3;
				}
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch, colorBG);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[indexRow - 1, colB - 1].CellType = generalCellType;
				generalCellType = new GeneralCellType();
				if (ImageURL != "" && File.Exists(ImageURL))
				{
					generalCellType.BackgroundImage = new Picture(Image.FromFile(ImageURL), RenderStyle.Stretch);
					generalCellType.CancelEditing();
				}
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].CellType = generalCellType;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colD - 1].Value = nodeText;
				if (!HTKKFormater.isNullString(PopID))
				{
					winSpMenu.Sheets[0].Cells[indexRow - 1, colD - 1].Font = new Font(new FontFamily("Tahoma"), 8f);
				}
				else
				{
					winSpMenu.Sheets[0].Cells[indexRow - 1, colD - 1].Font = new Font(new FontFamily("Tahoma"), 8f, (FontStyle)1);
				}
				winSpMenu.Sheets[0].Cells[indexRow - 1, colE - 1].Value = nodeValue;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colF - 1].Value = 2;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colG - 1].Value = parentNode;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colH - 1].Value = PopID;
				winSpMenu.Sheets[0].Rows[indexRow - 1].Locked = true;
				indexRow++;
			}
			else if (!HTKKFormater.isNullString(PopID))
			{
				generalCellType = new GeneralCellType();
				Image i = ((indexRow != 1) ? out_subline1 : out_subline1);
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[indexRow - 1, colA - 1].CellType = generalCellType;
				generalCellType = new GeneralCellType();
				i = ((indexRow != 1) ? out_subline3 : out_subline1);
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[indexRow - 1, colB - 1].CellType = generalCellType;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].ColumnSpan = 2;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].Value = nodeText;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colC - 1].Font = new Font(new FontFamily("Tahoma"), 8f);
				winSpMenu.Sheets[0].Cells[indexRow - 1, colE - 1].Value = nodeValue;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colF - 1].Value = 3;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colG - 1].Value = parentNode;
				winSpMenu.Sheets[0].Cells[indexRow - 1, colH - 1].Value = PopID;
				winSpMenu.Sheets[0].Rows[indexRow - 1].Locked = true;
				if ("I5~K1~K6~K8~K9~M7~N1~N9~O8~O9~".IndexOf(nodeValue) >= 0)
				{
					winSpMenu.Sheets[0].Rows[indexRow - 1].Height = 0f;
				}
				if ("35~L5~L6~29~G6~37~B8~38~".IndexOf(nodeValue) >= 0)
				{
					winSpMenu.Sheets[0].Rows[indexRow - 1].Height = 0f;
				}
				indexRow++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void trvMenu_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
	{
	}

	private void form_DoWorkBackup(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		WinDoWorkEvent.WinDoWorkBackup(fileSelect);
	}

	private void form_DoWorkRestore(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		WinDoWorkEvent.WinDoWorkRestore(fileSelect);
	}

	private void form_DoWorkConvert(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		WinDoWorkEvent.WinDoWorkConvert(fileSelect);
	}

	public void LoadListMessage()
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		XmlDocument val = new XmlDocument();
		val.Load(Application.StartupPath + "\\Message.xml");
		val = null;
	}

	public void ExpandMenu(int indexMenu)
	{
		try
		{
			if (indexMenu == -1)
			{
				return;
			}
			string text = winSpMenu.Sheets[0].Cells[indexMenu, colG - 1].Text;
			string text2 = winSpMenu.Sheets[0].Cells[indexMenu, colE - 1].Text;
			for (int i = 0; i < winSpMenu.Sheets[0].RowCount; i++)
			{
				string text3 = winSpMenu.Sheets[0].Cells[i, colG - 1].Text;
				string text4 = winSpMenu.Sheets[0].Cells[i, colE - 1].Text;
				if (HTKKFormater.isNullString(text4))
				{
					continue;
				}
				if (text == text3 || text == text4 || text2 == text3)
				{
					winSpMenu.Sheets[0].Rows[i].Visible = true;
				}
				else
				{
					string text5 = winSpMenu.Sheets[0].Cells[i, colF - 1].Text;
					string text6 = winSpMenu.Sheets[0].Cells[i, colH - 1].Text;
					if (text5 != "1" && !HTKKFormater.isNullString(text6))
					{
						winSpMenu.Sheets[0].Rows[i].Visible = false;
					}
				}
				if (text == text4)
				{
					ExpandMenu(i);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void ChildFormClosed(object sender, FormClosedEventArgs e)
	{
		FormProcess.ReLoadfrmSystems();
	}

	private void ShowBangKe(string sFileName)
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			string text = Path.GetTempPath() + sFileName;
			File.Copy(HTKKConfig.INTERFACE_TEMPLATES + "\\excel\\" + sFileName, text, overwrite: true);
			Process.Start(text);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage = new frmMessage("0433", "3", "", "Thông báo");
			((Form)frmMessage).ShowDialog();
		}
	}

	private void frmHTKKSystems_FormClosed(object sender, FormClosedEventArgs e)
	{
		Application.Exit();
	}

	private void winSpMenu_MouseMove(object sender, MouseEventArgs e)
	{
		//IL_012f: Unknown result type (might be due to invalid IL or missing references)
		//IL_013e: Expected O, but got Unknown
		//IL_0139: Unknown result type (might be due to invalid IL or missing references)
		//IL_0143: Expected O, but got Unknown
		try
		{
			if (cellsOld != null)
			{
				winSpMenu.Sheets[0].Cells[cellsOld.Row, colD - 1].Border = boder;
				winSpMenu.Sheets[0].Cells[cellsOld.Row, colC - 1].Border = boder;
			}
			winSpMenu.SetCursor(CursorType.Normal, curMouse);
			winSpMenu.SetCursor(CursorType.LockedCell, curMouse);
			CellRange cellFromPixel = winSpMenu.GetCellFromPixel(0, 0, e.X, e.Y);
			if (cellFromPixel != null && cellFromPixel.Row >= 0 && cellFromPixel.Row <= winSpMenu.Sheets[0].RowCount)
			{
				winSpMenu.Sheets[0].Rows[cellFromPixel.Row].Font = new Font(new FontFamily("Tahoma"), 8f);
				Color color = default(Color);
				color = Color.FromArgb(168, 9, 13);
				string text = winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colF - 1].Text;
				string text2 = winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colH - 1].Text;
				ComplexBorderSide complexBorderSide = new ComplexBorderSide(ComplexBorderSideStyle.Dashed, color);
				if (text == "1")
				{
					complexBorderSide = new ComplexBorderSide(ComplexBorderSideStyle.ThinLine, color);
					ComplexBorder border = new ComplexBorder(null, complexBorderSide, null, complexBorderSide);
					winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colD - 1].Border = border;
					winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colC - 1].Border = border;
				}
				else if (text == "2" && HTKKFormater.isNullString(text2))
				{
					complexBorderSide = new ComplexBorderSide(ComplexBorderSideStyle.ThinLine, color);
					ComplexBorder border = new ComplexBorder(null, complexBorderSide, null, complexBorderSide);
					winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colD - 1].Border = border;
				}
				else
				{
					ComplexBorder border = new ComplexBorder(null, complexBorderSide, null, complexBorderSide);
					winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colD - 1].Border = border;
					winSpMenu.Sheets[0].Cells[cellFromPixel.Row, colC - 1].Border = border;
				}
				cellsOld = cellFromPixel;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winSpMenu_MouseHover(object sender, EventArgs e)
	{
		try
		{
			Cursor current = Cursor.Current;
			winSpMenu.SetCursor(CursorType.Normal, curMouse);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winSpMenu_MouseDown(object sender, MouseEventArgs e)
	{
	}

	private void winSpMenu_CellClick(object sender, CellClickEventArgs e)
	{
		//IL_0612: Unknown result type (might be due to invalid IL or missing references)
		//IL_061c: Expected O, but got Unknown
		//IL_0606: Unknown result type (might be due to invalid IL or missing references)
		//IL_09f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a17: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a37: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a57: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a77: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a97: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ab7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0af5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0afc: Expected O, but got Unknown
		//IL_0b5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b62: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b64: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b67: Invalid comparison between Unknown and I4
		//IL_0bbb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bc2: Expected O, but got Unknown
		//IL_0bfd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c02: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c04: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c07: Invalid comparison between Unknown and I4
		//IL_0c78: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c7f: Expected O, but got Unknown
		//IL_0cba: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cbf: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cc1: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cc4: Invalid comparison between Unknown and I4
		//IL_0f31: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d54: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (e.Row > winSpMenu.Sheets[0].RowCount || e.Column > winSpMenu.Sheets[0].ColumnCount || e.Row < 0 || e.Column < 0)
			{
				return;
			}
			GeneralCellType generalCellType = new GeneralCellType();
			string text = "";
			string text2 = "";
			if (indexOld != -1 && indexOld != e.Row)
			{
				if (indexOld == 0)
				{
					generalCellType = new GeneralCellType();
					Image i = out_pluscol1;
					generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
					generalCellType.CancelEditing();
					winSpMenu.Sheets[0].Cells[0, colA - 1].CellType = generalCellType;
				}
				else if (indexOld == indexCap1Last)
				{
					generalCellType = new GeneralCellType();
					Image i = out_pluscol3;
					generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
					generalCellType.CancelEditing();
					winSpMenu.Sheets[0].Cells[indexCap1Last, colA - 1].CellType = generalCellType;
				}
				else
				{
					generalCellType = new GeneralCellType();
					text = winSpMenu.Sheets[0].Cells[indexOld, colF - 1].Text;
					text2 = winSpMenu.Sheets[0].Cells[indexOld, colH - 1].Text;
					if (text == "1")
					{
						Image i = out_pluscol2;
						generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
						generalCellType.CancelEditing();
						winSpMenu.Sheets[0].Cells[indexOld, colA - 1].CellType = generalCellType;
					}
					if (text == "2" && HTKKFormater.isNullString(text2))
					{
						Image i = out_pluscol2;
						generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
						generalCellType.CancelEditing();
						winSpMenu.Sheets[0].Cells[indexOld, colB - 1].CellType = generalCellType;
					}
				}
			}
			generalCellType = new GeneralCellType();
			text = winSpMenu.Sheets[0].Cells[e.Row, colF - 1].Text;
			text2 = winSpMenu.Sheets[0].Cells[e.Row, colH - 1].Text;
			if (e.Row == 0)
			{
				generalCellType = new GeneralCellType();
				Image i = out_minuscol1;
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[0, colA - 1].CellType = generalCellType;
			}
			else if (e.Row == indexCap1Last)
			{
				generalCellType = new GeneralCellType();
				Image i = out_minuscol3;
				generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
				generalCellType.CancelEditing();
				winSpMenu.Sheets[0].Cells[indexCap1Last, colA - 1].CellType = generalCellType;
			}
			else
			{
				if (text == "1")
				{
					Image i = out_minuscol2;
					generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
					generalCellType.CancelEditing();
					winSpMenu.Sheets[0].Cells[e.Row, colA - 1].CellType = generalCellType;
				}
				if (text == "2" && HTKKFormater.isNullString(text2))
				{
					Image i = out_minuscol2;
					generalCellType.BackgroundImage = new Picture(i, RenderStyle.Stretch);
					generalCellType.CancelEditing();
					winSpMenu.Sheets[0].Cells[e.Row, colB - 1].CellType = generalCellType;
				}
			}
			indexOld = e.Row;
			string text3 = winSpMenu.Sheets[0].Cells[e.Row, colE - 1].Text;
			WinSettings.treeIndex = e.Row;
			ExpandMenu(WinSettings.treeIndex);
			XmlNodeList val = ((XmlNode)WinSettings.xmlDocMenu).SelectNodes("/Root/Menu[@ID='" + text3 + "']");
			WinSettings.xmlMenuNode = val[0];
			WinSettings.MenuID = text3;
			HTKKPeriodInfo tTPeriodInfo = WinSettings.TTPeriodInfo;
			string versionXML = (WinSettings.ToKhaiInfo.VersionXML = ((WinSettings.xmlMenuNode.Attributes["VersionXML"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["VersionXML"]).Value));
			tTPeriodInfo.VersionXML = versionXML;
			WinSettings.ToKhaiInfo.VersionKBSXML = ((WinSettings.xmlMenuNode.Attributes["VersionKHBS"] == null) ? "" : ((XmlNode)WinSettings.xmlMenuNode.Attributes["VersionKHBS"]).Value);
			if (!GeneralFunctions.CheckThongTinNNT(WinSettings.HTKKHeaderInfo) && WinSettings.MenuID.IndexOf("_") < 0 && WinSettings.MenuID.Length == 2)
			{
				frmMessage frmMessage = new frmMessage("0450", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			frmChild = new Control();
			formName = "";
			switch (WinSettings.MenuID)
			{
			case "100_4":
				Application.Exit();
				break;
			case "100_5":
			{
				WinSettings.chuyenMSTKhac = true;
				((Control)this).Hide();
				frmFormStart frmFormStart = new frmFormStart();
				((Control)frmFormStart).Show();
				break;
			}
			case "100_1":
			{
				Functions.ResetInfo(WinSettings.DicSheet, WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo);
				WinSettings.InstanceClass = WinDoWorkEvent.GetInstance(((XmlNode)WinSettings.xmlMenuNode.Attributes["Class"]).Value.ToString());
				WinSettings.ToKhaiInfo = WinFnProcess.GetToKhaiInfo(WinSettings.xmlMenuNode, WinSettings.TTPeriodInfo);
				frmWinToKhai frmWinToKhai2 = new frmWinToKhai(WinSettings.xmlMenuNode);
				((Control)frmWinToKhai2).Text = HTKKFile.GetMessage("0436");
				((Control)this).Hide();
				((Form)frmWinToKhai2).ShowDialog();
				break;
			}
			case "100_7":
			{
				frmRegistrationCatalogues frmRegistrationCatalogues4 = new frmRegistrationCatalogues(WinSettings.xmlMenuNode);
				((Control)this).Hide();
				((Form)frmRegistrationCatalogues4).ShowDialog();
				break;
			}
			case "100_13":
			{
				frmRegistrationCatalogues frmRegistrationCatalogues3 = new frmRegistrationCatalogues(WinSettings.xmlMenuNode);
				((Control)this).Hide();
				((Form)frmRegistrationCatalogues3).ShowDialog();
				break;
			}
			case "100_14":
			{
				frmRegistrationCatalogues frmRegistrationCatalogues2 = new frmRegistrationCatalogues(WinSettings.xmlMenuNode);
				((Control)this).Hide();
				((Form)frmRegistrationCatalogues2).ShowDialog();
				break;
			}
			case "100_15":
			{
				frmRegistrationCatalogues frmRegistrationCatalogues7 = new frmRegistrationCatalogues(WinSettings.xmlMenuNode);
				((Control)this).Hide();
				((Form)frmRegistrationCatalogues7).ShowDialog();
				break;
			}
			case "100_9":
			{
				frmRegistrationCatalogues frmRegistrationCatalogues6 = new frmRegistrationCatalogues(WinSettings.xmlMenuNode);
				((Control)this).Hide();
				((Form)frmRegistrationCatalogues6).ShowDialog();
				break;
			}
			case "100_10":
			{
				frmRegistrationCatalogues frmRegistrationCatalogues5 = new frmRegistrationCatalogues(WinSettings.xmlMenuNode);
				((Control)this).Hide();
				((Form)frmRegistrationCatalogues5).ShowDialog();
				break;
			}
			case "104":
				frmControl = (Form)(object)new frmTraCuuToKhai(WinSettings.xmlMenuNode);
				AddControl(frmControl);
				formName = ((Control)frmControl).Name;
				break;
			case "100_2":
			{
				SaveFileDialog val7 = new SaveFileDialog();
				((FileDialog)val7).InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
				((FileDialog)val7).Title = "Nhâ\u0323p tên file đê\u0309 sao lưu";
				((FileDialog)val7).DefaultExt = "zip";
				((FileDialog)val7).Filter = "zip files|*.zip";
				((FileDialog)val7).FileName = "Backup_" + DateTime.Now.ToString("yyyy_MM_dd_HHmmssfff") + ".zip";
				DialogResult val8 = ((CommonDialog)val7).ShowDialog();
				if ((int)val8 == 1)
				{
					fileSelect = ((FileDialog)val7).FileName;
					ProgressForm progressForm2 = new ProgressForm("Đang sao lưu dư\u0303 liê\u0323u...");
					progressForm2.Argument = false;
					progressForm2.DoWork += form_DoWorkBackup;
					WinSpreadProcess.StartForm(progressForm2);
				}
				break;
			}
			case "100_3":
			{
				OpenFileDialog val5 = new OpenFileDialog();
				((FileDialog)val5).DefaultExt = "zip";
				((FileDialog)val5).Filter = "Zip file (*.zip)|*.zip";
				((FileDialog)val5).AddExtension = true;
				((FileDialog)val5).RestoreDirectory = true;
				((FileDialog)val5).Title = "Chọn file phu\u0323c hô\u0300i dư\u0303 liê\u0323u?";
				DialogResult val6 = ((CommonDialog)val5).ShowDialog();
				if ((int)val6 == 1)
				{
					fileSelect = ((FileDialog)val5).FileName;
					ProgressForm progressForm = new ProgressForm("Đang phu\u0323c hô\u0300i dư\u0303 liê\u0323u...");
					progressForm.Argument = false;
					progressForm.DoWork += form_DoWorkRestore;
					WinSpreadProcess.StartForm(progressForm);
				}
				break;
			}
			case "100_11":
				frmControl = (Form)(object)new frmConfigUpdate();
				AddControl(frmControl);
				break;
			case "100_12":
			{
				OpenFileDialog val3 = new OpenFileDialog();
				((FileDialog)val3).DefaultExt = "zip";
				((FileDialog)val3).Filter = "Zip file (*.zip)|*.zip";
				((FileDialog)val3).AddExtension = true;
				((FileDialog)val3).RestoreDirectory = true;
				((FileDialog)val3).Title = "Đường dẫn đến file sao lưu";
				DialogResult val4 = ((CommonDialog)val3).ShowDialog();
				if ((int)val4 == 1)
				{
					fileSelect = ((FileDialog)val3).FileName;
					ProgressForm progressForm = new ProgressForm("Đang phu\u0323c hô\u0300i dư\u0303 liê\u0323u...");
					progressForm.Argument = false;
					progressForm.DoWork += form_DoWorkConvert;
					WinSpreadProcess.StartForm(progressForm);
				}
				break;
			}
			case "102_2":
				if (File.Exists(Application.StartupPath + "\\HTKK.chm"))
				{
					Process.Start(Application.StartupPath + "\\HTKK.chm");
				}
				else
				{
					MessageBox.Show("Không tìm thấy file !");
				}
				break;
			case "102_3":
				ShowBangKe("Bang_Ke_01NTNN.xls");
				break;
			case "102_4":
				ShowBangKe("Bang_Ke_05_1_QTT.xls");
				break;
			case "102_5":
				ShowBangKe("Bang_Ke_05_2_QTT.xls");
				break;
			case "102_6":
				ShowBangKe("Bang_Ke_05_3_QTT.xls");
				break;
			case "102_21":
				ShowBangKe("Bang_Ke_05DS.xls");
				break;
			case "102_7":
				ShowBangKe("Bang_Ke_01_1_XSBHDC.xls");
				break;
			case "102_9":
				ShowBangKe("Bang_Ke_16TH_DangKyCapMa.xls");
				break;
			case "102_10":
				ShowBangKe("Bao_Cao_BC26_AC.xls");
				break;
			case "102_11":
				ShowBangKe("Bangke_BC26_Mau_3_10.xls");
				break;
			case "102_12":
				ShowBangKe("Bangke_BC26_Mau_3_12.xls");
				break;
			case "102_22":
				ShowBangKe("Bao_Cao_01BC_SDHD_CNKD.xls");
				break;
			case "102_23":
				ShowBangKe("Bao_Cao_BC26_AC_BLP.xls");
				break;
			case "102_13":
				ShowBangKe("PL_01_7GTGT_TT154_MaSoThue.xls");
				break;
			case "102_14":
				ShowBangKe("PL_02MT_GTGT_TT140_MaSoThue.xls");
				break;
			case "102_15":
				ShowBangKe("PL_02GH_GTGT_TT16_MaSoThue.xls");
				break;
			case "102_16":
				ShowBangKe("PL_01GH_TNDN_TT16_MaSoThue.xls");
				break;
			case "102_17":
				ShowBangKe("BKGiaHan_TT170_MaSoThue.doc");
				break;
			case "102_18":
				ShowBangKe("PL_01MGT_TNDN_TT140_MaSoThue.xls");
				break;
			case "102_19":
				ShowBangKe("PL_01_TNDN_TT199_MaSoThue.xls");
				break;
			case "102_20":
				ShowBangKe("PL_25MGT_TNCN_TT154_MaSoThue.xls");
				break;
			case "102_24":
				ShowBangKe("bangkekhautru_01TTDB.xls");
				break;
			case "102_25":
				ShowBangKe("Bangke_01TBAC.xls");
				break;
			case "102_26":
				ShowBangKe("Bangke_01TBAC_BLP.xls");
				break;
			case "102_27":
				ShowBangKe("Bangke_03_7.xls");
				break;
			case "102_28":
				ShowBangKe("Bangke_01BCAC.xls");
				break;
			case "102_29":
				ShowBangKe("Bangke_01BCAC_BLP.xls");
				break;
			case "102_30":
				ShowBangKe("Bang_Ke_05DK_TH.xls");
				break;
			case "104_1":
			{
				frmHoTroKeKhai frmHoTroKeKhai = new frmHoTroKeKhai();
				((Form)frmHoTroKeKhai).ShowDialog();
				break;
			}
			case "104_2":
				WinDoWorkEvent.WinDoWorkSignXml();
				break;
			default:
			{
				XmlNodeList val2 = ((XmlNode)WinSettings.xmlDocMenu).SelectNodes("/Root/Menu[@ParentID='" + WinSettings.MenuID + "']");
				if (val2.Count == 0)
				{
					frmControl = (Form)(object)new frmHTKKPeriod(this, WinSettings.xmlMenuNode, 0);
					panelPeriod.BorderStyle = (BorderStyle)1;
					AddControl(frmControl);
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

	private void statusStrip1_Resize(object sender, EventArgs e)
	{
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		int num = (((Control)this).Width - ((ToolStripItem)toolStripStatusfooter).Width) / 2;
		((ToolStripItem)toolStripStatusfooter).Margin = new Padding(num, 3, 0, 2);
		ResizeForm((Control)(object)panelPeriod, frmControl);
	}

	public void AddControl(Form obj)
	{
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Expected O, but got Unknown
		try
		{
			FormProcess.BeginControlUpdate((Control)(object)this);
			((Control)panelMenu).Visible = false;
			((Control)panelPeriod).Controls.Clear();
			((Control)panelPeriod).Visible = true;
			obj.TopLevel = false;
			obj.FormBorderStyle = (FormBorderStyle)0;
			obj.FormClosed += new FormClosedEventHandler(ChildFormClosed);
			((Control)obj).Show();
			((Control)panelPeriod).Controls.AddRange((Control[])(object)new Control[1] { (Control)obj });
			((Control)panelPeriod).Width = ((Control)obj).Width;
			((Control)panelPeriod).Height = ((Control)obj).Height;
			((Control)obj).Dock = (DockStyle)5;
			ResizeForm((Control)(object)panelPeriod, obj);
			FormProcess.EndControlUpdate((Control)(object)this);
		}
		catch (Exception)
		{
		}
	}

	private void getListControl(List<Control> listCtl, Control pControl)
	{
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002b: Expected O, but got Unknown
		try
		{
			if (pControl.HasChildren)
			{
				foreach (Control item in (ArrangedElementCollection)pControl.Controls)
				{
					Control pControl2 = item;
					getListControl(listCtl, pControl2);
				}
			}
			if (pControl.Visible)
			{
				listCtl.Add(pControl);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ResizeForm(Control control, Form obj)
	{
		try
		{
			int x = (((Control)this).Width - ((Control)obj).Width) / 2;
			int y = (((Control)this).Height - ((Control)obj).Height) / 2;
			control.Location = new Point(x, y);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmHTKKSystems_Resize(object sender, EventArgs e)
	{
		try
		{
			int height = ((Control)statusStrip).Height;
			int height2 = ((Control)statusStrip1).Height;
			((Control)panelMenu).Height = ((Control)this).Height - height - height2 - 2;
			((Control)winSpMenu).Height = ((Control)this).Height - height - height2 - 2;
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
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		//IL_003d: Expected O, but got Unknown
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Expected O, but got Unknown
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
		//IL_0096: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a0: Expected O, but got Unknown
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Expected O, but got Unknown
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cc: Expected O, but got Unknown
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Expected O, but got Unknown
		//IL_00d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e2: Expected O, but got Unknown
		//IL_023a: Unknown result type (might be due to invalid IL or missing references)
		//IL_02cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0467: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_04de: Expected O, but got Unknown
		//IL_04fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0592: Unknown result type (might be due to invalid IL or missing references)
		//IL_061e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0628: Expected O, but got Unknown
		//IL_064e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0658: Expected O, but got Unknown
		//IL_06c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ca: Expected O, but got Unknown
		//IL_06e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_06f1: Expected O, but got Unknown
		//IL_070e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0718: Expected O, but got Unknown
		//IL_0735: Unknown result type (might be due to invalid IL or missing references)
		//IL_073f: Expected O, but got Unknown
		//IL_075c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0766: Expected O, but got Unknown
		//IL_0783: Unknown result type (might be due to invalid IL or missing references)
		//IL_078d: Expected O, but got Unknown
		//IL_07aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_07b4: Expected O, but got Unknown
		//IL_07d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_07db: Expected O, but got Unknown
		//IL_07f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0802: Expected O, but got Unknown
		//IL_081f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0829: Expected O, but got Unknown
		//IL_0846: Unknown result type (might be due to invalid IL or missing references)
		//IL_0850: Expected O, but got Unknown
		//IL_086d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0877: Expected O, but got Unknown
		//IL_0894: Unknown result type (might be due to invalid IL or missing references)
		//IL_089e: Expected O, but got Unknown
		//IL_08bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_08c5: Expected O, but got Unknown
		//IL_08e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_08ec: Expected O, but got Unknown
		//IL_0909: Unknown result type (might be due to invalid IL or missing references)
		//IL_0913: Expected O, but got Unknown
		//IL_0930: Unknown result type (might be due to invalid IL or missing references)
		//IL_093a: Expected O, but got Unknown
		//IL_0957: Unknown result type (might be due to invalid IL or missing references)
		//IL_0961: Expected O, but got Unknown
		//IL_097e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0988: Expected O, but got Unknown
		//IL_09a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_09af: Expected O, but got Unknown
		//IL_09cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_09d6: Expected O, but got Unknown
		//IL_09f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_09fd: Expected O, but got Unknown
		//IL_0a1a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a24: Expected O, but got Unknown
		//IL_0a41: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a4b: Expected O, but got Unknown
		//IL_0a68: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a72: Expected O, but got Unknown
		//IL_0a8f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a99: Expected O, but got Unknown
		//IL_0ab6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ac0: Expected O, but got Unknown
		//IL_0add: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ae7: Expected O, but got Unknown
		//IL_0b04: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b0e: Expected O, but got Unknown
		//IL_0b2b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b35: Expected O, but got Unknown
		//IL_0b52: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b5c: Expected O, but got Unknown
		//IL_0b79: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b83: Expected O, but got Unknown
		//IL_0ba0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0baa: Expected O, but got Unknown
		//IL_0bc7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bd1: Expected O, but got Unknown
		//IL_0bee: Unknown result type (might be due to invalid IL or missing references)
		//IL_0bf8: Expected O, but got Unknown
		//IL_0c15: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c1f: Expected O, but got Unknown
		//IL_0c3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c47: Expected O, but got Unknown
		//IL_0c65: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c6f: Expected O, but got Unknown
		//IL_0c8d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c97: Expected O, but got Unknown
		//IL_0cb5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0cbf: Expected O, but got Unknown
		//IL_0cdd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ce7: Expected O, but got Unknown
		//IL_0d05: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d0f: Expected O, but got Unknown
		//IL_0d2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d37: Expected O, but got Unknown
		//IL_0d55: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d5f: Expected O, but got Unknown
		//IL_0d7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d87: Expected O, but got Unknown
		//IL_0da5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0daf: Expected O, but got Unknown
		//IL_0dcd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0dd7: Expected O, but got Unknown
		//IL_0df5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0dff: Expected O, but got Unknown
		//IL_0e1d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e27: Expected O, but got Unknown
		//IL_0e45: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e4f: Expected O, but got Unknown
		//IL_0e6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e77: Expected O, but got Unknown
		//IL_0e95: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e9f: Expected O, but got Unknown
		//IL_0ebd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ec7: Expected O, but got Unknown
		//IL_0ee5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0eef: Expected O, but got Unknown
		//IL_0f0d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f17: Expected O, but got Unknown
		//IL_0f35: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f3f: Expected O, but got Unknown
		//IL_0f5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f67: Expected O, but got Unknown
		//IL_0f85: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f8f: Expected O, but got Unknown
		//IL_0fad: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fb7: Expected O, but got Unknown
		//IL_0fd5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fdf: Expected O, but got Unknown
		//IL_0ffd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1007: Expected O, but got Unknown
		//IL_1025: Unknown result type (might be due to invalid IL or missing references)
		//IL_102f: Expected O, but got Unknown
		//IL_104d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1057: Expected O, but got Unknown
		//IL_1075: Unknown result type (might be due to invalid IL or missing references)
		//IL_107f: Expected O, but got Unknown
		//IL_109d: Unknown result type (might be due to invalid IL or missing references)
		//IL_10a7: Expected O, but got Unknown
		//IL_10c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_10cf: Expected O, but got Unknown
		//IL_10ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_10f7: Expected O, but got Unknown
		//IL_1115: Unknown result type (might be due to invalid IL or missing references)
		//IL_111f: Expected O, but got Unknown
		//IL_113d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1147: Expected O, but got Unknown
		//IL_1165: Unknown result type (might be due to invalid IL or missing references)
		//IL_116f: Expected O, but got Unknown
		//IL_118d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1197: Expected O, but got Unknown
		//IL_11b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_11bf: Expected O, but got Unknown
		//IL_11dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_11e7: Expected O, but got Unknown
		//IL_1205: Unknown result type (might be due to invalid IL or missing references)
		//IL_120f: Expected O, but got Unknown
		//IL_122d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1237: Expected O, but got Unknown
		//IL_1255: Unknown result type (might be due to invalid IL or missing references)
		//IL_125f: Expected O, but got Unknown
		//IL_127d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1287: Expected O, but got Unknown
		//IL_12a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_12af: Expected O, but got Unknown
		//IL_12cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_12d7: Expected O, but got Unknown
		//IL_12f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_12ff: Expected O, but got Unknown
		//IL_131d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1327: Expected O, but got Unknown
		//IL_1345: Unknown result type (might be due to invalid IL or missing references)
		//IL_134f: Expected O, but got Unknown
		//IL_136d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1377: Expected O, but got Unknown
		//IL_1395: Unknown result type (might be due to invalid IL or missing references)
		//IL_139f: Expected O, but got Unknown
		//IL_13bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_13c7: Expected O, but got Unknown
		//IL_13e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_13ef: Expected O, but got Unknown
		//IL_140d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1417: Expected O, but got Unknown
		//IL_1435: Unknown result type (might be due to invalid IL or missing references)
		//IL_143f: Expected O, but got Unknown
		//IL_145d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1467: Expected O, but got Unknown
		//IL_1485: Unknown result type (might be due to invalid IL or missing references)
		//IL_148f: Expected O, but got Unknown
		//IL_14ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_14b7: Expected O, but got Unknown
		//IL_14d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_14df: Expected O, but got Unknown
		//IL_14fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1507: Expected O, but got Unknown
		//IL_1525: Unknown result type (might be due to invalid IL or missing references)
		//IL_152f: Expected O, but got Unknown
		//IL_154d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1557: Expected O, but got Unknown
		//IL_1575: Unknown result type (might be due to invalid IL or missing references)
		//IL_157f: Expected O, but got Unknown
		//IL_159d: Unknown result type (might be due to invalid IL or missing references)
		//IL_15a7: Expected O, but got Unknown
		//IL_15c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_15cf: Expected O, but got Unknown
		//IL_15ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_15f7: Expected O, but got Unknown
		//IL_1615: Unknown result type (might be due to invalid IL or missing references)
		//IL_161f: Expected O, but got Unknown
		//IL_163d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1647: Expected O, but got Unknown
		//IL_1665: Unknown result type (might be due to invalid IL or missing references)
		//IL_166f: Expected O, but got Unknown
		//IL_168d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1697: Expected O, but got Unknown
		//IL_16b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_16bf: Expected O, but got Unknown
		//IL_16dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_16e7: Expected O, but got Unknown
		//IL_1705: Unknown result type (might be due to invalid IL or missing references)
		//IL_170f: Expected O, but got Unknown
		//IL_172d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1737: Expected O, but got Unknown
		//IL_1755: Unknown result type (might be due to invalid IL or missing references)
		//IL_175f: Expected O, but got Unknown
		//IL_177d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1787: Expected O, but got Unknown
		//IL_17a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_17af: Expected O, but got Unknown
		//IL_17cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_17d7: Expected O, but got Unknown
		//IL_17f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_17ff: Expected O, but got Unknown
		//IL_181d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1827: Expected O, but got Unknown
		//IL_1845: Unknown result type (might be due to invalid IL or missing references)
		//IL_184f: Expected O, but got Unknown
		//IL_186d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1877: Expected O, but got Unknown
		//IL_1895: Unknown result type (might be due to invalid IL or missing references)
		//IL_189f: Expected O, but got Unknown
		//IL_18bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_18c7: Expected O, but got Unknown
		//IL_18e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_18ef: Expected O, but got Unknown
		//IL_190d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1917: Expected O, but got Unknown
		//IL_1935: Unknown result type (might be due to invalid IL or missing references)
		//IL_193f: Expected O, but got Unknown
		//IL_195d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1967: Expected O, but got Unknown
		//IL_1985: Unknown result type (might be due to invalid IL or missing references)
		//IL_198f: Expected O, but got Unknown
		//IL_19ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_19b7: Expected O, but got Unknown
		//IL_19d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_19df: Expected O, but got Unknown
		//IL_19fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1a07: Expected O, but got Unknown
		//IL_1a25: Unknown result type (might be due to invalid IL or missing references)
		//IL_1a2f: Expected O, but got Unknown
		//IL_1a4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1a57: Expected O, but got Unknown
		//IL_1a75: Unknown result type (might be due to invalid IL or missing references)
		//IL_1a7f: Expected O, but got Unknown
		//IL_1a9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1aa7: Expected O, but got Unknown
		//IL_1ac5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1acf: Expected O, but got Unknown
		//IL_1aed: Unknown result type (might be due to invalid IL or missing references)
		//IL_1af7: Expected O, but got Unknown
		//IL_1b15: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b1f: Expected O, but got Unknown
		//IL_1b3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b47: Expected O, but got Unknown
		//IL_1b65: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b6f: Expected O, but got Unknown
		//IL_1b8d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1b97: Expected O, but got Unknown
		//IL_1bb5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1bbf: Expected O, but got Unknown
		//IL_1bdd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1be7: Expected O, but got Unknown
		//IL_1c05: Unknown result type (might be due to invalid IL or missing references)
		//IL_1c0f: Expected O, but got Unknown
		//IL_1c2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1c37: Expected O, but got Unknown
		//IL_1c55: Unknown result type (might be due to invalid IL or missing references)
		//IL_1c5f: Expected O, but got Unknown
		//IL_1c7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1c87: Expected O, but got Unknown
		//IL_1ca5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1caf: Expected O, but got Unknown
		//IL_1ccd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1cd7: Expected O, but got Unknown
		//IL_1cf5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1cff: Expected O, but got Unknown
		//IL_1d1d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d27: Expected O, but got Unknown
		//IL_1d45: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d4f: Expected O, but got Unknown
		//IL_1d6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d77: Expected O, but got Unknown
		//IL_1d95: Unknown result type (might be due to invalid IL or missing references)
		//IL_1d9f: Expected O, but got Unknown
		//IL_1dbd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1dc7: Expected O, but got Unknown
		//IL_1de5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1def: Expected O, but got Unknown
		//IL_1e0d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e17: Expected O, but got Unknown
		//IL_1e35: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e3f: Expected O, but got Unknown
		//IL_1e5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e67: Expected O, but got Unknown
		//IL_1e85: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e8f: Expected O, but got Unknown
		//IL_1ead: Unknown result type (might be due to invalid IL or missing references)
		//IL_1eb7: Expected O, but got Unknown
		//IL_1ed5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1edf: Expected O, but got Unknown
		//IL_1efd: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f07: Expected O, but got Unknown
		//IL_1f25: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f2f: Expected O, but got Unknown
		//IL_1f4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f57: Expected O, but got Unknown
		//IL_1f75: Unknown result type (might be due to invalid IL or missing references)
		//IL_1f7f: Expected O, but got Unknown
		//IL_1f9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_1fa7: Expected O, but got Unknown
		//IL_1fc5: Unknown result type (might be due to invalid IL or missing references)
		//IL_1fcf: Expected O, but got Unknown
		//IL_1fed: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ff7: Expected O, but got Unknown
		//IL_2015: Unknown result type (might be due to invalid IL or missing references)
		//IL_201f: Expected O, but got Unknown
		//IL_203d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2047: Expected O, but got Unknown
		//IL_2065: Unknown result type (might be due to invalid IL or missing references)
		//IL_206f: Expected O, but got Unknown
		//IL_208d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2097: Expected O, but got Unknown
		//IL_20b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_20bf: Expected O, but got Unknown
		//IL_20dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_20e7: Expected O, but got Unknown
		//IL_2105: Unknown result type (might be due to invalid IL or missing references)
		//IL_210f: Expected O, but got Unknown
		//IL_212d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2137: Expected O, but got Unknown
		//IL_2155: Unknown result type (might be due to invalid IL or missing references)
		//IL_215f: Expected O, but got Unknown
		//IL_217d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2187: Expected O, but got Unknown
		//IL_21a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_21af: Expected O, but got Unknown
		//IL_21cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_21d7: Expected O, but got Unknown
		//IL_21f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_21ff: Expected O, but got Unknown
		//IL_221d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2227: Expected O, but got Unknown
		//IL_2245: Unknown result type (might be due to invalid IL or missing references)
		//IL_224f: Expected O, but got Unknown
		//IL_226d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2277: Expected O, but got Unknown
		//IL_2295: Unknown result type (might be due to invalid IL or missing references)
		//IL_229f: Expected O, but got Unknown
		//IL_22bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_22c7: Expected O, but got Unknown
		//IL_22e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_22ef: Expected O, but got Unknown
		//IL_230d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2317: Expected O, but got Unknown
		//IL_2335: Unknown result type (might be due to invalid IL or missing references)
		//IL_233f: Expected O, but got Unknown
		//IL_235d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2367: Expected O, but got Unknown
		//IL_2385: Unknown result type (might be due to invalid IL or missing references)
		//IL_238f: Expected O, but got Unknown
		//IL_23ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_23b7: Expected O, but got Unknown
		//IL_23d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_23df: Expected O, but got Unknown
		//IL_23fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2407: Expected O, but got Unknown
		//IL_2425: Unknown result type (might be due to invalid IL or missing references)
		//IL_242f: Expected O, but got Unknown
		//IL_244d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2457: Expected O, but got Unknown
		//IL_2475: Unknown result type (might be due to invalid IL or missing references)
		//IL_247f: Expected O, but got Unknown
		//IL_249d: Unknown result type (might be due to invalid IL or missing references)
		//IL_24a7: Expected O, but got Unknown
		//IL_24c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_24cf: Expected O, but got Unknown
		//IL_24ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_24f7: Expected O, but got Unknown
		//IL_2515: Unknown result type (might be due to invalid IL or missing references)
		//IL_251f: Expected O, but got Unknown
		//IL_253d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2547: Expected O, but got Unknown
		//IL_2565: Unknown result type (might be due to invalid IL or missing references)
		//IL_256f: Expected O, but got Unknown
		//IL_258d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2597: Expected O, but got Unknown
		//IL_25b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_25bf: Expected O, but got Unknown
		//IL_25dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_25e7: Expected O, but got Unknown
		//IL_2605: Unknown result type (might be due to invalid IL or missing references)
		//IL_260f: Expected O, but got Unknown
		//IL_262d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2637: Expected O, but got Unknown
		//IL_2655: Unknown result type (might be due to invalid IL or missing references)
		//IL_265f: Expected O, but got Unknown
		//IL_267d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2687: Expected O, but got Unknown
		//IL_26a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_26af: Expected O, but got Unknown
		//IL_26cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_26d7: Expected O, but got Unknown
		//IL_26f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_26ff: Expected O, but got Unknown
		//IL_271d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2727: Expected O, but got Unknown
		//IL_2745: Unknown result type (might be due to invalid IL or missing references)
		//IL_274f: Expected O, but got Unknown
		//IL_276d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2777: Expected O, but got Unknown
		//IL_2795: Unknown result type (might be due to invalid IL or missing references)
		//IL_279f: Expected O, but got Unknown
		//IL_27bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_27c7: Expected O, but got Unknown
		//IL_27e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_27ef: Expected O, but got Unknown
		//IL_280d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2817: Expected O, but got Unknown
		//IL_2835: Unknown result type (might be due to invalid IL or missing references)
		//IL_283f: Expected O, but got Unknown
		//IL_285d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2867: Expected O, but got Unknown
		//IL_2885: Unknown result type (might be due to invalid IL or missing references)
		//IL_288f: Expected O, but got Unknown
		//IL_28ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_28b7: Expected O, but got Unknown
		//IL_28d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_28df: Expected O, but got Unknown
		//IL_28fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2907: Expected O, but got Unknown
		//IL_2925: Unknown result type (might be due to invalid IL or missing references)
		//IL_292f: Expected O, but got Unknown
		//IL_294d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2957: Expected O, but got Unknown
		//IL_2975: Unknown result type (might be due to invalid IL or missing references)
		//IL_297f: Expected O, but got Unknown
		//IL_299d: Unknown result type (might be due to invalid IL or missing references)
		//IL_29a7: Expected O, but got Unknown
		//IL_29c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_29cf: Expected O, but got Unknown
		//IL_29ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_29f7: Expected O, but got Unknown
		//IL_2a15: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a1f: Expected O, but got Unknown
		//IL_2a3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a47: Expected O, but got Unknown
		//IL_2a65: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a6f: Expected O, but got Unknown
		//IL_2a8d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2a97: Expected O, but got Unknown
		//IL_2ab5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2abf: Expected O, but got Unknown
		//IL_2add: Unknown result type (might be due to invalid IL or missing references)
		//IL_2ae7: Expected O, but got Unknown
		//IL_2b05: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b0f: Expected O, but got Unknown
		//IL_2b2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b37: Expected O, but got Unknown
		//IL_2b55: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b5f: Expected O, but got Unknown
		//IL_2b7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2b87: Expected O, but got Unknown
		//IL_2ba5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2baf: Expected O, but got Unknown
		//IL_2bcd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2bd7: Expected O, but got Unknown
		//IL_2bf5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2bff: Expected O, but got Unknown
		//IL_2c1d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2c27: Expected O, but got Unknown
		//IL_2c45: Unknown result type (might be due to invalid IL or missing references)
		//IL_2c4f: Expected O, but got Unknown
		//IL_2c6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2c77: Expected O, but got Unknown
		//IL_2c95: Unknown result type (might be due to invalid IL or missing references)
		//IL_2c9f: Expected O, but got Unknown
		//IL_2cbd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2cc7: Expected O, but got Unknown
		//IL_2ce5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2cef: Expected O, but got Unknown
		//IL_2d0d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2d17: Expected O, but got Unknown
		//IL_2d35: Unknown result type (might be due to invalid IL or missing references)
		//IL_2d3f: Expected O, but got Unknown
		//IL_2d5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2d67: Expected O, but got Unknown
		//IL_2d85: Unknown result type (might be due to invalid IL or missing references)
		//IL_2d8f: Expected O, but got Unknown
		//IL_2dad: Unknown result type (might be due to invalid IL or missing references)
		//IL_2db7: Expected O, but got Unknown
		//IL_2dd5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2ddf: Expected O, but got Unknown
		//IL_2dfd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e07: Expected O, but got Unknown
		//IL_2e25: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e2f: Expected O, but got Unknown
		//IL_2e4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e57: Expected O, but got Unknown
		//IL_2e75: Unknown result type (might be due to invalid IL or missing references)
		//IL_2e7f: Expected O, but got Unknown
		//IL_2e9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2ea7: Expected O, but got Unknown
		//IL_2ec5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2ecf: Expected O, but got Unknown
		//IL_2eed: Unknown result type (might be due to invalid IL or missing references)
		//IL_2ef7: Expected O, but got Unknown
		//IL_2f15: Unknown result type (might be due to invalid IL or missing references)
		//IL_2f1f: Expected O, but got Unknown
		//IL_2f3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2f47: Expected O, but got Unknown
		//IL_2f65: Unknown result type (might be due to invalid IL or missing references)
		//IL_2f6f: Expected O, but got Unknown
		//IL_2f8d: Unknown result type (might be due to invalid IL or missing references)
		//IL_2f97: Expected O, but got Unknown
		//IL_2fb5: Unknown result type (might be due to invalid IL or missing references)
		//IL_2fbf: Expected O, but got Unknown
		//IL_2fdd: Unknown result type (might be due to invalid IL or missing references)
		//IL_2fe7: Expected O, but got Unknown
		//IL_3005: Unknown result type (might be due to invalid IL or missing references)
		//IL_300f: Expected O, but got Unknown
		//IL_302d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3037: Expected O, but got Unknown
		//IL_3055: Unknown result type (might be due to invalid IL or missing references)
		//IL_305f: Expected O, but got Unknown
		//IL_307d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3087: Expected O, but got Unknown
		//IL_30a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_30af: Expected O, but got Unknown
		//IL_30cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_30d7: Expected O, but got Unknown
		//IL_30f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_30ff: Expected O, but got Unknown
		//IL_311d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3127: Expected O, but got Unknown
		//IL_3145: Unknown result type (might be due to invalid IL or missing references)
		//IL_314f: Expected O, but got Unknown
		//IL_316d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3177: Expected O, but got Unknown
		//IL_3195: Unknown result type (might be due to invalid IL or missing references)
		//IL_319f: Expected O, but got Unknown
		//IL_31bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_31c7: Expected O, but got Unknown
		//IL_31e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_31ef: Expected O, but got Unknown
		//IL_320d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3217: Expected O, but got Unknown
		//IL_3235: Unknown result type (might be due to invalid IL or missing references)
		//IL_323f: Expected O, but got Unknown
		//IL_325d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3267: Expected O, but got Unknown
		//IL_3285: Unknown result type (might be due to invalid IL or missing references)
		//IL_328f: Expected O, but got Unknown
		//IL_32ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_32b7: Expected O, but got Unknown
		//IL_32d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_32df: Expected O, but got Unknown
		//IL_32fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3307: Expected O, but got Unknown
		//IL_3325: Unknown result type (might be due to invalid IL or missing references)
		//IL_332f: Expected O, but got Unknown
		//IL_334d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3357: Expected O, but got Unknown
		//IL_3375: Unknown result type (might be due to invalid IL or missing references)
		//IL_337f: Expected O, but got Unknown
		//IL_339d: Unknown result type (might be due to invalid IL or missing references)
		//IL_33a7: Expected O, but got Unknown
		//IL_33c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_33cf: Expected O, but got Unknown
		//IL_33ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_33f7: Expected O, but got Unknown
		//IL_3415: Unknown result type (might be due to invalid IL or missing references)
		//IL_341f: Expected O, but got Unknown
		//IL_343d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3447: Expected O, but got Unknown
		//IL_3465: Unknown result type (might be due to invalid IL or missing references)
		//IL_346f: Expected O, but got Unknown
		//IL_348d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3497: Expected O, but got Unknown
		//IL_34b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_34bf: Expected O, but got Unknown
		//IL_34dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_34e7: Expected O, but got Unknown
		//IL_3505: Unknown result type (might be due to invalid IL or missing references)
		//IL_350f: Expected O, but got Unknown
		//IL_352d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3537: Expected O, but got Unknown
		//IL_3555: Unknown result type (might be due to invalid IL or missing references)
		//IL_355f: Expected O, but got Unknown
		//IL_357d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3587: Expected O, but got Unknown
		//IL_35a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_35af: Expected O, but got Unknown
		//IL_35cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_35d7: Expected O, but got Unknown
		//IL_35f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_35ff: Expected O, but got Unknown
		//IL_361d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3627: Expected O, but got Unknown
		//IL_3645: Unknown result type (might be due to invalid IL or missing references)
		//IL_364f: Expected O, but got Unknown
		//IL_366d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3677: Expected O, but got Unknown
		//IL_3695: Unknown result type (might be due to invalid IL or missing references)
		//IL_369f: Expected O, but got Unknown
		//IL_36bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_36c7: Expected O, but got Unknown
		//IL_36e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_36ef: Expected O, but got Unknown
		//IL_370d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3717: Expected O, but got Unknown
		//IL_3735: Unknown result type (might be due to invalid IL or missing references)
		//IL_373f: Expected O, but got Unknown
		//IL_375d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3767: Expected O, but got Unknown
		//IL_3785: Unknown result type (might be due to invalid IL or missing references)
		//IL_378f: Expected O, but got Unknown
		//IL_37ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_37b7: Expected O, but got Unknown
		//IL_37d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_37df: Expected O, but got Unknown
		//IL_37fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3807: Expected O, but got Unknown
		//IL_3825: Unknown result type (might be due to invalid IL or missing references)
		//IL_382f: Expected O, but got Unknown
		//IL_384d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3857: Expected O, but got Unknown
		//IL_3875: Unknown result type (might be due to invalid IL or missing references)
		//IL_387f: Expected O, but got Unknown
		//IL_389d: Unknown result type (might be due to invalid IL or missing references)
		//IL_38a7: Expected O, but got Unknown
		//IL_38c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_38cf: Expected O, but got Unknown
		//IL_38ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_38f7: Expected O, but got Unknown
		//IL_3915: Unknown result type (might be due to invalid IL or missing references)
		//IL_391f: Expected O, but got Unknown
		//IL_393d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3947: Expected O, but got Unknown
		//IL_3965: Unknown result type (might be due to invalid IL or missing references)
		//IL_396f: Expected O, but got Unknown
		//IL_398d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3997: Expected O, but got Unknown
		//IL_39b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_39bf: Expected O, but got Unknown
		//IL_39dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_39e7: Expected O, but got Unknown
		//IL_3a05: Unknown result type (might be due to invalid IL or missing references)
		//IL_3a0f: Expected O, but got Unknown
		//IL_3a2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3a37: Expected O, but got Unknown
		//IL_3a55: Unknown result type (might be due to invalid IL or missing references)
		//IL_3a5f: Expected O, but got Unknown
		//IL_3a7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3a87: Expected O, but got Unknown
		//IL_3aa5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3aaf: Expected O, but got Unknown
		//IL_3acd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3ad7: Expected O, but got Unknown
		//IL_3af5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3aff: Expected O, but got Unknown
		//IL_3b1d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3b27: Expected O, but got Unknown
		//IL_3b45: Unknown result type (might be due to invalid IL or missing references)
		//IL_3b4f: Expected O, but got Unknown
		//IL_3b6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3b77: Expected O, but got Unknown
		//IL_3b95: Unknown result type (might be due to invalid IL or missing references)
		//IL_3b9f: Expected O, but got Unknown
		//IL_3bbd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3bc7: Expected O, but got Unknown
		//IL_3be5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3bef: Expected O, but got Unknown
		//IL_3c0d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c17: Expected O, but got Unknown
		//IL_3c35: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c3f: Expected O, but got Unknown
		//IL_3c5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c67: Expected O, but got Unknown
		//IL_3c85: Unknown result type (might be due to invalid IL or missing references)
		//IL_3c8f: Expected O, but got Unknown
		//IL_3cad: Unknown result type (might be due to invalid IL or missing references)
		//IL_3cb7: Expected O, but got Unknown
		//IL_3cd5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3cdf: Expected O, but got Unknown
		//IL_3cfd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3d07: Expected O, but got Unknown
		//IL_3d25: Unknown result type (might be due to invalid IL or missing references)
		//IL_3d2f: Expected O, but got Unknown
		//IL_3d4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3d57: Expected O, but got Unknown
		//IL_3d75: Unknown result type (might be due to invalid IL or missing references)
		//IL_3d7f: Expected O, but got Unknown
		//IL_3d9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3da7: Expected O, but got Unknown
		//IL_3dc5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3dcf: Expected O, but got Unknown
		//IL_3ded: Unknown result type (might be due to invalid IL or missing references)
		//IL_3df7: Expected O, but got Unknown
		//IL_3e15: Unknown result type (might be due to invalid IL or missing references)
		//IL_3e1f: Expected O, but got Unknown
		//IL_3e3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3e47: Expected O, but got Unknown
		//IL_3e65: Unknown result type (might be due to invalid IL or missing references)
		//IL_3e6f: Expected O, but got Unknown
		//IL_3e8d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3e97: Expected O, but got Unknown
		//IL_3eb5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3ebf: Expected O, but got Unknown
		//IL_3edd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3ee7: Expected O, but got Unknown
		//IL_3f05: Unknown result type (might be due to invalid IL or missing references)
		//IL_3f0f: Expected O, but got Unknown
		//IL_3f2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3f37: Expected O, but got Unknown
		//IL_3f55: Unknown result type (might be due to invalid IL or missing references)
		//IL_3f5f: Expected O, but got Unknown
		//IL_3f7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_3f87: Expected O, but got Unknown
		//IL_3fa5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3faf: Expected O, but got Unknown
		//IL_3fcd: Unknown result type (might be due to invalid IL or missing references)
		//IL_3fd7: Expected O, but got Unknown
		//IL_3ff5: Unknown result type (might be due to invalid IL or missing references)
		//IL_3fff: Expected O, but got Unknown
		//IL_401d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4027: Expected O, but got Unknown
		//IL_4045: Unknown result type (might be due to invalid IL or missing references)
		//IL_404f: Expected O, but got Unknown
		//IL_406d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4077: Expected O, but got Unknown
		//IL_4095: Unknown result type (might be due to invalid IL or missing references)
		//IL_409f: Expected O, but got Unknown
		//IL_40bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_40c7: Expected O, but got Unknown
		//IL_40e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_40ef: Expected O, but got Unknown
		//IL_410d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4117: Expected O, but got Unknown
		//IL_4135: Unknown result type (might be due to invalid IL or missing references)
		//IL_413f: Expected O, but got Unknown
		//IL_415d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4167: Expected O, but got Unknown
		//IL_4185: Unknown result type (might be due to invalid IL or missing references)
		//IL_418f: Expected O, but got Unknown
		//IL_41ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_41b7: Expected O, but got Unknown
		//IL_41d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_41df: Expected O, but got Unknown
		//IL_41fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4207: Expected O, but got Unknown
		//IL_4225: Unknown result type (might be due to invalid IL or missing references)
		//IL_422f: Expected O, but got Unknown
		//IL_424d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4257: Expected O, but got Unknown
		//IL_4275: Unknown result type (might be due to invalid IL or missing references)
		//IL_427f: Expected O, but got Unknown
		//IL_429d: Unknown result type (might be due to invalid IL or missing references)
		//IL_42a7: Expected O, but got Unknown
		//IL_42c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_42cf: Expected O, but got Unknown
		//IL_42ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_42f7: Expected O, but got Unknown
		//IL_4315: Unknown result type (might be due to invalid IL or missing references)
		//IL_431f: Expected O, but got Unknown
		//IL_433d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4347: Expected O, but got Unknown
		//IL_4365: Unknown result type (might be due to invalid IL or missing references)
		//IL_436f: Expected O, but got Unknown
		//IL_438d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4397: Expected O, but got Unknown
		//IL_43b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_43bf: Expected O, but got Unknown
		//IL_43dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_43e7: Expected O, but got Unknown
		//IL_4405: Unknown result type (might be due to invalid IL or missing references)
		//IL_440f: Expected O, but got Unknown
		//IL_442d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4437: Expected O, but got Unknown
		//IL_4455: Unknown result type (might be due to invalid IL or missing references)
		//IL_445f: Expected O, but got Unknown
		//IL_447d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4487: Expected O, but got Unknown
		//IL_44a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_44af: Expected O, but got Unknown
		//IL_44cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_44d7: Expected O, but got Unknown
		//IL_44f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_44ff: Expected O, but got Unknown
		//IL_451d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4527: Expected O, but got Unknown
		//IL_4545: Unknown result type (might be due to invalid IL or missing references)
		//IL_454f: Expected O, but got Unknown
		//IL_456d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4577: Expected O, but got Unknown
		//IL_4595: Unknown result type (might be due to invalid IL or missing references)
		//IL_459f: Expected O, but got Unknown
		//IL_45bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_45c7: Expected O, but got Unknown
		//IL_45e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_45ef: Expected O, but got Unknown
		//IL_460d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4617: Expected O, but got Unknown
		//IL_4635: Unknown result type (might be due to invalid IL or missing references)
		//IL_463f: Expected O, but got Unknown
		//IL_465d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4667: Expected O, but got Unknown
		//IL_4685: Unknown result type (might be due to invalid IL or missing references)
		//IL_468f: Expected O, but got Unknown
		//IL_46ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_46b7: Expected O, but got Unknown
		//IL_46d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_46df: Expected O, but got Unknown
		//IL_46fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4707: Expected O, but got Unknown
		//IL_4725: Unknown result type (might be due to invalid IL or missing references)
		//IL_472f: Expected O, but got Unknown
		//IL_474d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4757: Expected O, but got Unknown
		//IL_4775: Unknown result type (might be due to invalid IL or missing references)
		//IL_477f: Expected O, but got Unknown
		//IL_479d: Unknown result type (might be due to invalid IL or missing references)
		//IL_47a7: Expected O, but got Unknown
		//IL_47c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_47cf: Expected O, but got Unknown
		//IL_47ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_47f7: Expected O, but got Unknown
		//IL_4815: Unknown result type (might be due to invalid IL or missing references)
		//IL_481f: Expected O, but got Unknown
		//IL_483d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4847: Expected O, but got Unknown
		//IL_4865: Unknown result type (might be due to invalid IL or missing references)
		//IL_486f: Expected O, but got Unknown
		//IL_488d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4897: Expected O, but got Unknown
		//IL_48b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_48bf: Expected O, but got Unknown
		//IL_48dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_48e7: Expected O, but got Unknown
		//IL_4905: Unknown result type (might be due to invalid IL or missing references)
		//IL_490f: Expected O, but got Unknown
		//IL_492d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4937: Expected O, but got Unknown
		//IL_4955: Unknown result type (might be due to invalid IL or missing references)
		//IL_495f: Expected O, but got Unknown
		//IL_497d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4987: Expected O, but got Unknown
		//IL_49a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_49af: Expected O, but got Unknown
		//IL_49cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_49d7: Expected O, but got Unknown
		//IL_49f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_49ff: Expected O, but got Unknown
		//IL_4a1d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4a27: Expected O, but got Unknown
		//IL_4a45: Unknown result type (might be due to invalid IL or missing references)
		//IL_4a4f: Expected O, but got Unknown
		//IL_4a6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4a77: Expected O, but got Unknown
		//IL_4a95: Unknown result type (might be due to invalid IL or missing references)
		//IL_4a9f: Expected O, but got Unknown
		//IL_4abd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4ac7: Expected O, but got Unknown
		//IL_4ae5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4aef: Expected O, but got Unknown
		//IL_4b0d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4b17: Expected O, but got Unknown
		//IL_4b35: Unknown result type (might be due to invalid IL or missing references)
		//IL_4b3f: Expected O, but got Unknown
		//IL_4b5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4b67: Expected O, but got Unknown
		//IL_4b85: Unknown result type (might be due to invalid IL or missing references)
		//IL_4b8f: Expected O, but got Unknown
		//IL_4bad: Unknown result type (might be due to invalid IL or missing references)
		//IL_4bb7: Expected O, but got Unknown
		//IL_4bd5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4bdf: Expected O, but got Unknown
		//IL_4bfd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4c07: Expected O, but got Unknown
		//IL_4c25: Unknown result type (might be due to invalid IL or missing references)
		//IL_4c2f: Expected O, but got Unknown
		//IL_4c4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4c57: Expected O, but got Unknown
		//IL_4c75: Unknown result type (might be due to invalid IL or missing references)
		//IL_4c7f: Expected O, but got Unknown
		//IL_4c9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4ca7: Expected O, but got Unknown
		//IL_4cc5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4ccf: Expected O, but got Unknown
		//IL_4ced: Unknown result type (might be due to invalid IL or missing references)
		//IL_4cf7: Expected O, but got Unknown
		//IL_4d15: Unknown result type (might be due to invalid IL or missing references)
		//IL_4d1f: Expected O, but got Unknown
		//IL_4d3d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4d47: Expected O, but got Unknown
		//IL_4d65: Unknown result type (might be due to invalid IL or missing references)
		//IL_4d6f: Expected O, but got Unknown
		//IL_4d8d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4d97: Expected O, but got Unknown
		//IL_4db5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4dbf: Expected O, but got Unknown
		//IL_4ddd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4de7: Expected O, but got Unknown
		//IL_4e05: Unknown result type (might be due to invalid IL or missing references)
		//IL_4e0f: Expected O, but got Unknown
		//IL_4e2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4e37: Expected O, but got Unknown
		//IL_4e55: Unknown result type (might be due to invalid IL or missing references)
		//IL_4e5f: Expected O, but got Unknown
		//IL_4e7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4e87: Expected O, but got Unknown
		//IL_4ea5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4eaf: Expected O, but got Unknown
		//IL_4ecd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4ed7: Expected O, but got Unknown
		//IL_4ef5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4eff: Expected O, but got Unknown
		//IL_4f1d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4f27: Expected O, but got Unknown
		//IL_4f45: Unknown result type (might be due to invalid IL or missing references)
		//IL_4f4f: Expected O, but got Unknown
		//IL_4f6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_4f77: Expected O, but got Unknown
		//IL_4f95: Unknown result type (might be due to invalid IL or missing references)
		//IL_4f9f: Expected O, but got Unknown
		//IL_4fbd: Unknown result type (might be due to invalid IL or missing references)
		//IL_4fc7: Expected O, but got Unknown
		//IL_4fe5: Unknown result type (might be due to invalid IL or missing references)
		//IL_4fef: Expected O, but got Unknown
		//IL_500d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5017: Expected O, but got Unknown
		//IL_5035: Unknown result type (might be due to invalid IL or missing references)
		//IL_503f: Expected O, but got Unknown
		//IL_505d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5067: Expected O, but got Unknown
		//IL_5085: Unknown result type (might be due to invalid IL or missing references)
		//IL_508f: Expected O, but got Unknown
		//IL_50ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_50b7: Expected O, but got Unknown
		//IL_50d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_50df: Expected O, but got Unknown
		//IL_50fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_5107: Expected O, but got Unknown
		//IL_5125: Unknown result type (might be due to invalid IL or missing references)
		//IL_512f: Expected O, but got Unknown
		//IL_514d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5157: Expected O, but got Unknown
		//IL_5175: Unknown result type (might be due to invalid IL or missing references)
		//IL_517f: Expected O, but got Unknown
		//IL_519d: Unknown result type (might be due to invalid IL or missing references)
		//IL_51a7: Expected O, but got Unknown
		//IL_51c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_51cf: Expected O, but got Unknown
		//IL_51ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_51f7: Expected O, but got Unknown
		//IL_5215: Unknown result type (might be due to invalid IL or missing references)
		//IL_521f: Expected O, but got Unknown
		//IL_523d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5247: Expected O, but got Unknown
		//IL_5265: Unknown result type (might be due to invalid IL or missing references)
		//IL_526f: Expected O, but got Unknown
		//IL_528d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5297: Expected O, but got Unknown
		//IL_52b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_52bf: Expected O, but got Unknown
		//IL_52dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_52e7: Expected O, but got Unknown
		//IL_5305: Unknown result type (might be due to invalid IL or missing references)
		//IL_530f: Expected O, but got Unknown
		//IL_532d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5337: Expected O, but got Unknown
		//IL_5355: Unknown result type (might be due to invalid IL or missing references)
		//IL_535f: Expected O, but got Unknown
		//IL_537d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5387: Expected O, but got Unknown
		//IL_53a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_53af: Expected O, but got Unknown
		//IL_53cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_53d7: Expected O, but got Unknown
		//IL_53f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_53ff: Expected O, but got Unknown
		//IL_541d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5427: Expected O, but got Unknown
		//IL_5445: Unknown result type (might be due to invalid IL or missing references)
		//IL_544f: Expected O, but got Unknown
		//IL_546d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5477: Expected O, but got Unknown
		//IL_5495: Unknown result type (might be due to invalid IL or missing references)
		//IL_549f: Expected O, but got Unknown
		//IL_54bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_54c7: Expected O, but got Unknown
		//IL_54e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_54ef: Expected O, but got Unknown
		//IL_550d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5517: Expected O, but got Unknown
		//IL_5535: Unknown result type (might be due to invalid IL or missing references)
		//IL_553f: Expected O, but got Unknown
		//IL_555d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5567: Expected O, but got Unknown
		//IL_5585: Unknown result type (might be due to invalid IL or missing references)
		//IL_558f: Expected O, but got Unknown
		//IL_55ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_55b7: Expected O, but got Unknown
		//IL_55d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_55df: Expected O, but got Unknown
		//IL_55fd: Unknown result type (might be due to invalid IL or missing references)
		//IL_5607: Expected O, but got Unknown
		//IL_5625: Unknown result type (might be due to invalid IL or missing references)
		//IL_562f: Expected O, but got Unknown
		//IL_564d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5657: Expected O, but got Unknown
		//IL_5675: Unknown result type (might be due to invalid IL or missing references)
		//IL_567f: Expected O, but got Unknown
		//IL_56a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_56aa: Expected O, but got Unknown
		//IL_56cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_56d5: Expected O, but got Unknown
		//IL_56f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_5700: Expected O, but got Unknown
		//IL_5721: Unknown result type (might be due to invalid IL or missing references)
		//IL_572b: Expected O, but got Unknown
		//IL_574c: Unknown result type (might be due to invalid IL or missing references)
		//IL_5756: Expected O, but got Unknown
		//IL_5777: Unknown result type (might be due to invalid IL or missing references)
		//IL_5781: Expected O, but got Unknown
		//IL_57a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_57ac: Expected O, but got Unknown
		//IL_57cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_57d7: Expected O, but got Unknown
		//IL_57f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_5802: Expected O, but got Unknown
		//IL_5823: Unknown result type (might be due to invalid IL or missing references)
		//IL_582d: Expected O, but got Unknown
		//IL_584e: Unknown result type (might be due to invalid IL or missing references)
		//IL_5858: Expected O, but got Unknown
		//IL_5879: Unknown result type (might be due to invalid IL or missing references)
		//IL_5883: Expected O, but got Unknown
		//IL_58a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_58ae: Expected O, but got Unknown
		//IL_58cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_58d9: Expected O, but got Unknown
		//IL_58fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_5904: Expected O, but got Unknown
		//IL_5925: Unknown result type (might be due to invalid IL or missing references)
		//IL_592f: Expected O, but got Unknown
		//IL_5950: Unknown result type (might be due to invalid IL or missing references)
		//IL_595a: Expected O, but got Unknown
		//IL_597b: Unknown result type (might be due to invalid IL or missing references)
		//IL_5985: Expected O, but got Unknown
		//IL_59a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_59b0: Expected O, but got Unknown
		//IL_59d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_59db: Expected O, but got Unknown
		//IL_59fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_5a06: Expected O, but got Unknown
		//IL_5a27: Unknown result type (might be due to invalid IL or missing references)
		//IL_5a31: Expected O, but got Unknown
		//IL_5a52: Unknown result type (might be due to invalid IL or missing references)
		//IL_5a5c: Expected O, but got Unknown
		//IL_5a7d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5a87: Expected O, but got Unknown
		//IL_5aa8: Unknown result type (might be due to invalid IL or missing references)
		//IL_5ab2: Expected O, but got Unknown
		//IL_5ad3: Unknown result type (might be due to invalid IL or missing references)
		//IL_5add: Expected O, but got Unknown
		//IL_5afe: Unknown result type (might be due to invalid IL or missing references)
		//IL_5b08: Expected O, but got Unknown
		//IL_5b29: Unknown result type (might be due to invalid IL or missing references)
		//IL_5b33: Expected O, but got Unknown
		//IL_5b54: Unknown result type (might be due to invalid IL or missing references)
		//IL_5b5e: Expected O, but got Unknown
		//IL_5b7f: Unknown result type (might be due to invalid IL or missing references)
		//IL_5b89: Expected O, but got Unknown
		//IL_5baa: Unknown result type (might be due to invalid IL or missing references)
		//IL_5bb4: Expected O, but got Unknown
		//IL_5bd5: Unknown result type (might be due to invalid IL or missing references)
		//IL_5bdf: Expected O, but got Unknown
		//IL_5c00: Unknown result type (might be due to invalid IL or missing references)
		//IL_5c0a: Expected O, but got Unknown
		//IL_5c2b: Unknown result type (might be due to invalid IL or missing references)
		//IL_5c35: Expected O, but got Unknown
		//IL_5c56: Unknown result type (might be due to invalid IL or missing references)
		//IL_5c60: Expected O, but got Unknown
		//IL_5c81: Unknown result type (might be due to invalid IL or missing references)
		//IL_5c8b: Expected O, but got Unknown
		//IL_5cac: Unknown result type (might be due to invalid IL or missing references)
		//IL_5cb6: Expected O, but got Unknown
		//IL_5cd7: Unknown result type (might be due to invalid IL or missing references)
		//IL_5ce1: Expected O, but got Unknown
		//IL_5d02: Unknown result type (might be due to invalid IL or missing references)
		//IL_5d0c: Expected O, but got Unknown
		//IL_5d2d: Unknown result type (might be due to invalid IL or missing references)
		//IL_5d37: Expected O, but got Unknown
		//IL_5d58: Unknown result type (might be due to invalid IL or missing references)
		//IL_5d62: Expected O, but got Unknown
		//IL_5d83: Unknown result type (might be due to invalid IL or missing references)
		//IL_5d8d: Expected O, but got Unknown
		//IL_5dae: Unknown result type (might be due to invalid IL or missing references)
		//IL_5db8: Expected O, but got Unknown
		//IL_5dd9: Unknown result type (might be due to invalid IL or missing references)
		//IL_5de3: Expected O, but got Unknown
		//IL_5e04: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e0e: Expected O, but got Unknown
		//IL_5e2f: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e39: Expected O, but got Unknown
		//IL_5e5a: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e64: Expected O, but got Unknown
		//IL_5e85: Unknown result type (might be due to invalid IL or missing references)
		//IL_5e8f: Expected O, but got Unknown
		//IL_5eb0: Unknown result type (might be due to invalid IL or missing references)
		//IL_5eba: Expected O, but got Unknown
		//IL_5edb: Unknown result type (might be due to invalid IL or missing references)
		//IL_5ee5: Expected O, but got Unknown
		//IL_5f06: Unknown result type (might be due to invalid IL or missing references)
		//IL_5f10: Expected O, but got Unknown
		//IL_5f31: Unknown result type (might be due to invalid IL or missing references)
		//IL_5f3b: Expected O, but got Unknown
		//IL_5f5c: Unknown result type (might be due to invalid IL or missing references)
		//IL_5f66: Expected O, but got Unknown
		//IL_5f87: Unknown result type (might be due to invalid IL or missing references)
		//IL_5f91: Expected O, but got Unknown
		//IL_5fb2: Unknown result type (might be due to invalid IL or missing references)
		//IL_5fbc: Expected O, but got Unknown
		//IL_5fdd: Unknown result type (might be due to invalid IL or missing references)
		//IL_5fe7: Expected O, but got Unknown
		//IL_6008: Unknown result type (might be due to invalid IL or missing references)
		//IL_6012: Expected O, but got Unknown
		//IL_6033: Unknown result type (might be due to invalid IL or missing references)
		//IL_603d: Expected O, but got Unknown
		//IL_605e: Unknown result type (might be due to invalid IL or missing references)
		//IL_6068: Expected O, but got Unknown
		//IL_6089: Unknown result type (might be due to invalid IL or missing references)
		//IL_6093: Expected O, but got Unknown
		//IL_60b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_60be: Expected O, but got Unknown
		//IL_60df: Unknown result type (might be due to invalid IL or missing references)
		//IL_60e9: Expected O, but got Unknown
		//IL_610a: Unknown result type (might be due to invalid IL or missing references)
		//IL_6114: Expected O, but got Unknown
		//IL_6135: Unknown result type (might be due to invalid IL or missing references)
		//IL_613f: Expected O, but got Unknown
		//IL_6160: Unknown result type (might be due to invalid IL or missing references)
		//IL_616a: Expected O, but got Unknown
		//IL_618b: Unknown result type (might be due to invalid IL or missing references)
		//IL_6195: Expected O, but got Unknown
		//IL_61b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_61c0: Expected O, but got Unknown
		//IL_61e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_61eb: Expected O, but got Unknown
		//IL_620c: Unknown result type (might be due to invalid IL or missing references)
		//IL_6216: Expected O, but got Unknown
		//IL_6237: Unknown result type (might be due to invalid IL or missing references)
		//IL_6241: Expected O, but got Unknown
		//IL_6262: Unknown result type (might be due to invalid IL or missing references)
		//IL_626c: Expected O, but got Unknown
		//IL_628d: Unknown result type (might be due to invalid IL or missing references)
		//IL_6297: Expected O, but got Unknown
		//IL_62b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_62c2: Expected O, but got Unknown
		//IL_62e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_62ed: Expected O, but got Unknown
		//IL_630e: Unknown result type (might be due to invalid IL or missing references)
		//IL_6318: Expected O, but got Unknown
		//IL_6339: Unknown result type (might be due to invalid IL or missing references)
		//IL_6343: Expected O, but got Unknown
		//IL_6364: Unknown result type (might be due to invalid IL or missing references)
		//IL_636e: Expected O, but got Unknown
		//IL_638f: Unknown result type (might be due to invalid IL or missing references)
		//IL_6399: Expected O, but got Unknown
		//IL_63ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_63c4: Expected O, but got Unknown
		//IL_63e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_63ef: Expected O, but got Unknown
		//IL_6410: Unknown result type (might be due to invalid IL or missing references)
		//IL_641a: Expected O, but got Unknown
		//IL_643b: Unknown result type (might be due to invalid IL or missing references)
		//IL_6445: Expected O, but got Unknown
		//IL_6466: Unknown result type (might be due to invalid IL or missing references)
		//IL_6470: Expected O, but got Unknown
		//IL_6491: Unknown result type (might be due to invalid IL or missing references)
		//IL_649b: Expected O, but got Unknown
		//IL_64bc: Unknown result type (might be due to invalid IL or missing references)
		//IL_64c6: Expected O, but got Unknown
		//IL_64e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_64f1: Expected O, but got Unknown
		//IL_6512: Unknown result type (might be due to invalid IL or missing references)
		//IL_651c: Expected O, but got Unknown
		//IL_653d: Unknown result type (might be due to invalid IL or missing references)
		//IL_6547: Expected O, but got Unknown
		//IL_6568: Unknown result type (might be due to invalid IL or missing references)
		//IL_6572: Expected O, but got Unknown
		//IL_6593: Unknown result type (might be due to invalid IL or missing references)
		//IL_659d: Expected O, but got Unknown
		//IL_65be: Unknown result type (might be due to invalid IL or missing references)
		//IL_65c8: Expected O, but got Unknown
		//IL_65e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_65f3: Expected O, but got Unknown
		//IL_6614: Unknown result type (might be due to invalid IL or missing references)
		//IL_661e: Expected O, but got Unknown
		//IL_663f: Unknown result type (might be due to invalid IL or missing references)
		//IL_6649: Expected O, but got Unknown
		//IL_666a: Unknown result type (might be due to invalid IL or missing references)
		//IL_6674: Expected O, but got Unknown
		//IL_6695: Unknown result type (might be due to invalid IL or missing references)
		//IL_669f: Expected O, but got Unknown
		//IL_66c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_66ca: Expected O, but got Unknown
		//IL_66eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_66f5: Expected O, but got Unknown
		//IL_6716: Unknown result type (might be due to invalid IL or missing references)
		//IL_6720: Expected O, but got Unknown
		//IL_6741: Unknown result type (might be due to invalid IL or missing references)
		//IL_674b: Expected O, but got Unknown
		//IL_676c: Unknown result type (might be due to invalid IL or missing references)
		//IL_6776: Expected O, but got Unknown
		//IL_6797: Unknown result type (might be due to invalid IL or missing references)
		//IL_67a1: Expected O, but got Unknown
		//IL_67c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_67cc: Expected O, but got Unknown
		//IL_67ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_67f7: Expected O, but got Unknown
		//IL_6818: Unknown result type (might be due to invalid IL or missing references)
		//IL_6822: Expected O, but got Unknown
		//IL_6843: Unknown result type (might be due to invalid IL or missing references)
		//IL_684d: Expected O, but got Unknown
		//IL_686e: Unknown result type (might be due to invalid IL or missing references)
		//IL_6878: Expected O, but got Unknown
		//IL_6899: Unknown result type (might be due to invalid IL or missing references)
		//IL_68a3: Expected O, but got Unknown
		//IL_68c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_68ce: Expected O, but got Unknown
		//IL_68ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_68f9: Expected O, but got Unknown
		//IL_691a: Unknown result type (might be due to invalid IL or missing references)
		//IL_6924: Expected O, but got Unknown
		//IL_6945: Unknown result type (might be due to invalid IL or missing references)
		//IL_694f: Expected O, but got Unknown
		//IL_6970: Unknown result type (might be due to invalid IL or missing references)
		//IL_697a: Expected O, but got Unknown
		//IL_699b: Unknown result type (might be due to invalid IL or missing references)
		//IL_69a5: Expected O, but got Unknown
		//IL_69c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_69d0: Expected O, but got Unknown
		//IL_69f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_69fb: Expected O, but got Unknown
		//IL_6a1c: Unknown result type (might be due to invalid IL or missing references)
		//IL_6a26: Expected O, but got Unknown
		//IL_6a47: Unknown result type (might be due to invalid IL or missing references)
		//IL_6a51: Expected O, but got Unknown
		//IL_6a72: Unknown result type (might be due to invalid IL or missing references)
		//IL_6a7c: Expected O, but got Unknown
		//IL_6a9d: Unknown result type (might be due to invalid IL or missing references)
		//IL_6aa7: Expected O, but got Unknown
		//IL_6ac8: Unknown result type (might be due to invalid IL or missing references)
		//IL_6ad2: Expected O, but got Unknown
		//IL_6af3: Unknown result type (might be due to invalid IL or missing references)
		//IL_6afd: Expected O, but got Unknown
		//IL_6b1e: Unknown result type (might be due to invalid IL or missing references)
		//IL_6b28: Expected O, but got Unknown
		//IL_6b49: Unknown result type (might be due to invalid IL or missing references)
		//IL_6b53: Expected O, but got Unknown
		//IL_6b74: Unknown result type (might be due to invalid IL or missing references)
		//IL_6b7e: Expected O, but got Unknown
		//IL_6b9f: Unknown result type (might be due to invalid IL or missing references)
		//IL_6ba9: Expected O, but got Unknown
		//IL_6bca: Unknown result type (might be due to invalid IL or missing references)
		//IL_6bd4: Expected O, but got Unknown
		//IL_6bf5: Unknown result type (might be due to invalid IL or missing references)
		//IL_6bff: Expected O, but got Unknown
		//IL_6c20: Unknown result type (might be due to invalid IL or missing references)
		//IL_6c2a: Expected O, but got Unknown
		//IL_6c4b: Unknown result type (might be due to invalid IL or missing references)
		//IL_6c55: Expected O, but got Unknown
		//IL_6c76: Unknown result type (might be due to invalid IL or missing references)
		//IL_6c80: Expected O, but got Unknown
		//IL_6ca1: Unknown result type (might be due to invalid IL or missing references)
		//IL_6cab: Expected O, but got Unknown
		//IL_6ccc: Unknown result type (might be due to invalid IL or missing references)
		//IL_6cd6: Expected O, but got Unknown
		//IL_6cf7: Unknown result type (might be due to invalid IL or missing references)
		//IL_6d01: Expected O, but got Unknown
		//IL_6d22: Unknown result type (might be due to invalid IL or missing references)
		//IL_6d2c: Expected O, but got Unknown
		//IL_6d4d: Unknown result type (might be due to invalid IL or missing references)
		//IL_6d57: Expected O, but got Unknown
		//IL_6d78: Unknown result type (might be due to invalid IL or missing references)
		//IL_6d82: Expected O, but got Unknown
		//IL_6da3: Unknown result type (might be due to invalid IL or missing references)
		//IL_6dad: Expected O, but got Unknown
		//IL_6dce: Unknown result type (might be due to invalid IL or missing references)
		//IL_6dd8: Expected O, but got Unknown
		//IL_6df9: Unknown result type (might be due to invalid IL or missing references)
		//IL_6e03: Expected O, but got Unknown
		//IL_6e24: Unknown result type (might be due to invalid IL or missing references)
		//IL_6e2e: Expected O, but got Unknown
		//IL_6e4f: Unknown result type (might be due to invalid IL or missing references)
		//IL_6e59: Expected O, but got Unknown
		//IL_6e7a: Unknown result type (might be due to invalid IL or missing references)
		//IL_6e84: Expected O, but got Unknown
		//IL_6ea5: Unknown result type (might be due to invalid IL or missing references)
		//IL_6eaf: Expected O, but got Unknown
		//IL_6ed0: Unknown result type (might be due to invalid IL or missing references)
		//IL_6eda: Expected O, but got Unknown
		//IL_6efb: Unknown result type (might be due to invalid IL or missing references)
		//IL_6f05: Expected O, but got Unknown
		//IL_6f26: Unknown result type (might be due to invalid IL or missing references)
		//IL_6f30: Expected O, but got Unknown
		//IL_6f51: Unknown result type (might be due to invalid IL or missing references)
		//IL_6f5b: Expected O, but got Unknown
		//IL_6f7c: Unknown result type (might be due to invalid IL or missing references)
		//IL_6f86: Expected O, but got Unknown
		//IL_6fa7: Unknown result type (might be due to invalid IL or missing references)
		//IL_6fb1: Expected O, but got Unknown
		//IL_6fd2: Unknown result type (might be due to invalid IL or missing references)
		//IL_6fdc: Expected O, but got Unknown
		//IL_6ffd: Unknown result type (might be due to invalid IL or missing references)
		//IL_7007: Expected O, but got Unknown
		//IL_7028: Unknown result type (might be due to invalid IL or missing references)
		//IL_7032: Expected O, but got Unknown
		//IL_7053: Unknown result type (might be due to invalid IL or missing references)
		//IL_705d: Expected O, but got Unknown
		//IL_707e: Unknown result type (might be due to invalid IL or missing references)
		//IL_7088: Expected O, but got Unknown
		//IL_70a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_70b3: Expected O, but got Unknown
		//IL_70d4: Unknown result type (might be due to invalid IL or missing references)
		//IL_70de: Expected O, but got Unknown
		//IL_70ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_7109: Expected O, but got Unknown
		//IL_712a: Unknown result type (might be due to invalid IL or missing references)
		//IL_7134: Expected O, but got Unknown
		//IL_7155: Unknown result type (might be due to invalid IL or missing references)
		//IL_715f: Expected O, but got Unknown
		//IL_7180: Unknown result type (might be due to invalid IL or missing references)
		//IL_718a: Expected O, but got Unknown
		//IL_71ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_71b5: Expected O, but got Unknown
		//IL_71d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_71e0: Expected O, but got Unknown
		//IL_7201: Unknown result type (might be due to invalid IL or missing references)
		//IL_720b: Expected O, but got Unknown
		//IL_722c: Unknown result type (might be due to invalid IL or missing references)
		//IL_7236: Expected O, but got Unknown
		//IL_7257: Unknown result type (might be due to invalid IL or missing references)
		//IL_7261: Expected O, but got Unknown
		//IL_7282: Unknown result type (might be due to invalid IL or missing references)
		//IL_728c: Expected O, but got Unknown
		//IL_72ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_72b7: Expected O, but got Unknown
		//IL_72d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_72e2: Expected O, but got Unknown
		//IL_7303: Unknown result type (might be due to invalid IL or missing references)
		//IL_730d: Expected O, but got Unknown
		//IL_732e: Unknown result type (might be due to invalid IL or missing references)
		//IL_7338: Expected O, but got Unknown
		//IL_7359: Unknown result type (might be due to invalid IL or missing references)
		//IL_7363: Expected O, but got Unknown
		//IL_7384: Unknown result type (might be due to invalid IL or missing references)
		//IL_738e: Expected O, but got Unknown
		//IL_73af: Unknown result type (might be due to invalid IL or missing references)
		//IL_73b9: Expected O, but got Unknown
		//IL_73da: Unknown result type (might be due to invalid IL or missing references)
		//IL_73e4: Expected O, but got Unknown
		//IL_7405: Unknown result type (might be due to invalid IL or missing references)
		//IL_740f: Expected O, but got Unknown
		//IL_7430: Unknown result type (might be due to invalid IL or missing references)
		//IL_743a: Expected O, but got Unknown
		//IL_745b: Unknown result type (might be due to invalid IL or missing references)
		//IL_7465: Expected O, but got Unknown
		//IL_7486: Unknown result type (might be due to invalid IL or missing references)
		//IL_7490: Expected O, but got Unknown
		//IL_74b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_74bb: Expected O, but got Unknown
		//IL_74dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_74e6: Expected O, but got Unknown
		//IL_7507: Unknown result type (might be due to invalid IL or missing references)
		//IL_7511: Expected O, but got Unknown
		//IL_7532: Unknown result type (might be due to invalid IL or missing references)
		//IL_753c: Expected O, but got Unknown
		//IL_755d: Unknown result type (might be due to invalid IL or missing references)
		//IL_7567: Expected O, but got Unknown
		//IL_7588: Unknown result type (might be due to invalid IL or missing references)
		//IL_7592: Expected O, but got Unknown
		//IL_75b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_75bd: Expected O, but got Unknown
		//IL_75de: Unknown result type (might be due to invalid IL or missing references)
		//IL_75e8: Expected O, but got Unknown
		//IL_7609: Unknown result type (might be due to invalid IL or missing references)
		//IL_7613: Expected O, but got Unknown
		//IL_7634: Unknown result type (might be due to invalid IL or missing references)
		//IL_763e: Expected O, but got Unknown
		//IL_765f: Unknown result type (might be due to invalid IL or missing references)
		//IL_7669: Expected O, but got Unknown
		//IL_768a: Unknown result type (might be due to invalid IL or missing references)
		//IL_7694: Expected O, but got Unknown
		//IL_76b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_76bf: Expected O, but got Unknown
		//IL_76e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_76ea: Expected O, but got Unknown
		//IL_770b: Unknown result type (might be due to invalid IL or missing references)
		//IL_7715: Expected O, but got Unknown
		//IL_7736: Unknown result type (might be due to invalid IL or missing references)
		//IL_7740: Expected O, but got Unknown
		//IL_7761: Unknown result type (might be due to invalid IL or missing references)
		//IL_776b: Expected O, but got Unknown
		//IL_778c: Unknown result type (might be due to invalid IL or missing references)
		//IL_7796: Expected O, but got Unknown
		//IL_77b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_77c1: Expected O, but got Unknown
		//IL_77e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_77ec: Expected O, but got Unknown
		//IL_780d: Unknown result type (might be due to invalid IL or missing references)
		//IL_7817: Expected O, but got Unknown
		//IL_7838: Unknown result type (might be due to invalid IL or missing references)
		//IL_7842: Expected O, but got Unknown
		//IL_7863: Unknown result type (might be due to invalid IL or missing references)
		//IL_786d: Expected O, but got Unknown
		//IL_788e: Unknown result type (might be due to invalid IL or missing references)
		//IL_7898: Expected O, but got Unknown
		//IL_78b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_78c3: Expected O, but got Unknown
		//IL_78e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_78ee: Expected O, but got Unknown
		//IL_790f: Unknown result type (might be due to invalid IL or missing references)
		//IL_7919: Expected O, but got Unknown
		//IL_793a: Unknown result type (might be due to invalid IL or missing references)
		//IL_7944: Expected O, but got Unknown
		//IL_7965: Unknown result type (might be due to invalid IL or missing references)
		//IL_796f: Expected O, but got Unknown
		//IL_7990: Unknown result type (might be due to invalid IL or missing references)
		//IL_799a: Expected O, but got Unknown
		//IL_79bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_79c5: Expected O, but got Unknown
		//IL_79e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_79f0: Expected O, but got Unknown
		//IL_7a11: Unknown result type (might be due to invalid IL or missing references)
		//IL_7a1b: Expected O, but got Unknown
		//IL_7a3c: Unknown result type (might be due to invalid IL or missing references)
		//IL_7a46: Expected O, but got Unknown
		//IL_7a67: Unknown result type (might be due to invalid IL or missing references)
		//IL_7a71: Expected O, but got Unknown
		//IL_7a92: Unknown result type (might be due to invalid IL or missing references)
		//IL_7a9c: Expected O, but got Unknown
		//IL_7abd: Unknown result type (might be due to invalid IL or missing references)
		//IL_7ac7: Expected O, but got Unknown
		//IL_7ae8: Unknown result type (might be due to invalid IL or missing references)
		//IL_7af2: Expected O, but got Unknown
		//IL_7b13: Unknown result type (might be due to invalid IL or missing references)
		//IL_7b1d: Expected O, but got Unknown
		//IL_7b3e: Unknown result type (might be due to invalid IL or missing references)
		//IL_7b48: Expected O, but got Unknown
		//IL_7b69: Unknown result type (might be due to invalid IL or missing references)
		//IL_7b73: Expected O, but got Unknown
		//IL_7b94: Unknown result type (might be due to invalid IL or missing references)
		//IL_7b9e: Expected O, but got Unknown
		//IL_7bbf: Unknown result type (might be due to invalid IL or missing references)
		//IL_7bc9: Expected O, but got Unknown
		//IL_7bea: Unknown result type (might be due to invalid IL or missing references)
		//IL_7bf4: Expected O, but got Unknown
		//IL_7c15: Unknown result type (might be due to invalid IL or missing references)
		//IL_7c1f: Expected O, but got Unknown
		//IL_7c40: Unknown result type (might be due to invalid IL or missing references)
		//IL_7c4a: Expected O, but got Unknown
		//IL_7c6b: Unknown result type (might be due to invalid IL or missing references)
		//IL_7c75: Expected O, but got Unknown
		//IL_7c96: Unknown result type (might be due to invalid IL or missing references)
		//IL_7ca0: Expected O, but got Unknown
		//IL_7cc1: Unknown result type (might be due to invalid IL or missing references)
		//IL_7ccb: Expected O, but got Unknown
		//IL_7cec: Unknown result type (might be due to invalid IL or missing references)
		//IL_7cf6: Expected O, but got Unknown
		//IL_7d17: Unknown result type (might be due to invalid IL or missing references)
		//IL_7d21: Expected O, but got Unknown
		//IL_7d42: Unknown result type (might be due to invalid IL or missing references)
		//IL_7d4c: Expected O, but got Unknown
		//IL_7d6d: Unknown result type (might be due to invalid IL or missing references)
		//IL_7d77: Expected O, but got Unknown
		//IL_7d98: Unknown result type (might be due to invalid IL or missing references)
		//IL_7da2: Expected O, but got Unknown
		//IL_7dc3: Unknown result type (might be due to invalid IL or missing references)
		//IL_7dcd: Expected O, but got Unknown
		//IL_7dee: Unknown result type (might be due to invalid IL or missing references)
		//IL_7df8: Expected O, but got Unknown
		//IL_7e19: Unknown result type (might be due to invalid IL or missing references)
		//IL_7e23: Expected O, but got Unknown
		//IL_7e44: Unknown result type (might be due to invalid IL or missing references)
		//IL_7e4e: Expected O, but got Unknown
		//IL_7e6f: Unknown result type (might be due to invalid IL or missing references)
		//IL_7e79: Expected O, but got Unknown
		//IL_7e9a: Unknown result type (might be due to invalid IL or missing references)
		//IL_7ea4: Expected O, but got Unknown
		//IL_7ec5: Unknown result type (might be due to invalid IL or missing references)
		//IL_7ecf: Expected O, but got Unknown
		//IL_7ef0: Unknown result type (might be due to invalid IL or missing references)
		//IL_7efa: Expected O, but got Unknown
		//IL_7f1b: Unknown result type (might be due to invalid IL or missing references)
		//IL_7f25: Expected O, but got Unknown
		//IL_7f46: Unknown result type (might be due to invalid IL or missing references)
		//IL_7f50: Expected O, but got Unknown
		//IL_7f71: Unknown result type (might be due to invalid IL or missing references)
		//IL_7f7b: Expected O, but got Unknown
		//IL_7f9c: Unknown result type (might be due to invalid IL or missing references)
		//IL_7fa6: Expected O, but got Unknown
		//IL_7fc7: Unknown result type (might be due to invalid IL or missing references)
		//IL_7fd1: Expected O, but got Unknown
		//IL_7ff2: Unknown result type (might be due to invalid IL or missing references)
		//IL_7ffc: Expected O, but got Unknown
		//IL_801d: Unknown result type (might be due to invalid IL or missing references)
		//IL_8027: Expected O, but got Unknown
		//IL_8048: Unknown result type (might be due to invalid IL or missing references)
		//IL_8052: Expected O, but got Unknown
		//IL_8073: Unknown result type (might be due to invalid IL or missing references)
		//IL_807d: Expected O, but got Unknown
		//IL_809e: Unknown result type (might be due to invalid IL or missing references)
		//IL_80a8: Expected O, but got Unknown
		//IL_80c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_80d3: Expected O, but got Unknown
		//IL_80f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_80fe: Expected O, but got Unknown
		//IL_811f: Unknown result type (might be due to invalid IL or missing references)
		//IL_8129: Expected O, but got Unknown
		//IL_814a: Unknown result type (might be due to invalid IL or missing references)
		//IL_8154: Expected O, but got Unknown
		//IL_8175: Unknown result type (might be due to invalid IL or missing references)
		//IL_817f: Expected O, but got Unknown
		//IL_81a0: Unknown result type (might be due to invalid IL or missing references)
		//IL_81aa: Expected O, but got Unknown
		//IL_81cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_81d5: Expected O, but got Unknown
		//IL_81f6: Unknown result type (might be due to invalid IL or missing references)
		//IL_8200: Expected O, but got Unknown
		//IL_8221: Unknown result type (might be due to invalid IL or missing references)
		//IL_822b: Expected O, but got Unknown
		//IL_824c: Unknown result type (might be due to invalid IL or missing references)
		//IL_8256: Expected O, but got Unknown
		//IL_8277: Unknown result type (might be due to invalid IL or missing references)
		//IL_8281: Expected O, but got Unknown
		//IL_82a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_82ac: Expected O, but got Unknown
		//IL_82cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_82d7: Expected O, but got Unknown
		//IL_82f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_8302: Expected O, but got Unknown
		//IL_8323: Unknown result type (might be due to invalid IL or missing references)
		//IL_832d: Expected O, but got Unknown
		//IL_834e: Unknown result type (might be due to invalid IL or missing references)
		//IL_8358: Expected O, but got Unknown
		//IL_8379: Unknown result type (might be due to invalid IL or missing references)
		//IL_8383: Expected O, but got Unknown
		//IL_83a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_83ae: Expected O, but got Unknown
		//IL_83cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_83d9: Expected O, but got Unknown
		//IL_83fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_8404: Expected O, but got Unknown
		//IL_8425: Unknown result type (might be due to invalid IL or missing references)
		//IL_842f: Expected O, but got Unknown
		//IL_8450: Unknown result type (might be due to invalid IL or missing references)
		//IL_845a: Expected O, but got Unknown
		//IL_847b: Unknown result type (might be due to invalid IL or missing references)
		//IL_8485: Expected O, but got Unknown
		//IL_84a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_84b0: Expected O, but got Unknown
		//IL_84d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_84db: Expected O, but got Unknown
		//IL_84fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_8506: Expected O, but got Unknown
		//IL_8527: Unknown result type (might be due to invalid IL or missing references)
		//IL_8531: Expected O, but got Unknown
		//IL_8552: Unknown result type (might be due to invalid IL or missing references)
		//IL_855c: Expected O, but got Unknown
		//IL_857d: Unknown result type (might be due to invalid IL or missing references)
		//IL_8587: Expected O, but got Unknown
		//IL_85a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_85b2: Expected O, but got Unknown
		//IL_85d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_85dd: Expected O, but got Unknown
		//IL_85fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_8608: Expected O, but got Unknown
		//IL_8629: Unknown result type (might be due to invalid IL or missing references)
		//IL_8633: Expected O, but got Unknown
		//IL_8b0f: Unknown result type (might be due to invalid IL or missing references)
		//IL_8b91: Unknown result type (might be due to invalid IL or missing references)
		//IL_8c1c: Unknown result type (might be due to invalid IL or missing references)
		//IL_8c26: Expected O, but got Unknown
		//IL_8c5d: Unknown result type (might be due to invalid IL or missing references)
		//IL_8cda: Unknown result type (might be due to invalid IL or missing references)
		//IL_8ce4: Expected O, but got Unknown
		//IL_8d71: Unknown result type (might be due to invalid IL or missing references)
		//IL_8d7b: Expected O, but got Unknown
		//IL_8d7e: Unknown result type (might be due to invalid IL or missing references)
		//IL_8dcf: Unknown result type (might be due to invalid IL or missing references)
		//IL_8dd9: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmHTKKSystems));
		columnHeader2 = new ColumnHeader();
		columnHeader5 = new ColumnHeader();
		columnHeaderName = new ColumnHeader();
		columnHeader4 = new ColumnHeader();
		columnHeader3 = new ColumnHeader();
		columnHeader1 = new ColumnHeader();
		statusStrip = new StatusStrip();
		toolStripStatusMST = new ToolStripStatusLabel();
		toolStripStatusUserName = new ToolStripStatusLabel();
		toolStripStatusLabel5 = new ToolStripStatusLabel();
		toolStripStatusTimeLogin = new ToolStripStatusLabel();
		txtVersion = new ToolStripStatusLabel();
		statusStrip1 = new StatusStrip();
		toolStripStatusfooter = new ToolStripStatusLabel();
		winSpMenu = new FpSpread();
		winSpMenu_Sheet1 = new SheetView();
		panelMenu = new Panel();
		panelPeriod = new Panel();
		lblTitle_Version_bg = new Label();
		((Control)statusStrip).SuspendLayout();
		((Control)statusStrip1).SuspendLayout();
		((ISupportInitialize)winSpMenu).BeginInit();
		((ISupportInitialize)winSpMenu_Sheet1).BeginInit();
		((Control)panelMenu).SuspendLayout();
		((Control)this).SuspendLayout();
		columnHeader2.Text = "STT";
		columnHeader2.TextAlign = (HorizontalAlignment)2;
		columnHeader2.Width = 40;
		columnHeader5.Text = "Trạng thái";
		columnHeader5.TextAlign = (HorizontalAlignment)2;
		columnHeader5.Width = 100;
		columnHeaderName.Text = "Tên đăng nhập";
		columnHeaderName.Width = 150;
		columnHeader4.Text = "Email";
		columnHeader4.Width = 150;
		columnHeader3.Text = "Điện thoại";
		columnHeader3.Width = 150;
		columnHeader1.Text = "Tên đầy đủ";
		columnHeader1.Width = 250;
		((Control)statusStrip).AutoSize = false;
		((ToolStrip)statusStrip).BackColor = Color.Firebrick;
		((Control)statusStrip).Dock = (DockStyle)1;
		((ToolStrip)statusStrip).GripMargin = new Padding(2, 0, 2, 0);
		((ToolStrip)statusStrip).ImageScalingSize = new Size(20, 20);
		((ToolStrip)statusStrip).Items.AddRange((ToolStripItem[])(object)new ToolStripItem[5]
		{
			(ToolStripItem)toolStripStatusMST,
			(ToolStripItem)toolStripStatusUserName,
			(ToolStripItem)toolStripStatusLabel5,
			(ToolStripItem)toolStripStatusTimeLogin,
			(ToolStripItem)txtVersion
		});
		((Control)statusStrip).Location = new Point(0, 0);
		((Control)statusStrip).Name = "statusStrip";
		statusStrip.Padding = new Padding(1, 0, 19, 0);
		((Control)statusStrip).Size = new Size(1434, 28);
		((Control)statusStrip).TabIndex = 27;
		((Control)statusStrip).Text = "statusStrip1";
		((ToolStripItem)toolStripStatusMST).Name = "toolStripStatusMST";
		((ToolStripItem)toolStripStatusMST).Size = new Size(0, 23);
		((ToolStripItem)toolStripStatusUserName).Name = "toolStripStatusUserName";
		((ToolStripItem)toolStripStatusUserName).Size = new Size(0, 23);
		((ToolStripItem)toolStripStatusLabel5).Name = "toolStripStatusLabel5";
		((ToolStripItem)toolStripStatusLabel5).Size = new Size(0, 23);
		((ToolStripItem)toolStripStatusTimeLogin).Name = "toolStripStatusTimeLogin";
		((ToolStripItem)toolStripStatusTimeLogin).Size = new Size(0, 23);
		((ToolStripItem)txtVersion).Name = "txtVersion";
		((ToolStripItem)txtVersion).Size = new Size(0, 23);
		((Control)statusStrip1).AutoSize = false;
		((ToolStrip)statusStrip1).BackColor = Color.Firebrick;
		((ToolStrip)statusStrip1).GripMargin = new Padding(2, 0, 2, 0);
		((ToolStrip)statusStrip1).ImageScalingSize = new Size(20, 20);
		((ToolStrip)statusStrip1).Items.AddRange((ToolStripItem[])(object)new ToolStripItem[1] { (ToolStripItem)toolStripStatusfooter });
		((Control)statusStrip1).Location = new Point(0, 670);
		((Control)statusStrip1).Name = "statusStrip1";
		statusStrip1.Padding = new Padding(1, 0, 19, 0);
		((Control)statusStrip1).Size = new Size(1434, 91);
		((Control)statusStrip1).TabIndex = 34;
		((Control)statusStrip1).Text = "statusStrip1";
		((Control)statusStrip1).Resize += statusStrip1_Resize;
		((ToolStripItem)toolStripStatusfooter).Font = new Font("Segoe UI", 9f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((ToolStripItem)toolStripStatusfooter).ForeColor = Color.White;
		((ToolStripItem)toolStripStatusfooter).Margin = new Padding(10, 3, 0, 2);
		((ToolStripItem)toolStripStatusfooter).Name = "toolStripStatusfooter";
		((ToolStripItem)toolStripStatusfooter).Size = new Size(650, 69);
		((ToolStripItem)toolStripStatusfooter).Text = "Bản quyền thuộc Cục Thuế - Bộ Tài chính\nĐịa chỉ: Số 1A Nguyễn Công Trứ, Phường Hai Bà Trưng, Thành phố Hà Nội; Website: http://www.gdt.gov.vn\nPhần mềm cung cấp miễn phí cho người nộp thuế";
		((Control)winSpMenu).AccessibleDescription = "winSpMenu, Sheet1, Row 0, Column 0, ";
		((Control)winSpMenu).BackgroundImageLayout = (ImageLayout)3;
		winSpMenu.BorderStyle = (BorderStyle)0;
		((Control)winSpMenu).Dock = (DockStyle)5;
		((Control)winSpMenu).Location = new Point(0, 0);
		((Control)winSpMenu).Margin = new Padding(4);
		winSpMenu.Name = "winSpMenu";
		winSpMenu.Sheets.AddRange(new SheetView[1] { winSpMenu_Sheet1 });
		((Control)winSpMenu).Size = new Size(616, 642);
		((Control)winSpMenu).TabIndex = 35;
		winSpMenu.CellClick += winSpMenu_CellClick;
		((Control)winSpMenu).MouseDown += new MouseEventHandler(winSpMenu_MouseDown);
		((Control)winSpMenu).MouseHover += winSpMenu_MouseHover;
		((Control)winSpMenu).MouseMove += new MouseEventHandler(winSpMenu_MouseMove);
		winSpMenu.SetViewportTopRow(0, 0, 216);
		winSpMenu_Sheet1.Reset();
		winSpMenu_Sheet1.SheetName = "Sheet1";
		winSpMenu_Sheet1.ReferenceStyle = ReferenceStyle.R1C1;
		winSpMenu_Sheet1.ColumnCount = 12;
		winSpMenu_Sheet1.Cells.Get(0, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(0, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(0, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(0, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(1, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(1, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(1, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(1, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(2, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(2, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(2, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(2, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(3, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(3, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(3, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(3, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(4, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(4, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(4, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(4, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(5, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(5, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(5, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(5, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(6, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(6, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(6, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(6, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(7, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(7, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(7, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(7, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(8, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(8, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(8, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(8, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(9, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(9, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(9, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(9, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(10, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(10, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(10, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(10, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(11, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(11, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(11, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(11, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(12, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(12, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(12, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(12, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(13, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(13, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(13, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(13, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(14, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(14, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(14, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(14, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(15, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(15, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(15, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(15, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(16, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(16, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(16, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(16, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(17, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(17, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(17, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(17, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(18, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(18, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(18, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(18, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(19, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(19, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(19, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(19, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(20, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(20, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(20, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(20, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(21, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(21, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(21, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(21, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(22, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(22, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(22, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(22, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(23, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(23, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(23, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(23, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(24, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(24, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(24, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(24, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(25, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(25, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(25, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(25, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(26, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(26, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(26, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(26, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(27, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(27, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(27, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(27, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(28, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(28, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(28, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(28, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(29, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(29, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(29, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(29, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(30, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(30, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(30, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(30, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(31, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(31, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(31, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(31, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(32, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(32, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(32, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(32, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(33, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(33, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(33, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(33, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(34, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(34, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(34, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(34, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(35, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(35, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(35, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(35, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(36, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(36, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(36, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(36, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(37, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(37, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(37, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(37, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(38, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(38, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(38, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(38, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(39, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(39, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(39, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(39, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(40, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(40, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(40, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(40, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(41, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(41, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(41, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(41, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(42, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(42, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(42, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(42, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(43, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(43, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(43, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(43, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(44, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(44, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(44, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(44, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(45, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(45, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(45, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(45, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(46, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(46, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(46, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(46, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(47, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(47, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(47, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(47, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(48, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(48, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(48, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(48, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(49, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(49, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(49, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(49, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(50, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(50, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(50, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(50, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(51, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(51, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(51, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(51, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(52, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(52, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(52, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(52, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(53, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(53, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(53, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(53, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(54, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(54, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(54, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(54, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(55, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(55, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(55, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(55, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(56, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(56, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(56, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(56, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(57, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(57, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(57, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(57, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(58, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(58, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(58, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(58, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(59, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(59, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(59, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(59, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(60, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(60, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(60, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(60, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(61, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(61, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(61, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(61, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(62, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(62, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(62, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(62, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(63, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(63, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(63, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(63, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(64, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(64, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(64, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(64, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(65, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(65, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(65, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(65, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(66, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(66, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(66, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(66, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(67, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(67, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(67, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(67, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(68, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(68, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(68, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(68, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(69, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(69, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(69, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(69, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(70, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(70, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(70, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(70, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(71, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(71, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(71, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(71, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(72, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(72, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(72, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(72, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(73, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(73, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(73, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(73, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(74, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(74, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(74, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(74, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(75, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(75, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(75, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(75, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(76, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(76, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(76, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(76, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(77, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(77, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(77, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(77, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(78, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(78, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(78, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(78, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(79, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(79, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(79, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(79, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(80, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(80, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(80, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(80, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(81, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(81, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(81, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(81, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(82, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(82, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(82, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(82, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(83, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(83, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(83, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(83, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(84, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(84, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(84, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(84, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(85, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(85, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(85, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(85, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(86, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(86, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(86, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(86, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(87, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(87, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(87, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(87, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(88, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(88, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(88, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(88, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(89, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(89, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(89, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(89, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(90, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(90, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(90, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(90, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(91, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(91, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(91, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(91, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(92, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(92, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(92, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(92, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(93, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(93, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(93, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(93, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(94, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(94, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(94, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(94, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(95, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(95, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(95, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(95, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(96, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(96, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(96, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(96, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(97, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(97, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(97, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(97, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(98, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(98, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(98, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(98, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(99, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(99, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(99, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(99, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(100, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(100, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(100, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(100, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(101, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(101, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(101, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(101, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(102, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(102, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(102, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(102, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(103, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(103, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(103, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(103, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(104, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(104, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(104, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(104, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(105, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(105, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(105, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(105, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(106, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(106, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(106, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(106, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(107, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(107, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(107, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(107, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(108, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(108, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(108, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(108, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(109, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(109, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(109, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(109, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(110, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(110, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(110, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(110, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(111, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(111, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(111, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(111, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(112, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(112, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(112, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(112, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(113, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(113, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(113, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(113, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(114, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(114, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(114, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(114, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(115, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(115, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(115, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(115, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(116, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(116, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(116, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(116, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(117, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(117, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(117, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(117, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(118, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(118, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(118, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(118, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(119, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(119, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(119, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(119, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(120, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(120, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(120, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(120, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(121, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(121, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(121, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(121, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(122, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(122, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(122, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(122, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(123, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(123, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(123, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(123, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(124, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(124, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(124, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(124, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(125, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(125, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(125, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(125, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(126, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(126, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(126, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(126, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(127, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(127, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(127, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(127, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(128, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(128, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(128, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(128, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(129, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(129, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(129, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(129, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(130, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(130, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(130, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(130, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(131, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(131, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(131, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(131, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(132, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(132, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(132, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(132, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(133, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(133, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(133, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(133, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(134, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(134, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(134, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(134, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(135, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(135, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(135, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(135, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(136, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(136, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(136, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(136, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(137, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(137, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(137, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(137, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(138, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(138, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(138, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(138, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(139, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(139, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(139, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(139, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(140, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(140, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(140, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(140, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(141, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(141, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(141, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(141, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(142, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(142, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(142, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(142, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(143, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(143, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(143, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(143, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(144, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(144, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(144, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(144, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(145, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(145, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(145, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(145, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(146, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(146, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(146, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(146, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(147, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(147, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(147, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(147, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(148, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(148, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(148, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(148, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(149, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(149, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(149, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(149, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(150, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(150, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(150, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(150, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(151, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(151, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(151, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(151, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(152, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(152, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(152, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(152, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(153, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(153, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(153, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(153, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(154, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(154, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(154, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(154, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(155, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(155, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(155, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(155, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(156, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(156, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(156, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(156, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(157, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(157, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(157, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(157, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(158, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(158, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(158, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(158, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(159, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(159, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(159, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(159, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(160, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(160, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(160, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(160, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(161, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(161, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(161, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(161, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(162, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(162, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(162, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(162, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(163, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(163, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(163, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(163, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(164, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(164, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(164, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(164, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(165, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(165, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(165, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(165, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(166, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(166, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(166, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(166, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(167, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(167, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(167, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(167, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(168, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(168, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(168, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(168, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(169, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(169, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(169, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(169, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(170, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(170, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(170, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(170, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(171, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(171, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(171, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(171, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(172, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(172, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(172, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(172, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(173, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(173, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(173, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(173, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(174, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(174, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(174, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(174, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(175, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(175, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(175, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(175, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(176, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(176, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(176, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(176, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(177, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(177, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(177, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(177, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(178, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(178, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(178, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(178, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(179, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(179, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(179, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(179, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(180, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(180, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(180, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(180, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(181, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(181, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(181, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(181, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(182, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(182, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(182, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(182, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(183, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(183, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(183, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(183, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(184, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(184, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(184, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(184, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(185, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(185, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(185, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(185, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(186, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(186, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(186, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(186, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(187, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(187, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(187, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(187, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(188, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(188, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(188, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(188, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(189, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(189, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(189, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(189, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(190, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(190, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(190, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(190, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(191, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(191, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(191, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(191, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(192, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(192, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(192, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(192, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(193, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(193, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(193, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(193, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(194, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(194, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(194, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(194, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(195, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(195, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(195, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(195, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(196, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(196, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(196, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(196, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(197, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(197, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(197, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(197, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(198, 1).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(198, 2).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(198, 3).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.Cells.Get(198, 4).Font = new Font("Tahoma", 8f);
		winSpMenu_Sheet1.ColumnFooter.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		winSpMenu_Sheet1.ColumnFooter.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerFooterDefaultEnhanced";
		winSpMenu_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		winSpMenu_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.ColumnHeader.Rows.Get(0).Height = 19f;
		winSpMenu_Sheet1.Columns.Get(0).Width = 11f;
		winSpMenu_Sheet1.Columns.Get(1).Width = 18f;
		winSpMenu_Sheet1.Columns.Get(2).Width = 18f;
		winSpMenu_Sheet1.Columns.Get(3).Width = 18f;
		winSpMenu_Sheet1.Columns.Get(4).Width = 351f;
		winSpMenu_Sheet1.Columns.Get(5).Visible = false;
		winSpMenu_Sheet1.Columns.Get(6).Visible = false;
		winSpMenu_Sheet1.Columns.Get(7).Visible = false;
		winSpMenu_Sheet1.Columns.Get(8).Visible = false;
		winSpMenu_Sheet1.Columns.Get(9).Visible = false;
		winSpMenu_Sheet1.Columns.Get(10).Visible = false;
		winSpMenu_Sheet1.Columns.Get(10).Width = 54f;
		winSpMenu_Sheet1.Columns.Get(11).Visible = false;
		winSpMenu_Sheet1.Columns.Get(11).Width = 72f;
		winSpMenu_Sheet1.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.DefaultStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.DefaultStyle.Parent = "DataAreaDefault";
		winSpMenu_Sheet1.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefaultEnhanced";
		winSpMenu_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.FilterBarHeaderStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.FilterBarHeaderStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.FilterBarHeaderStyle.Parent = "RowHeaderDefaultEnhanced";
		winSpMenu_Sheet1.FilterBarHeaderStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.RowHeader.Columns.Default.Resizable = false;
		winSpMenu_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefaultEnhanced";
		winSpMenu_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.SheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winSpMenu_Sheet1.SheetCornerStyle.NoteIndicatorColor = Color.Red;
		winSpMenu_Sheet1.SheetCornerStyle.Parent = "CornerDefaultEnhanced";
		winSpMenu_Sheet1.SheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		winSpMenu_Sheet1.ReferenceStyle = ReferenceStyle.A1;
		((Control)panelMenu).BackColor = Color.Transparent;
		((Control)panelMenu).Controls.Add((Control)(object)winSpMenu);
		((Control)panelMenu).Dock = (DockStyle)3;
		((Control)panelMenu).Location = new Point(0, 28);
		((Control)panelMenu).Margin = new Padding(4);
		((Control)panelMenu).Name = "panelMenu";
		((Control)panelMenu).Size = new Size(616, 642);
		((Control)panelMenu).TabIndex = 36;
		((Control)panelPeriod).Anchor = (AnchorStyles)15;
		panelPeriod.BorderStyle = (BorderStyle)1;
		((Control)panelPeriod).Location = new Point(859, 734);
		((Control)panelPeriod).Margin = new Padding(4);
		((Control)panelPeriod).Name = "panelPeriod";
		((Control)panelPeriod).Size = new Size(13, 12);
		((Control)panelPeriod).TabIndex = 37;
		((Control)panelPeriod).Visible = false;
		((Control)lblTitle_Version_bg).Anchor = (AnchorStyles)10;
		((Control)lblTitle_Version_bg).AutoSize = true;
		((Control)lblTitle_Version_bg).BackColor = Color.Transparent;
		((Control)lblTitle_Version_bg).Font = new Font("Microsoft Sans Serif", 10f, (FontStyle)1, (GraphicsUnit)3, (byte)0);
		((Control)lblTitle_Version_bg).ForeColor = Color.DarkRed;
		((Control)lblTitle_Version_bg).Location = new Point(1268, 639);
		((Control)lblTitle_Version_bg).Margin = new Padding(4, 0, 4, 0);
		((Control)lblTitle_Version_bg).Name = "lblTitle_Version_bg";
		((Control)lblTitle_Version_bg).Size = new Size(144, 20);
		((Control)lblTitle_Version_bg).TabIndex = 38;
		((Control)lblTitle_Version_bg).Text = "Phiên bản: 4.0.0";
		((ContainerControl)this).AutoScaleDimensions = new SizeF(8f, 16f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackgroundImage = (Image)componentResourceManager.GetObject("$this.BackgroundImage");
		((Control)this).BackgroundImageLayout = (ImageLayout)3;
		((Form)this).ClientSize = new Size(1434, 761);
		((Control)this).Controls.Add((Control)(object)lblTitle_Version_bg);
		((Control)this).Controls.Add((Control)(object)panelPeriod);
		((Control)this).Controls.Add((Control)(object)panelMenu);
		((Control)this).Controls.Add((Control)(object)statusStrip1);
		((Control)this).Controls.Add((Control)(object)statusStrip);
		((Control)this).DoubleBuffered = true;
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).Margin = new Padding(4);
		((Control)this).MinimumSize = new Size(1061, 727);
		((Control)this).Name = "frmHTKKSystems";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "frmHTKKSystems";
		((Form)this).WindowState = (FormWindowState)2;
		((Form)this).FormClosed += new FormClosedEventHandler(frmHTKKSystems_FormClosed);
		((Form)this).Load += frmHTKKSystems_Load;
		((Control)this).Resize += frmHTKKSystems_Resize;
		((Control)statusStrip).ResumeLayout(false);
		((Control)statusStrip).PerformLayout();
		((Control)statusStrip1).ResumeLayout(false);
		((Control)statusStrip1).PerformLayout();
		((ISupportInitialize)winSpMenu).EndInit();
		((ISupportInitialize)winSpMenu_Sheet1).EndInit();
		((Control)panelMenu).ResumeLayout(false);
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
