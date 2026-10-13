using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Xml;
using FarPoint.Excel;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using FarPoint.Win.Spread.Model;
using HTKK.Class;
using HTKK.Forms;
using HTKK.Forms.Controls;
using HTKKDanhmuc;
using TAX.HTKK.Business;
using TAX.HTKK.Utilities.Control;
using TAX.HTKK.Utilities.Core;
using TAX.HTKK.Utilities.Entity;
using TAX.HTKK.Utilities.Implement;

namespace HTKK;

public class frmWinToKhai : Form
{
	private static XmlDocument xmlConfigHeader = new XmlDocument();

	private static string FileNameLoading = string.Empty;

	private static Dictionary<string, int> dicTimKiemValue;

	private static int rowIndexConfig = 0;

	private bool isDataChanged = false;

	private bool isDeleteEvent = false;

	private bool isDataValid = false;

	private bool isClickTongHopSNNC = false;

	private bool isClickTongHopKHBS = false;

	private bool isDataChangedCallTHKhbs = false;

	private string maTinhTP = "";

	private string maQuanHuyen = "";

	private Dictionary<string, string> selectValueOld = new Dictionary<string, string>();

	private int index = 21;

	private int indexDADT = 20;

	private int startRowTimKiem = 21;

	private int startRowTimKiemDADT = 20;

	private int totalRow;

	private Row ar;

	private FarPoint.Win.Spread.Cell ac;

	private FarPoint.Win.Spread.Cell cellOld;

	private FarPoint.Win.Spread.Cell cellActiveWhileTab;

	private bool isErrCellOld;

	private List<string> result = new List<string>();

	private List<ControlProperties> listControlProperties = new List<ControlProperties>();

	private static Dictionary<string, int> dicScroll = new Dictionary<string, int>();

	private static int scrollCurrent = 0;

	private int chonMCTTruocKHBS = 0;

	private Dictionary<string, string> dicToKhaiPhuLucKHBS = new Dictionary<string, string>();

	private Dictionary<string, string> dicTieuMucKHBS = new Dictionary<string, string>();

	private string[] listTenChiTieuKHBS;

	private XmlNode _mapMct = null;

	private XmlNode _mapMctTieuMuc = null;

	private string _oldValue = null;

	private string _tenDai = "";

	private IContainer components = null;

	private OpenFileDialog opfdlg;

	private FpSpread winFpSpread;

	private SheetView winFpSpread_Sheet1;

	private Panel pnWinSpread;

	private Panel pnTimKiem;

	private Button btnTimKiem;

	private TextBox txtTimKiem;

	private ComboBox cbbTimKiem;

	private Button cmdClear;

	private Button cmdLoadToKhai;

	private Button cmdSave;

	private Button cmdPrint;

	private Button cmdDelete;

	private Button cmdKiemTra;

	private Button btnThemPhuLuc;

	private Button cmdExport;

	private Button cmdImportXML;

	private Button cmdExit;

	private Button btnTaiBangKe;

	private Button btnTongHopKHBS;

	private Panel pnButon;

	private Panel pnF7;

	private Label label1;

	private Label label2;

	private Label lblInfo;

	private Button btnCustom;

	public Panel panelChild;

	private Panel pnF5;

	private Label lblThemDong;

	private Label lblXoaDong;

	private Panel pnInsert;

	private Panel panelCQT;

	private ComboBox cbbCQTChiCuc;

	private Label label3;

	private Label lbCQTCapCuc;

	private ComboBox cbbCQTCuc;

	private Button btnEndPage;

	private Button btnNextPage;

	private Label labelTotalPage;

	private Label label4;

	private TextBox txtPage;

	private Button btnPrivPage;

	private Button btnFirstPage;

	private Panel panelPage;

	public frmWinToKhai(XmlNode _xmlNode)
	{
		if (File.Exists("ProgressIndicatorTest.exe"))
		{
			Process.Start("ProgressIndicatorTest.exe");
		}
		InitializeComponent();
	}

	private void frmToKhai_Load(object sender, EventArgs e)
	{
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e6: Expected O, but got Unknown
		//IL_00fc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0102: Expected O, but got Unknown
		//IL_0121: Unknown result type (might be due to invalid IL or missing references)
		//IL_0128: Expected O, but got Unknown
		//IL_0148: Unknown result type (might be due to invalid IL or missing references)
		//IL_014f: Expected O, but got Unknown
		//IL_016f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0176: Expected O, but got Unknown
		//IL_0196: Unknown result type (might be due to invalid IL or missing references)
		//IL_019d: Expected O, but got Unknown
		//IL_01bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c4: Expected O, but got Unknown
		//IL_02b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0268: Unknown result type (might be due to invalid IL or missing references)
		MouseEventHandler val = null;
		MouseEventHandler val2 = null;
		MouseEventHandler val3 = null;
		MouseEventHandler val4 = null;
		MouseEventHandler val5 = null;
		MouseEventHandler val6 = null;
		try
		{
			FormProcess.ChildFormClosed((Form)(object)this, (Control)(object)panelChild);
			frmToKhai_ReLoad();
			InputMap inputMap = winFpSpread.GetInputMap(InputMapMode.WhenAncestorOfFocused);
			inputMap.Put(new Keystroke((Keys)33, (Keys)131072), SpreadActions.MoveToNextSheet);
			inputMap.Put(new Keystroke((Keys)34, (Keys)131072), SpreadActions.MoveToPreviousSheet);
			inputMap.Put(new Keystroke((Keys)13, (Keys)0), SpreadActions.None);
			inputMap = winFpSpread.GetInputMap(InputMapMode.WhenFocused);
			inputMap.Put(new Keystroke((Keys)33, (Keys)131072), SpreadActions.MoveToNextSheet);
			inputMap.Put(new Keystroke((Keys)34, (Keys)131072), SpreadActions.MoveToPreviousSheet);
			inputMap.Put(new Keystroke((Keys)13, (Keys)0), SpreadActions.None);
			WinSettings.myWinToKhai = this;
			((Control)winFpSpread).MouseWheel += new MouseEventHandler(winFpSpread_MouseWheel);
			SpreadVScrollBar verticalScrollBar = winFpSpread.VerticalScrollBar;
			if (val == null)
			{
				val = (MouseEventHandler)delegate
				{
					if (!winFpSpread.IsLayoutSuspended)
					{
						winFpSpread.SuspendLayout();
						SetActiveCellToObject(winFpSpread.ActiveSheetIndex);
						winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveRowIndex, 0);
					}
				};
			}
			verticalScrollBar.MouseDown += val;
			SpreadVScrollBar verticalScrollBar2 = winFpSpread.VerticalScrollBar;
			if (val2 == null)
			{
				val2 = (MouseEventHandler)delegate
				{
					if (!winFpSpread.IsLayoutSuspended)
					{
						winFpSpread.SuspendLayout();
						ac = winFpSpread.ActiveSheet.ActiveCell;
						scrollCurrent = ac.Row.Index;
						winFpSpread.ActiveSheet.SetActiveCell(scrollCurrent, 0);
						winFpSpread.ResumeLayout();
					}
				};
			}
			((Control)verticalScrollBar2).MouseWheel += val2;
			SpreadVScrollBar verticalScrollBar3 = winFpSpread.VerticalScrollBar;
			if (val3 == null)
			{
				val3 = (MouseEventHandler)delegate
				{
					if (winFpSpread.IsLayoutSuspended)
					{
						int[] activeCellFromObject = GetActiveCellFromObject(winFpSpread.ActiveSheetIndex);
						winFpSpread.ActiveSheet.SetActiveCell(activeCellFromObject[0], activeCellFromObject[1]);
						winFpSpread.ResumeLayout();
					}
				};
			}
			verticalScrollBar3.MouseUp += val3;
			SpreadHScrollBar horizontalScrollBar = winFpSpread.HorizontalScrollBar;
			if (val4 == null)
			{
				val4 = (MouseEventHandler)delegate
				{
					if (!winFpSpread.IsLayoutSuspended)
					{
						winFpSpread.SuspendLayout();
						SetActiveCellToObject(winFpSpread.ActiveSheetIndex);
						winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveRowIndex, 0);
					}
				};
			}
			horizontalScrollBar.MouseDown += val4;
			SpreadHScrollBar horizontalScrollBar2 = winFpSpread.HorizontalScrollBar;
			if (val5 == null)
			{
				val5 = (MouseEventHandler)delegate
				{
					if (!winFpSpread.IsLayoutSuspended)
					{
						winFpSpread.SuspendLayout();
						ac = winFpSpread.ActiveSheet.ActiveCell;
						scrollCurrent = ac.Row.Index;
						winFpSpread.ActiveSheet.SetActiveCell(scrollCurrent, 0);
						winFpSpread.ResumeLayout();
					}
				};
			}
			((Control)horizontalScrollBar2).MouseWheel += val5;
			SpreadHScrollBar horizontalScrollBar3 = winFpSpread.HorizontalScrollBar;
			if (val6 == null)
			{
				val6 = (MouseEventHandler)delegate
				{
					if (winFpSpread.IsLayoutSuspended)
					{
						int[] activeCellFromObject = GetActiveCellFromObject(winFpSpread.ActiveSheetIndex);
						winFpSpread.ActiveSheet.SetActiveCell(activeCellFromObject[0], activeCellFromObject[1]);
						winFpSpread.ResumeLayout();
					}
				};
			}
			horizontalScrollBar3.MouseUp += val6;
			if (WinSettings.TTPeriodInfo.dicVar != null && WinSettings.TTPeriodInfo.dicVar.Count > 0 && WinSettings.TTPeriodInfo.dicVar.ContainsKey("IS_DM_MOI") && (WinSettings.MenuID == "06" || WinSettings.MenuID == "77" || WinSettings.MenuID == "28"))
			{
				frmMessage frmMessage = new frmMessage("0497", "3", "1", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
			}
			if (WinSettings.MenuID == "86" || WinSettings.MenuID == "87")
			{
				frmMessage frmMessage = new frmMessage("0500", "3", "1", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
			}
			if (WinSettings.MenuID == "U8")
			{
				showPanelPage(0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmToKhai_ReLoad()
	{
		try
		{
			Stopwatch stopwatch = Stopwatch.StartNew();
			stopwatch.Reset();
			stopwatch.Start();
			((Control)pnTimKiem).Visible = false;
			((Control)panelCQT).Visible = false;
			InitTimKiem(0);
			LoadButton();
			LoadToKhai();
			winFpSpread.ButtonDrawMode = ButtonDrawModes.Always;
			winFpSpread.ActiveSheetIndex = 0;
			if (WinSettings.MenuID == "Y7")
			{
				if (!WinSettings.ToKhaiInfo.SheetInfos[1].IsNullOrEmpty())
				{
					winFpSpread.ActiveSheetIndex = 1;
				}
				else if (!WinSettings.ToKhaiInfo.SheetInfos[2].IsNullOrEmpty())
				{
					winFpSpread.ActiveSheetIndex = 2;
				}
			}
			DisplayF5F7(winFpSpread.ActiveSheetIndex);
			stopwatch.Stop();
			long elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
			HTKKFile.LogRun(MethodBase.GetCurrentMethod().Name, "All_Load - " + elapsedMilliseconds);
			((Control)pnTimKiem).Visible = WinSettings.MenuID.Equals("57");
			if (WinSettings.MenuID.Equals("G2"))
			{
				((Control)pnTimKiem).Visible = true;
				((Control)panelCQT).Visible = true;
				((Control)panelCQT).Location = new Point(12, 5);
				((Control)pnTimKiem).Location = new Point(20 + ((Control)panelCQT).Size.Width, 5);
			}
			((Control)this).Text = "Hỗ trợ kê khai thuế -  Phiên bản " + WinSettings.VersionApp;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		finally
		{
			Process[] processesByName = Process.GetProcessesByName("ProgressIndicatorTest");
			foreach (Process process in processesByName)
			{
				process.Kill();
			}
		}
	}

	private void LoadToKhai()
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			FileInfo file = new FileInfo(WinSettings.ToKhaiInfo.InterfaceTemplateExcel);
			if (HTKKFile.IsFileLocked(file))
			{
				frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
				((Form)this).Close();
				return;
			}
			winFpSpread.SuspendLayout();
			bool flag = WinFnProcess.WinLoadFromFile(winFpSpread, WinSettings.ToKhaiInfo.InterfaceTemplateExcel);
			WinSpreadProcess.fpSpDisplayTemplate(winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
			if (flag & (WinSettings.ToKhaiInfo.FolderName == "HEADER"))
			{
				winFpSpread.Sheets[0].Cells[1, 19].Text = WinSettings.HTKKHeaderInfo.mst;
			}
			winFpSpread.ResumeLayout();
			Default_01_1_KHBS();
			ArrangeHorizonPanel(pnButon);
			GeneralFunctions.checkLockUnLockCells(WinSettings.ToKhaiInfo, winFpSpread);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void form_DoWorkDelete(ProgressForm sender, DoWorkEventArgs e)
	{
		WinDoWorkEvent.WinDoWorkDelete(winFpSpread, WinSettings.DicSheet, winFpSpread.ActiveSheetIndex);
	}

	private void fpSpread1_ScrollTipFetch(object sender, ScrollTipFetchEventArgs e)
	{
		e.ShowTip = true;
		e.TipText = "Di chuyển đến dòng " + e.Row;
	}

	private void form_DoWorkKiemTra(ProgressForm sender, DoWorkEventArgs e)
	{
		bool flag = (bool)e.Argument;
		if (!isDataValid || isDataChanged)
		{
			WinDoWorkEvent.WinDoWorkKiemTra(winFpSpread);
		}
	}

	private void form_DoWorkReindexTable(ProgressForm sender, DoWorkEventArgs e)
	{
		int num = ConvertHelpers.ConvertToInt(e.Argument);
		WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.ActiveSheet, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[num]);
	}

	private void fpSpread1_EditModeStarting(object sender, EditModeStartingEventArgs e)
	{
	}

	private void winFpSpread_EditModeOn(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		KeyPressEventHandler val = new KeyPressEventHandler(winFpSpread_KeyPress);
		winFpSpread.EditingControl.KeyPress += val;
		_oldValue = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveRowIndex, winFpSpread.ActiveSheet.ActiveColumnIndex].Text;
	}

	private void winFpSpread_EditModeOff(object sender, EventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		KeyPressEventHandler val = new KeyPressEventHandler(winFpSpread_KeyPress);
		winFpSpread.EditingControl.KeyPress -= val;
		string text = winFpSpread.ActiveSheet.ActiveCell.Text;
		if (!string.IsNullOrEmpty(text) && winFpSpread.ActiveSheet.ActiveRow.Index < 900)
		{
			Row activeRow = winFpSpread.ActiveSheet.ActiveRow;
			float preferredHeight = activeRow.GetPreferredHeight();
			if (winFpSpread.ActiveSheet.SheetName.Equals("KHBS"))
			{
				if (activeRow.Height < preferredHeight)
				{
					activeRow.Height = preferredHeight;
				}
			}
			else if (activeRow.Height < preferredHeight - 5f || activeRow.Height > preferredHeight + 5f)
			{
				activeRow.Height = preferredHeight;
			}
		}
		if (!string.IsNullOrEmpty(_tenDai))
		{
			winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveRowIndex, winFpSpread.ActiveSheet.ActiveColumnIndex].Value = _tenDai;
			_tenDai = "";
		}
	}

	private void winFpSpread_KeyPress(object sender, KeyPressEventArgs e)
	{
		if (e.KeyChar == '\r' && sender.GetType().ToString() != "FarPoint.Win.FpCombo")
		{
			e.Handled = true;
			SendKeys.Send("{TAB}");
			return;
		}
		int activeSheetIndex = winFpSpread.ActiveSheetIndex;
		int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
		int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
		HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(activeRowIndex, activeColumnIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
		if (staticCell != null)
		{
			if (staticCell.Controltype == "14" || staticCell.Controltype == "19" || staticCell.Controltype == "104" || staticCell.Controltype == "109" || staticCell.Controltype == "96" || staticCell.Controltype == "94")
			{
				GeneralFunctions.CompareKeycode(e, 1);
			}
		}
		else
		{
			int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
			if (tableIndex >= 0)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
				staticCell = WinFnProcess.GetStaticCell(activeRowIndex, activeColumnIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Header.StaticCells);
				if (staticCell != null)
				{
					if (staticCell.Controltype == "14" || staticCell.Controltype == "19" || staticCell.Controltype == "104" || staticCell.Controltype == "109" || staticCell.Controltype == "96" || staticCell.Controltype == "94")
					{
						GeneralFunctions.CompareKeycode(e, 1);
					}
				}
				else
				{
					staticCell = WinFnProcess.GetStaticCell(activeRowIndex, activeColumnIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Footer.StaticCells);
					if (staticCell != null && (staticCell.Controltype == "14" || staticCell.Controltype == "19" || staticCell.Controltype == "104" || staticCell.Controltype == "109" || staticCell.Controltype == "96" || staticCell.Controltype == "94"))
					{
						GeneralFunctions.CompareKeycode(e, 1);
					}
					if ((activeRowIndex + 1 >= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart) & (activeRowIndex + 1 <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count))
					{
						for (int i = 0; i < hTKKCMCTableInfo.RowsInfo.Cells.Count; i++)
						{
							if (hTKKCMCTableInfo.RowsInfo.Cells[i].CellTemplate.ColIndex == activeColumnIndex + 1)
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "14" || hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "19" || hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "104" || hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "109" || hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "96" || hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "94")
								{
									GeneralFunctions.CompareKeycode(e, 1);
								}
								string[] source = "106~".Split(new char[1] { '~' });
								if (source.Contains(hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype) && e.KeyChar == '\r')
								{
									autoCombobox(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRowIndex + 1, winFpSpread.ActiveSheet.ActiveColumnIndex + 1, hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex]);
								}
							}
						}
					}
				}
			}
		}
		WinDoWorkEvent.KeyPress(winFpSpread, activeSheetIndex, activeRowIndex, activeColumnIndex, e, (Keys)0);
	}

	protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Invalid comparison between Unknown and I4
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001a: Invalid comparison between Unknown and I4
		//IL_059f: Unknown result type (might be due to invalid IL or missing references)
		//IL_05a2: Invalid comparison between Unknown and I4
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_001f: Invalid comparison between Unknown and I4
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Invalid comparison between Unknown and I4
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Invalid comparison between Unknown and I4
		//IL_08f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_08f7: Invalid comparison between Unknown and I4
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Invalid comparison between Unknown and I4
		//IL_09af: Unknown result type (might be due to invalid IL or missing references)
		//IL_09b2: Invalid comparison between Unknown and I4
		//IL_0943: Unknown result type (might be due to invalid IL or missing references)
		//IL_12da: Unknown result type (might be due to invalid IL or missing references)
		//IL_12dd: Invalid comparison between Unknown and I4
		//IL_1e9e: Unknown result type (might be due to invalid IL or missing references)
		//IL_1ea1: Invalid comparison between Unknown and I4
		//IL_2112: Unknown result type (might be due to invalid IL or missing references)
		//IL_2115: Invalid comparison between Unknown and I4
		//IL_17bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_1e68: Unknown result type (might be due to invalid IL or missing references)
		//IL_0df9: Unknown result type (might be due to invalid IL or missing references)
		//IL_2103: Unknown result type (might be due to invalid IL or missing references)
		//IL_128b: Unknown result type (might be due to invalid IL or missing references)
		//IL_25e1: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			GeneralFunctions.KEY_PRESS = keyData;
			if ((int)keyData == 116 || (int)keyData == 117 || (int)keyData == 118 || (int)keyData == 119 || (int)keyData != 112)
			{
			}
			if ((int)keyData == 113)
			{
				try
				{
					int activeSheetIndex = winFpSpread.ActiveSheetIndex;
					int activeRowIndex = winFpSpread.ActiveSheet.ActiveRowIndex;
					int activeColumnIndex = winFpSpread.ActiveSheet.ActiveColumnIndex;
					bool flag = false;
					ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
					ICellType cellType = winFpSpread.Sheets[activeSheetIndex].Cells[activeRowIndex, activeColumnIndex].CellType;
					if (cellType != null && (cellType.ToString() == "ComboBoxCellType" || cellType.ToString().Contains("ComboBoxCellType")))
					{
						HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
						for (int i = 0; i < hTKKCMCSheetInfo.StaticCells.Count; i++)
						{
							if (((hTKKCMCSheetInfo.StaticCells[i].CellTemplate.RowIndex == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (hTKKCMCSheetInfo.StaticCells[i].CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1)) && hTKKCMCSheetInfo.StaticCells[i].CellSelectedValue != null)
							{
								winFpSpread.ActiveSheet.Cells[hTKKCMCSheetInfo.StaticCells[i].CellSelectedValue.RowIndex - 1, hTKKCMCSheetInfo.StaticCells[i].CellSelectedValue.ColIndex - 1].Text = "";
								break;
							}
						}
						for (int j = 0; j < hTKKCMCSheetInfo.Tables.Count; j++)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo.Tables[j];
							HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
							HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
							int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
							List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
							if (hTKKCMCTableInfo.Header.StaticCells.Count > 0)
							{
								for (int i = 0; i < hTKKCMCTableInfo.Header.StaticCells.Count; i++)
								{
									HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells[i];
									if ((hTKKCMCCellInfo.CellTemplate.RowIndex == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (hTKKCMCCellInfo.CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1))
									{
										winFpSpread.ActiveSheet.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Text = "";
										break;
									}
								}
							}
							if (hTKKCMCTableInfo.Footer.StaticCells.Count > 0)
							{
								for (int i = 0; i < hTKKCMCTableInfo.Footer.StaticCells.Count; i++)
								{
									HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Footer.StaticCells[i];
									if ((hTKKCMCCellInfo.CellTemplate.RowIndex == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (hTKKCMCCellInfo.CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1))
									{
										winFpSpread.ActiveSheet.Cells[hTKKCMCCellInfo.CellSelectedValue.RowIndex - 1, hTKKCMCCellInfo.CellSelectedValue.ColIndex - 1].Text = "";
										break;
									}
								}
							}
							for (int i = 0; i < rowsValues.Count; i++)
							{
								for (int k = 0; k < rowsInfo.Cells.Count; k++)
								{
									if ((dataLocationTemplateRowStart + i == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (rowsInfo.Cells[k].CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1))
									{
										winFpSpread.ActiveSheet.Cells[dataLocationTemplateRowStart + i - 1, HTKKFile.ColumnLettersToIndex(rowsInfo.Cells[k].indexColumnSelectedValue) - 1].Text = "";
										flag = true;
										break;
									}
								}
								if (flag)
								{
									break;
								}
							}
						}
						selectValueOld = new Dictionary<string, string>();
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
			if ((int)keyData == 112)
			{
				for (int i = 0; i < sInfo.StaticCells.Count; i++)
				{
					if ((sInfo.StaticCells[i].CellTemplate.RowIndex == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (sInfo.StaticCells[i].CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1))
					{
						string fullPath = Path.GetFullPath("HTKK.chm");
						HelpUtil.HelpPath = fullPath;
						HelpUtil.HtmFolder = "HTKK";
						HelpUtil.Show((Form)(object)this, sInfo.StaticCells[i].HelpContextID);
						return true;
					}
				}
				for (int j = 0; j < sInfo.Tables.Count; j++)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[j];
					HTKKCMCDataLocation location = hTKKCMCTableInfo.Location;
					HTKKCMCRowInfo rowsInfo = hTKKCMCTableInfo.RowsInfo;
					int dataLocationTemplateRowStart = location.DataLocationTemplateRowStart;
					List<HTKKCMCRowValue> rowsValues = hTKKCMCTableInfo.RowsValues;
					if (hTKKCMCTableInfo.Header.StaticCells.Count > 0)
					{
						for (int i = 0; i < hTKKCMCTableInfo.Header.StaticCells.Count; i++)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells[i];
							if ((hTKKCMCCellInfo.CellTemplate.RowIndex == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (hTKKCMCCellInfo.CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1))
							{
								string fullPath = Path.GetFullPath("HTKK.chm");
								HelpUtil.HelpPath = fullPath;
								HelpUtil.HtmFolder = "HTKK";
								HelpUtil.Show((Form)(object)this, hTKKCMCCellInfo.HelpContextID);
								return true;
							}
						}
					}
					for (int i = 0; i < rowsValues.Count; i++)
					{
						for (int k = 0; k < rowsInfo.Cells.Count; k++)
						{
							if ((dataLocationTemplateRowStart + i == winFpSpread.ActiveSheet.ActiveRow.Index + 1) & (rowsInfo.Cells[k].CellTemplate.ColIndex == winFpSpread.ActiveSheet.ActiveColumn.Index + 1))
							{
								string fullPath = Path.GetFullPath("HTKK.chm");
								HelpUtil.HelpPath = fullPath;
								HelpUtil.HtmFolder = "HTKK";
								HelpUtil.Show((Form)(object)this, rowsInfo.Cells[k].HelpContextID);
								return true;
							}
						}
					}
				}
			}
			if (WinSettings.MenuID.Equals("B8") && (int)keyData == 114)
			{
				WinDoWorkEvent.KeyPress(winFpSpread, winFpSpread.ActiveSheetIndex, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, keyData);
				return true;
			}
			int iCurTable = sInfo.GetTableIndex(winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
			if (iCurTable < 0)
			{
				return false;
			}
			if ((int)keyData == 116 && ((Control)pnF5).Visible)
			{
				if (sInfo.TaiDuLieuRatLon == "1")
				{
					winFpSpread.SuspendLayout();
					isDataChanged = true;
					int num = winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1;
					int num2 = GeneralFunctions._page * GeneralFunctions._pageSize + num;
					int num3 = GeneralFunctions.timBangBatDauSView(sInfo, num2);
					if (num3 < 0)
					{
						return false;
					}
					HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num3];
					HTKKCMCRowValue rowValueDefault = GeneralFunctions.GetRowValueDefault(hTKKCMCTableInfo.RowsInfo);
					int num4 = 0;
					if (num3 > 0 && !GeneralFunctions.dataInfoOld.listTable[num3 - 1].isShow)
					{
						num4++;
					}
					int num5 = num2 - (hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + num4);
					sInfo.Tables[num3].RowsValues.Insert(num5 + 1, rowValueDefault);
					GeneralFunctions.InsertTableRowNoSView(sInfo.Tables[num3].Location.DataLocationTemplateRowStart, 1, sInfo);
					GeneralFunctions.ReindexColStt(sInfo.Tables[num3]);
					WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex] = sInfo;
					WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
					GeneralFunctions._totalPage = GeneralFunctions.tinhTotalPage_Large(sInfo, WinSettings.MenuID);
					((Control)labelTotalPage).Text = GeneralFunctions._totalPage.ToString();
					if (Convert.ToInt32(((Control)txtPage).Text) > Convert.ToInt32(((Control)labelTotalPage).Text))
					{
						((Control)txtPage).Text = ((Control)labelTotalPage).Text;
					}
					winFpSpread.ResumeLayout();
					return true;
				}
				if ((WinSettings.MenuID.Equals("G1") || WinSettings.MenuID.Equals("G4") || WinSettings.MenuID.Equals("N5") || WinSettings.MenuID.Equals("H5") || WinSettings.MenuID.Equals("O7") || WinSettings.MenuID.Equals("P5") || WinSettings.MenuID.Equals("U6")) && (iCurTable == 1 || iCurTable == 2))
				{
					return false;
				}
				if (WinSettings.MenuID.Equals("N8") && iCurTable == 0 && winFpSpread.ActiveSheetIndex == 0 && Convert.ToInt32(WinSettings.TTPeriodInfo.Nam) < 2025)
				{
					return false;
				}
				if (WinSettings.MenuID.Equals("L9"))
				{
					return false;
				}
				if (WinSettings.MenuID.Equals("99") && iCurTable == 0 && winFpSpread.ActiveSheetIndex == 0)
				{
					return false;
				}
				if ((WinSettings.MenuID.Equals("D4") || WinSettings.MenuID.Equals("Z3")) && iCurTable == 0 && winFpSpread.ActiveSheetIndex == 0 && sInfo.Tables[iCurTable].RowsValues.Count >= 2)
				{
					frmMessage frmMessage = new frmMessage("0585", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return true;
				}
				if ("01~03~94~96~97~90~".Contains(WinSettings.MenuID))
				{
					winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveRow.Index, 0);
				}
				int row = 0;
				isDataChanged = true;
				if (winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 2 && winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 3)
				{
					isDataChangedCallTHKhbs = true;
				}
				if (sInfo.Tables[iCurTable].RowsValues.Count == 1 && sInfo.Tables[iCurTable].RowsValues[0].RowDisplay == "0")
				{
					winFpSpread.ActiveSheet.Rows[sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart - 1].Visible = true;
					sInfo.Tables[iCurTable].RowsValues[0].RowDisplay = "1";
					for (int i = 0; i < sInfo.Tables[iCurTable].RowsInfo.Cells.Count; i++)
					{
						WinSpreadProcess.fpSpBindDataCell(winFpSpread.ActiveSheet, sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[iCurTable].RowsInfo.Cells[i].Controltype, sInfo.Tables[iCurTable].RowsInfo.Cells[i].DefaultValue);
						row = sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart;
					}
				}
				else
				{
					int num = winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1;
					if (num < sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart)
					{
						num = sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + sInfo.Tables[iCurTable].RowsValues.Count - 1;
					}
					else if (num > sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + sInfo.Tables[iCurTable].RowsValues.Count)
					{
						num = sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart + sInfo.Tables[iCurTable].RowsValues.Count - 1;
					}
					WinSpreadProcess.InsertTableRow(winFpSpread.ActiveSheet, num, -1, num, -1, -1, sInfo, iCurTable);
					row = num + 1;
				}
				WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.ActiveSheet, sInfo.Tables[iCurTable]);
				WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
				WinDoWorkEvent.KeyPress(winFpSpread, winFpSpread.ActiveSheetIndex, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, keyData);
				if (winFpSpread.ActiveSheet.SheetName == "01-1/KHBS")
				{
					SetDefaultManualRow_01_1_KHBS(winFpSpread.ActiveSheet, iCurTable, row);
				}
			}
			Dictionary<int, List<string>> dictionary = new Dictionary<int, List<string>>();
			if ((int)keyData == 117 && ((Control)pnF5).Visible)
			{
				isDataChanged = true;
				if (sInfo.TaiDuLieuRatLon == "1")
				{
					winFpSpread.SuspendLayout();
					int num = winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1;
					int num2 = GeneralFunctions._page * GeneralFunctions._pageSize + num;
					int num3 = GeneralFunctions.timBangBatDauSView(sInfo, num2);
					if (num3 < 0)
					{
						return false;
					}
					HTKKCMCTableInfo hTKKCMCTableInfo = sInfo.Tables[num3];
					int num5 = num2 - hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					if (WinSettings.MenuID == "P2" && num3 > 0 && !GeneralFunctions.dataInfoOld.listTable[0].isShow)
					{
						num5--;
					}
					sInfo.Tables[num3].RowsValues.RemoveAt(num5);
					if (sInfo.Tables[num3].RowsValues.Count > 0)
					{
						GeneralFunctions.InsertTableRowNoSView(sInfo.Tables[num3].Location.DataLocationTemplateRowStart, -1, sInfo);
					}
					if (num5 == 0 && sInfo.Tables[num3].RowsValues.Count == 0)
					{
						sInfo.Tables[num3].RowsValues.Add(GeneralFunctions.GetRowValueDefault(sInfo.Tables[num3].RowsInfo));
					}
					GeneralFunctions.ReindexColStt(sInfo.Tables[num3]);
					WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex] = sInfo;
					WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
					GeneralFunctions._totalPage = GeneralFunctions.tinhTotalPage_Large(sInfo, WinSettings.MenuID);
					((Control)labelTotalPage).Text = GeneralFunctions._totalPage.ToString();
					if (Convert.ToInt32(((Control)txtPage).Text) > Convert.ToInt32(((Control)labelTotalPage).Text))
					{
						((Control)txtPage).Text = ((Control)labelTotalPage).Text;
					}
					winFpSpread.ResumeLayout();
					return true;
				}
				if (winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 2 && winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 3)
				{
					isDataChangedCallTHKhbs = true;
				}
				int num6 = winFpSpread.ActiveSheet.ActiveRow.Index;
				int column = winFpSpread.ActiveSheet.ActiveColumn.Index;
				int frRow = winFpSpread.ActiveSheet.ActiveRow.Index;
				dictionary = saveHeaderIfDeleteF6F8(winFpSpread, WinSettings.ToKhaiInfo, winFpSpread.ActiveSheetIndex, num6 + 1);
				if (sInfo.Tables[iCurTable].RowsValues.Count > 1)
				{
					num6 = ((winFpSpread.ActiveSheet.ActiveRow.Index != sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart - 1 + sInfo.Tables[iCurTable].RowsValues.Count - 1) ? winFpSpread.ActiveSheet.ActiveRow.Index : (winFpSpread.ActiveSheet.ActiveRow.Index - 1));
				}
				if (WinSettings.MenuID == "G2")
				{
					string text = "";
					int num7 = HTKKFile.ColumnLettersToIndex("AC");
					text = winFpSpread.Sheets[0].Cells[winFpSpread.ActiveSheet.ActiveRow.Index, num7 - 1].Text;
					if (text.Equals("0"))
					{
						frmMessage frmMessage = new frmMessage("0527", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						return true;
					}
				}
				if (winFpSpread.ActiveSheet.ActiveRow.Index + 1 == sInfo.Tables[iCurTable].Location.TableTemplateRowStart && sInfo.Tables[iCurTable].RowsInfo.Cells[0].DefaultValue.Equals("NA"))
				{
					for (int i = 1; i < sInfo.Tables[iCurTable].RowsInfo.Cells.Count; i++)
					{
						WinSpreadProcess.fpSpBindDataCell(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index + 1, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[iCurTable].RowsInfo.Cells[i].Controltype, sInfo.Tables[iCurTable].RowsInfo.Cells[i].DefaultValue);
					}
					return true;
				}
				if (sInfo.Tables[iCurTable].RowsValues.Count == 1)
				{
					string cellFormat = HTKKFile.getCellFormat(WinSettings.MenuID, winFpSpread.ActiveSheetIndex, iCurTable, 2);
					string cellFormat2 = HTKKFile.getCellFormat(WinSettings.MenuID, winFpSpread.ActiveSheetIndex, iCurTable, 3);
					for (int i = 0; i < sInfo.Tables[iCurTable].RowsInfo.Cells.Count; i++)
					{
						WinSpreadProcess.fpSpBindDataCell(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index + 1, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex, sInfo.Tables[iCurTable].RowsInfo.Cells[i].Controltype, sInfo.Tables[iCurTable].RowsInfo.Cells[i].DefaultValue);
						if (cellFormat == null)
						{
							continue;
						}
						if (cellFormat.IndexOf("~" + HTKKFile.ColumnIndexToLetters(sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex) + "~") > -1)
						{
							HTKKFile.RemoveValidateBackGround(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1);
							continue;
						}
						string text2 = HTKKFile.ColumnIndexToLetters(sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex);
						if (HTKKFormater.isNullString(cellFormat2) || cellFormat2.IndexOf("~" + HTKKFile.ColumnIndexToLetters(sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex) + "~") == -1)
						{
							HTKKFile.RemoveValidate(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index, sInfo.Tables[iCurTable].RowsInfo.Cells[i].CellTemplate.ColIndex - 1);
						}
					}
				}
				if (sInfo.Tables[iCurTable].RowsValues.Count > 1)
				{
					winFpSpread.SuspendLayout();
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveRow.Index, winFpSpread.ActiveSheet.ActiveColumn.Index].Value = GeneralFunctions.GetValue(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index + 2, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
					sInfo.Tables[iCurTable].RowsValues.RemoveAt(winFpSpread.ActiveSheet.ActiveRow.Index + 1 - sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart);
					WinSpreadProcess.RemoveDataRows(winFpSpread.ActiveSheet, frRow, 1, sInfo);
					winFpSpread.ActiveSheet.SetActiveCell(num6, column);
					winFpSpread.ResumeLayout();
				}
				WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.ActiveSheet, sInfo.Tables[iCurTable]);
				WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
				WinDoWorkEvent.KeyPress(winFpSpread, winFpSpread.ActiveSheetIndex, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, keyData);
				revertHeaderIfDeleteF6(winFpSpread, WinSettings.ToKhaiInfo, winFpSpread.ActiveSheetIndex, dictionary);
				dictionary.Clear();
				return true;
			}
			if ((int)keyData == 118 && ((Control)pnF7).Visible)
			{
				bool flag2 = false;
				if (sInfo.Tables[iCurTable].Header != null && sInfo.Tables[iCurTable].Header.StaticCells.Count > 0)
				{
					flag2 = true;
				}
				if (sInfo.Tables[iCurTable].Footer != null && sInfo.Tables[iCurTable].Footer.StaticCells.Count > 0)
				{
					flag2 = true;
				}
				if (!flag2)
				{
					return true;
				}
				isDataChanged = true;
				if (winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 2 && winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 3)
				{
					isDataChangedCallTHKhbs = true;
				}
				if ("01~03~94~96~97~90~".Contains(WinSettings.MenuID))
				{
					winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveRow.Index, 0);
				}
				if (!"G9,Q8,R7,U5".Contains(WinSettings.MenuID))
				{
					WinSpreadProcess.InsertTable(winFpSpread.ActiveSheet, iCurTable, sInfo);
					WinSpreadProcess.fpSpReindexingSheetTemplate(winFpSpread.ActiveSheet, sInfo);
					WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
				}
				else
				{
					WinSpreadProcess.InsertTableNew(winFpSpread.ActiveSheet, iCurTable, sInfo);
					WinSpreadProcess.fpSpReindexingSheetTemplate(winFpSpread.ActiveSheet, sInfo);
				}
				WinDoWorkEvent.KeyPress(winFpSpread, winFpSpread.ActiveSheetIndex, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, keyData);
				return true;
			}
			if ((int)keyData == 119 && ((Control)pnF7).Visible)
			{
				bool flag2 = false;
				if (sInfo.Tables[iCurTable].Header != null && sInfo.Tables[iCurTable].Header.StaticCells.Count > 0)
				{
					flag2 = true;
				}
				if (sInfo.Tables[iCurTable].Footer != null && sInfo.Tables[iCurTable].Footer.StaticCells.Count > 0)
				{
					flag2 = true;
				}
				if (!flag2)
				{
					return true;
				}
				isDataChanged = true;
				if (winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 2 && winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 3)
				{
					isDataChangedCallTHKhbs = true;
				}
				dictionary = saveHeaderIfDeleteF6F8(winFpSpread, WinSettings.ToKhaiInfo, winFpSpread.ActiveSheetIndex, sInfo.Tables[iCurTable].Location.DataLocationTemplateRowStart);
				if ("G9,Q8,R7,U5".Contains(WinSettings.MenuID))
				{
					bool flag3 = GeneralFunctions.RemoveTableF8New(winFpSpread.ActiveSheet, sInfo, iCurTable, setActiveCell: true);
					WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
					revertHeaderIfDeleteF6(winFpSpread, WinSettings.ToKhaiInfo, winFpSpread.ActiveSheetIndex, dictionary);
					dictionary.Clear();
					return flag3;
				}
				if (iCurTable == 0)
				{
					if (sInfo.Tables.Count((HTKKCMCTableInfo p) => p.TableName == sInfo.Tables[iCurTable].TableName) == 1)
					{
						return false;
					}
				}
				else if (iCurTable > 0 && sInfo.Tables[iCurTable].TableName != sInfo.Tables[iCurTable - 1].TableName)
				{
					return false;
				}
				int num6 = winFpSpread.ActiveSheet.ActiveRow.Index;
				int column = winFpSpread.ActiveSheet.ActiveColumn.Index;
				winFpSpread.ActiveSheet.SetActiveCell(0, column);
				if (iCurTable == sInfo.Tables.Count - 1 && iCurTable > 0)
				{
					if (sInfo.Tables[iCurTable - 1].RowsValues.Count > 0)
					{
						num6 = sInfo.Tables[iCurTable - 1].Location.TableTemplateRowStart - 1;
					}
					else
					{
						int num8 = winFpSpread.ActiveSheet.ActiveRow.Index + 1 - sInfo.Tables[iCurTable].Location.TableTemplateRowStart;
						num6 = sInfo.Tables[iCurTable - 1].Location.TableTemplateRowStart - 1 + num8;
					}
				}
				else
				{
					num6 = sInfo.Tables[iCurTable].Location.TableTemplateRowStart - 1;
				}
				WinSpreadProcess.RemoveTable(winFpSpread.ActiveSheet, iCurTable, sInfo);
				WinSpreadProcess.fpSpReindexingSheetTemplate(winFpSpread.ActiveSheet, sInfo);
				WinDoWorkEvent.UpdateFormula(winFpSpread, winFpSpread.ActiveSheetIndex);
				winFpSpread.ActiveSheet.SetActiveCell(num6, column);
				WinDoWorkEvent.KeyPress(winFpSpread, winFpSpread.ActiveSheetIndex, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, null, keyData);
				revertHeaderIfDeleteF6(winFpSpread, WinSettings.ToKhaiInfo, winFpSpread.ActiveSheetIndex, dictionary);
				dictionary.Clear();
				return true;
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private void autoCombobox(SheetView sView, int iRow, int iCol, string controltype = "", HTKKCMCSheetInfo sInfo = null)
	{
		ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
		ICellType cellType = sView.Cells[iRow - 1, iCol - 1].CellType;
		if ((cellType == null || !(cellType.ToString() == "ComboBoxCellType")) && !cellType.ToString().Contains("ComboBoxCellType"))
		{
			return;
		}
		Dictionary<string, string> dictionary = null;
		if (!GeneralFunctions.dicDms.ContainsKey(iCol.ToString()))
		{
			dictionary = GeneralFunctions.GetDMFromCombobox(winFpSpread.ActiveSheet, iRow, iCol);
			GeneralFunctions.dicDms.Add(iCol.ToString(), dictionary);
		}
		else
		{
			dictionary = GeneralFunctions.dicDms[iCol.ToString()];
		}
		string txt = winFpSpread.ActiveSheet.Cells[iRow - 1, iCol - 1].Text;
		if (!string.IsNullOrEmpty(txt))
		{
			IEnumerable<KeyValuePair<string, string>> source = dictionary.Where((KeyValuePair<string, string> x) => x.Value.ToLower().IndexOf(txt.ToLower()) >= 0);
			if (source.Count() > 0)
			{
				WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow, iCol, source.ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value), controltype, deleteText: false, sInfo);
			}
			else
			{
				WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow, iCol, dictionary, controltype, deleteText: true, sInfo);
			}
		}
		else
		{
			WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow, iCol, dictionary, controltype, deleteText: true, sInfo);
		}
	}

	private void cmdSave_Click(object sender, EventArgs e)
	{
		//IL_09f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_07a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_077b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0428: Unknown result type (might be due to invalid IL or missing references)
		//IL_09a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0830: Unknown result type (might be due to invalid IL or missing references)
		//IL_0930: Unknown result type (might be due to invalid IL or missing references)
		//IL_08ba: Unknown result type (might be due to invalid IL or missing references)
		//IL_0478: Unknown result type (might be due to invalid IL or missing references)
		//IL_05b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0582: Unknown result type (might be due to invalid IL or missing references)
		//IL_0507: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_06a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_065a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0633: Unknown result type (might be due to invalid IL or missing references)
		//IL_0551: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
			WinSettings.ToKhaiInfo.TaiDuLieuLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuLon");
			bool flag = true;
			if (WinSettings.ToKhaiInfo.KieuKHBS != "" && !isClickTongHopKHBS && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && !File.Exists(WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[0].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML)))
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
			}
			else if (isDataChangedCallTHKhbs && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
				isDataChangedCallTHKhbs = false;
			}
			bool flag2 = false;
			string text = "#B3#B4#B5#A2#A5#A6#A7#C1#C2#C3#C4#C5#03#";
			if (text.Contains("#" + WinSettings.MenuID + "#"))
			{
				flag2 = true;
			}
			if (!isDataValid || isDataChanged || flag2)
			{
				winFpSpread.SuspendLayout();
				WinDoWorkEvent.WinDoWorkKiemTra(winFpSpread);
				winFpSpread.ResumeLayout();
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				Validate_01_KHBS();
				if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].isValidData)
				{
					flag = false;
					winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].BackColor = Color.Pink;
				}
				if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].isValidData)
				{
					flag = false;
					winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].BackColor = Color.Pink;
				}
			}
			bool flag3 = false;
			for (int i = 0; i < winFpSpread.Sheets.Count; i++)
			{
				if (WinSettings.MenuID == "35")
				{
					if (winFpSpread.TabStrip[i].BackColor == Color.Pink)
					{
						flag = false;
						break;
					}
					if (winFpSpread.TabStrip[i].BackColor == Color.Yellow)
					{
						flag = false;
						flag3 = true;
						break;
					}
				}
				else if (winFpSpread.TabStrip[i].BackColor == Color.Pink)
				{
					flag = false;
					break;
				}
			}
			if (WinDoWorkEvent.firstSheetErr != -1)
			{
				DisplayF5F7(WinDoWorkEvent.firstSheetErr);
			}
			if (!flag)
			{
				if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
				{
					((Control)txtPage).Text = Convert.ToString(GeneralFunctions._page + 1);
				}
				isDataValid = false;
				if (WinSettings.TTPeriodInfo.menuID == "35" && flag3)
				{
					WinDoWorkEvent.WinDoWorkSave(winFpSpread);
					frmMessage frmMessage = new frmMessage("0002", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					isDataChanged = false;
					return;
				}
				bool flag4 = false;
				string text2 = HTKKFile.readLogErrorBigData();
				if (showErrorBigData() && !HTKKFormater.isNullString(text2))
				{
					frmMessageForm frmMessageForm = new frmMessageForm("0015", "5", "Chi tiết lỗi", text2);
					((Form)frmMessageForm).ShowDialog();
					flag4 = frmMessageForm.Yes;
				}
				else if (WinSettings.TTPeriodInfo.menuID == "37" && WinSettings.TTPeriodInfo.LoaiToKhai == "C")
				{
					if (!WinSpreadProcess.CheckNghiaVuNTTNCN(winFpSpread))
					{
						frmMessage frmMessage2 = new frmMessage("0524", "5", "2", "Thông báo");
						frmMessage2.setLabelButton();
						((Form)frmMessage2).ShowDialog();
						if (frmMessage2.Yes)
						{
							isDataChanged = false;
							((Form)this).Close();
						}
						else
						{
							frmMessage2 = new frmMessage("0015", "5", "2", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							flag4 = frmMessage2.Yes;
						}
					}
					else
					{
						frmMessage frmMessage2 = new frmMessage("0015", "5", "2", "Thông báo");
						((Form)frmMessage2).ShowDialog();
						flag4 = frmMessage2.Yes;
					}
				}
				else
				{
					frmMessage frmMessage2 = new frmMessage("0015", "5", "2", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					flag4 = frmMessage2.Yes;
				}
				if (!flag4)
				{
					return;
				}
				if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
				{
					string text3 = WinDoWorkEvent.WinDoWorkSave_Large1(winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
					if (!string.IsNullOrEmpty(text3))
					{
						frmMessageForm frmMessageForm = new frmMessageForm("0015", "5", "Chi tiết lỗi", text3);
						((Form)frmMessageForm).ShowDialog();
						return;
					}
					frmMessage frmMessage3 = new frmMessage("0002", "3", "1", "Thông báo");
					((Form)frmMessage3).ShowDialog();
					isDataChanged = false;
					isDataChangedCallTHKhbs = false;
				}
				else if (!WinDoWorkEvent.WinDoWorkSave(winFpSpread))
				{
					frmMessage frmMessage3 = new frmMessage("0003", "3", "1", "Thông báo");
					((Form)frmMessage3).ShowDialog();
				}
				else
				{
					frmMessage frmMessage3 = new frmMessage("0002", "3", "1", "Thông báo");
					((Form)frmMessage3).ShowDialog();
					isDataChanged = false;
					isDataChangedCallTHKhbs = false;
				}
				return;
			}
			WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
			WinSettings.ToKhaiInfo.TaiDuLieuLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuLon");
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
			{
				string text3 = WinDoWorkEvent.WinDoWorkSave_Large1(winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
				if (!string.IsNullOrEmpty(text3))
				{
					frmMessageForm frmMessageForm = new frmMessageForm("0015", "5", "Chi tiết lỗi", text3);
					((Form)frmMessageForm).ShowDialog();
					return;
				}
				frmMessage frmMessage3 = new frmMessage("0002", "3", "1", "Thông báo");
				((Form)frmMessage3).ShowDialog();
				isDataChanged = false;
				isDataChangedCallTHKhbs = false;
			}
			else if (WinSettings.MenuID == "37" && WinSettings.TTPeriodInfo.LoaiToKhai == "C")
			{
				if (!WinSpreadProcess.CheckNghiaVuNTTNCN(winFpSpread))
				{
					frmMessage frmMessage2 = new frmMessage("0524", "5", "2", "Thông báo");
					frmMessage2.setLabelButton();
					((Form)frmMessage2).ShowDialog();
					if (frmMessage2.Yes)
					{
						isDataChanged = false;
						((Form)this).Close();
						return;
					}
					WinDoWorkEvent.WinDoWorkSave(winFpSpread);
					if (WinSettings.MenuID.Equals("100_1"))
					{
						HTKKFactory.GetInstanceHTKKHeader().SetPropertyHeader(ref WinSettings.HTKKHeaderInfo, WinSettings.HTKKHeaderInfo.mst);
					}
					frmMessage frmMessage3 = new frmMessage("0002", "3", "1", "Thông báo");
					((Form)frmMessage3).ShowDialog();
					isDataChanged = false;
					isDataValid = true;
				}
				else
				{
					WinDoWorkEvent.WinDoWorkSave(winFpSpread);
					if (WinSettings.MenuID.Equals("100_1"))
					{
						HTKKFactory.GetInstanceHTKKHeader().SetPropertyHeader(ref WinSettings.HTKKHeaderInfo, WinSettings.HTKKHeaderInfo.mst);
					}
					frmMessage frmMessage3 = new frmMessage("0002", "3", "1", "Thông báo");
					((Form)frmMessage3).ShowDialog();
					isDataChanged = false;
					isDataValid = true;
				}
			}
			else
			{
				WinDoWorkEvent.WinDoWorkSave(winFpSpread);
				if (WinSettings.MenuID.Equals("100_1"))
				{
					HTKKFactory.GetInstanceHTKKHeader().SetPropertyHeader(ref WinSettings.HTKKHeaderInfo, WinSettings.HTKKHeaderInfo.mst);
				}
				frmMessage frmMessage3 = new frmMessage("0002", "3", "1", "Thông báo");
				((Form)frmMessage3).ShowDialog();
				isDataChanged = false;
				isDataValid = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage3 = new frmMessage("0003", "3", "1", "Cảnh báo");
			((Form)frmMessage3).ShowDialog();
		}
		finally
		{
			winFpSpread.ResumeLayout();
		}
	}

	private void cmdExport_Click(object sender, EventArgs e)
	{
		//IL_0590: Unknown result type (might be due to invalid IL or missing references)
		//IL_0504: Unknown result type (might be due to invalid IL or missing references)
		//IL_0606: Unknown result type (might be due to invalid IL or missing references)
		//IL_06f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_06a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_072f: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			WinDoWorkEvent.CellChangeBeforClickTab(winFpSpread, winFpSpread.ActiveSheetIndex, cellActiveWhileTab.Row.Index + 1, cellActiveWhileTab.Column.Index + 1);
			bool flag = true;
			bool flag2 = false;
			string text = "#B3#B4#B5#A2#A5#A6#A7#C1#C2#C3#C4#C5#";
			if (text.Contains("#" + WinSettings.MenuID + "#"))
			{
				flag2 = true;
			}
			if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && !isClickTongHopKHBS && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && !File.Exists(WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[0].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML)))
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
			}
			else if (isDataChangedCallTHKhbs && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
				isDataChangedCallTHKhbs = false;
			}
			string accessibleDefaultActionDescription = ((Control)winFpSpread).AccessibleDefaultActionDescription;
			((Control)winFpSpread).AccessibleDefaultActionDescription = "NotCheck";
			winFpSpread.SuspendLayout();
			if (!isDataValid || isDataChanged || flag2)
			{
				WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
				WinSettings.ToKhaiInfo.TaiDuLieuLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuLon");
				WinDoWorkEvent.WinDoWorkKiemTra(winFpSpread);
				if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
				{
					Validate_01_KHBS();
					if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].isValidData)
					{
						flag = false;
						winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].BackColor = Color.Pink;
					}
					if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].isValidData)
					{
						flag = false;
						winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].BackColor = Color.Pink;
					}
				}
			}
			else if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				Validate_01_KHBS();
				if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].isValidData)
				{
					flag = false;
					winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].BackColor = Color.Pink;
				}
				if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].isValidData)
				{
					flag = false;
					winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].BackColor = Color.Pink;
				}
			}
			winFpSpread.ResumeLayout();
			for (int i = 0; i < winFpSpread.Sheets.Count; i++)
			{
				if (winFpSpread.TabStrip[i].BackColor == Color.Pink)
				{
					flag = false;
					break;
				}
			}
			if ((WinSettings.MenuID.Equals("78") || WinSettings.MenuID.Equals("A1") || WinSettings.MenuID.Equals("U9")) && ((WinSettings.ToKhaiInfo.SheetInfos[2] == null && WinSettings.ToKhaiInfo.SheetInfos[3] == null) || (WinSettings.ToKhaiInfo.SheetInfos[2] != null && WinSettings.ToKhaiInfo.SheetInfos[3] != null)))
			{
				frmMessage frmMessage = new frmMessage("0478", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("C1") && (WinSettings.ToKhaiInfo.SheetInfos[1] == null || (WinSettings.ToKhaiInfo.SheetInfos[2] == null && WinSettings.ToKhaiInfo.SheetInfos[3] == null)))
			{
				frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("C2") && (WinSettings.ToKhaiInfo.SheetInfos[1] == null || WinSettings.ToKhaiInfo.SheetInfos[2] == null))
			{
				frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if ((WinSettings.MenuID.Equals("C3") || WinSettings.MenuID.Equals("C4") || WinSettings.MenuID.Equals("C5")) && (WinSettings.ToKhaiInfo.SheetInfos[1] == null || WinSettings.ToKhaiInfo.SheetInfos[2] == null))
			{
				frmMessage frmMessage = new frmMessage("0499", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID == "35" && !flag)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage2 = new frmMessage("0039", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (!flag)
			{
				winFpSpread.ResumeLayout();
				frmMessage frmMessage3 = new frmMessage("0016", "3", "3", "Thông báo");
				((Form)frmMessage3).ShowDialog();
			}
			else
			{
				frmWinSelectExport form = new frmWinSelectExport(winFpSpread);
				FormProcess.AddControl((Form)(object)form, (Control)(object)WinSettings.myWinToKhai.panelChild);
				((ContainerControl)this).ActiveControl = (Control)(object)WinSettings.myWinToKhai.panelChild;
			}
			((Control)winFpSpread).AccessibleDefaultActionDescription = accessibleDefaultActionDescription;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cmdLoadToKhai_Click(object sender, EventArgs e)
	{
		//IL_102f: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Expected O, but got Unknown
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Invalid comparison between Unknown and I4
		//IL_0097: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a4: Unknown result type (might be due to invalid IL or missing references)
		//IL_031f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0290: Unknown result type (might be due to invalid IL or missing references)
		//IL_04ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_040b: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0926: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b88: Unknown result type (might be due to invalid IL or missing references)
		//IL_05c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_0d2c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0826: Unknown result type (might be due to invalid IL or missing references)
		//IL_0636: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a88: Unknown result type (might be due to invalid IL or missing references)
		//IL_0898: Unknown result type (might be due to invalid IL or missing references)
		//IL_0c9e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0afa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fe5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0fc0: Unknown result type (might be due to invalid IL or missing references)
		//IL_0ebc: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e18: Unknown result type (might be due to invalid IL or missing references)
		//IL_0f52: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			winFpSpread.SuspendLayout();
			OpenFileDialog val = new OpenFileDialog();
			((FileDialog)val).DefaultExt = "xls";
			((FileDialog)val).Filter = "Excel file (*.xls, *.xlsx)|*.xls; *.xlsx";
			((FileDialog)val).AddExtension = true;
			((FileDialog)val).RestoreDirectory = true;
			((FileDialog)val).Title = "Chọn file tải lên?";
			DialogResult val2 = ((CommonDialog)val).ShowDialog();
			if ((int)val2 == 1)
			{
				FileInfo file = new FileInfo(((FileDialog)val).FileName);
				if (HTKKFile.IsFileLocked(file))
				{
					frmMessage frmMessage = new frmMessage("0105", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				FpSpread fpSpread = new FpSpread();
				if (fpSpread.OpenExcel(((FileDialog)val).FileName, ExcelOpenFlags.DoNotRecalculateAfterLoad))
				{
					string text = "";
					string text2 = "";
					string text3 = "";
					string text4 = "";
					string text5 = "";
					string text6 = "";
					string text7 = "";
					string text8 = "";
					string text9 = "";
					string text10 = "";
					if (WinSettings.MenuID == "33")
					{
						text = fpSpread.Sheets[0].Cells[9, HTKKFile.ColumnLettersToIndex("D") - 1].Text;
						if (text.Trim().Replace("-", "") != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text3 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("H") - 1].Text;
						text6 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("L") - 1].Text;
						text7 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("N") - 1].Text;
						if (text3 != WinSettings.TTinChung.kyKKhai || text6 != WinSettings.TTinChung.kyKKhaiTuThang || text7 != WinSettings.TTinChung.kyKKhaiDenThang)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "37")
					{
						text = fpSpread.Sheets[0].Cells[9, HTKKFile.ColumnLettersToIndex("F") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text3 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("J") - 1].Text;
						text6 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("N") - 1].Text;
						text7 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("P") - 1].Text;
						if (text3 != WinSettings.TTinChung.kyKKhai || text6 != WinSettings.TTinChung.kyKKhaiTuThang || text7 != WinSettings.TTinChung.kyKKhaiDenThang)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "78" || WinSettings.MenuID.Equals("U9"))
					{
						text = fpSpread.Sheets[0].Cells[6, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text3 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text4 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text5 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						if (text3 != WinSettings.TTinChung.kyKKhai || text4 != WinSettings.TTinChung.kyKKhaiTuNgay || text5 != WinSettings.TTinChung.kyKKhaiDenNgay)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text10 = fpSpread.Sheets[0].Cells[39, HTKKFile.ColumnLettersToIndex("F") - 1].Text;
						if (text10.Trim() != WinSettings.TTPeriodInfo.DonViTinhMa)
						{
							frmMessage frmMessage = new frmMessage("0452", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "A1")
					{
						text = fpSpread.Sheets[0].Cells[6, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text8 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text9 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text4 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text5 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						if (text8 != WinSettings.TTinChung.quy || text9 != WinSettings.TTinChung.nam || text4 != WinSettings.TTinChung.kyKKhaiTuNgay || text5 != WinSettings.TTinChung.kyKKhaiDenNgay)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text10 = fpSpread.Sheets[0].Cells[39, HTKKFile.ColumnLettersToIndex("F") - 1].Text;
						if (text10.Trim() != WinSettings.TTPeriodInfo.DonViTinhMa)
						{
							frmMessage frmMessage = new frmMessage("0452", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "A3")
					{
						text = fpSpread.Sheets[0].Cells[6, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text8 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[46, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text9 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text4 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text5 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						if (text8 != WinSettings.TTinChung.quy || text9 != WinSettings.TTinChung.nam || text4 != WinSettings.TTinChung.kyKKhaiTuNgay || text5 != WinSettings.TTinChung.kyKKhaiDenNgay)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text10 = fpSpread.Sheets[0].Cells[39, HTKKFile.ColumnLettersToIndex("F") - 1].Text;
						if (text10.Trim() != WinSettings.TTPeriodInfo.DonViTinhMa)
						{
							frmMessage frmMessage = new frmMessage("0452", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "53")
					{
						text = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text3 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[47, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text4 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[31, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						text5 = fpSpread.Sheets[fpSpread.Sheets.Count - 1].Cells[32, HTKKFile.ColumnLettersToIndex("V") - 1].Text;
						if (text3 != WinSettings.TTinChung.kyKKhai || text4 != WinSettings.TTinChung.kyKKhaiTuNgay || text5 != WinSettings.TTinChung.kyKKhaiDenNgay)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "39")
					{
						text = fpSpread.Sheets[0].Cells[8, HTKKFile.ColumnLettersToIndex("D") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text3 = fpSpread.Sheets[0].Cells[4, HTKKFile.ColumnLettersToIndex("E") - 1].Text;
						text6 = fpSpread.Sheets[0].Cells[4, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
						text7 = fpSpread.Sheets[0].Cells[4, HTKKFile.ColumnLettersToIndex("I") - 1].Text;
						if (text3 != WinSettings.TTinChung.kyKKhai || text6 != WinSettings.TTinChung.kyKKhaiTuThang || text7 != WinSettings.TTinChung.kyKKhaiDenThang)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					if (WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9")
					{
						text = fpSpread.Sheets[0].Cells[11, HTKKFile.ColumnLettersToIndex("G") - 1].Text;
						if (text.Trim() != WinSettings.TTinChung.mst)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
						text3 = fpSpread.Sheets[0].Cells[3, HTKKFile.ColumnLettersToIndex("M") - 1].Text;
						if (text3 != WinSettings.TTinChung.kyKKhai || text6 != WinSettings.TTinChung.kyKKhaiTuThang || text7 != WinSettings.TTinChung.kyKKhaiDenThang)
						{
							frmMessage frmMessage = new frmMessage("0438", "3", "3", "Thông báo");
							((Form)frmMessage).ShowDialog();
							return;
						}
					}
					isDataChanged = true;
					isDataChangedCallTHKhbs = true;
					if (WinSpreadProcess.setValueSpreadImpToForm(winFpSpread, fpSpread, WinSettings.ToKhaiInfo.SheetInfos))
					{
						frmToKhai_ReLoad();
						WinDoWorkEvent.FinishImportExcel(winFpSpread, WinSettings.ToKhaiInfo);
						frmMessage frmMessage = new frmMessage("0432", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
					}
					else
					{
						frmMessage frmMessage = new frmMessage("0433", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
					}
				}
			}
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			frmMessage frmMessage2 = new frmMessage("0106", "3", "3", "Thông báo");
			((Form)frmMessage2).ShowDialog();
		}
	}

	private void cmdDelete_Click(object sender, EventArgs e)
	{
		//IL_0087: Unknown result type (might be due to invalid IL or missing references)
		//IL_02ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_0439: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0348: Unknown result type (might be due to invalid IL or missing references)
		//IL_05ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_0508: Unknown result type (might be due to invalid IL or missing references)
		//IL_0125: Unknown result type (might be due to invalid IL or missing references)
		//IL_0602: Unknown result type (might be due to invalid IL or missing references)
		//IL_06da: Unknown result type (might be due to invalid IL or missing references)
		//IL_0732: Unknown result type (might be due to invalid IL or missing references)
		//IL_0682: Unknown result type (might be due to invalid IL or missing references)
		//IL_085e: Unknown result type (might be due to invalid IL or missing references)
		//IL_079c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0806: Unknown result type (might be due to invalid IL or missing references)
		//IL_08cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0923: Unknown result type (might be due to invalid IL or missing references)
		//IL_09b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b5b: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			int num = -1;
			string text = "";
			if (WinSettings.MenuID == "99")
			{
				string tieuMucHachToanVal = WinSettings.TTPeriodInfo.TieuMucHachToanVal;
				if (tieuMucHachToanVal == "0002" && (winFpSpread.ActiveSheetIndex == 1 || winFpSpread.ActiveSheetIndex == 2))
				{
					frmMessage frmMessage = new frmMessage("0665", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				if (tieuMucHachToanVal == "0007" && winFpSpread.ActiveSheetIndex == 5)
				{
					frmMessage frmMessage = new frmMessage("0663", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				if (tieuMucHachToanVal == "0006" && winFpSpread.ActiveSheetIndex == 3)
				{
					frmMessage frmMessage = new frmMessage("0664", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
			}
			if (WinSettings.MenuID == "G5")
			{
				int[] array = GeneralFunctions.ArrayCellId("B_6");
				int[] array2 = GeneralFunctions.ArrayCellId("B_8");
				int[] array3 = GeneralFunctions.ArrayCellId("B_10");
				int[] array4 = GeneralFunctions.ArrayCellId("B_13");
				int[] array5 = GeneralFunctions.ArrayCellId("B_15");
				int[] array6 = GeneralFunctions.ArrayCellId("B_17");
				bool flag = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array[1], array[0]].Value);
				bool flag2 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array2[1], array2[0]].Value);
				bool flag3 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array3[1], array3[0]].Value);
				bool flag4 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array4[1], array4[0]].Value);
				bool flag5 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array5[1], array5[0]].Value);
				bool flag6 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array6[1], array6[0]].Value);
				if (flag3 && winFpSpread.ActiveSheetIndex == 1)
				{
					frmMessage frmMessage2 = new frmMessage("0562", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
				if (!flag && !flag2 && !flag3 && flag4 && !flag5 && !flag6 && winFpSpread.ActiveSheetIndex == 2)
				{
					frmMessage frmMessage2 = new frmMessage("0563", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if (WinSettings.MenuID == "G7")
			{
				int[] array7 = GeneralFunctions.ArrayCellId("Q_10");
				int[] array8 = GeneralFunctions.ArrayCellId("Q_12");
				bool flag7 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array7[1], array7[0]].Value);
				bool flag8 = Convert.ToBoolean(winFpSpread.Sheets[0].Cells[array8[1], array8[0]].Value);
				if ((winFpSpread.ActiveSheetIndex == 1 && flag7) || (flag8 && winFpSpread.ActiveSheetIndex == 2))
				{
					frmMessage frmMessage2 = new frmMessage("0537", "3", "1", "Thông báo");
					((Form)frmMessage2).ShowDialog();
					return;
				}
			}
			if ((WinSettings.MenuID == "31" || WinSettings.MenuID == "G8") && winFpSpread.ActiveSheetIndex == 1 && ((WinSettings.TTPeriodInfo.KieuKy == "Q" && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Quy) == 4) || (WinSettings.TTPeriodInfo.KieuKy == "M" && ConvertHelpers.ConvertToInt(WinSettings.TTPeriodInfo.Thang) == 12)))
			{
				frmMessage frmMessage2 = new frmMessage("0368", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("78") || WinSettings.MenuID.Equals("53") || WinSettings.MenuID.Equals("A1") || WinSettings.MenuID.Equals("A3") || WinSettings.MenuID.Equals("U9"))
			{
				if (winFpSpread.ActiveSheetIndex == 1)
				{
					frmMessage frmMessage = new frmMessage("0475", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
				if (WinSettings.MenuID.Equals("A3") && winFpSpread.ActiveSheetIndex == 2)
				{
					frmMessage frmMessage = new frmMessage("0479", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					return;
				}
			}
			if ((WinSettings.MenuID.Equals("A2") || WinSettings.MenuID.Equals("A5")) && (winFpSpread.ActiveSheetIndex == 1 || winFpSpread.ActiveSheetIndex == 4))
			{
				frmMessage frmMessage = new frmMessage("0485", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("A6") && winFpSpread.ActiveSheetIndex == 1)
			{
				frmMessage frmMessage = new frmMessage("0486", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("A7") && winFpSpread.ActiveSheetIndex == 1)
			{
				frmMessage frmMessage = new frmMessage("0484", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("B3") && (winFpSpread.ActiveSheetIndex == 1 || winFpSpread.ActiveSheetIndex == 4))
			{
				frmMessage frmMessage = new frmMessage("0491", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("B5") && (winFpSpread.ActiveSheetIndex == 1 || winFpSpread.ActiveSheetIndex == 4))
			{
				frmMessage frmMessage = new frmMessage("0493", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("B4") && winFpSpread.ActiveSheetIndex == 1)
			{
				frmMessage frmMessage = new frmMessage("0492", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if ((WinSettings.MenuID.Equals("C1") || WinSettings.MenuID.Equals("C2")) && winFpSpread.ActiveSheetIndex == 1)
			{
				frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("C2") && winFpSpread.ActiveSheetIndex == 2)
			{
				frmMessage frmMessage = new frmMessage("0498", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			if ((WinSettings.MenuID.Equals("C3") || WinSettings.MenuID.Equals("C4") || WinSettings.MenuID.Equals("C5")) && (winFpSpread.ActiveSheetIndex == 1 || winFpSpread.ActiveSheetIndex == 2))
			{
				frmMessage frmMessage = new frmMessage("0499", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				return;
			}
			string text2 = "";
			text2 = ((winFpSpread.ActiveSheetIndex != 0 && !(WinSettings.MenuID == "33") && !(WinSettings.MenuID == "B9") && !(WinSettings.MenuID == "P2") && !(WinSettings.MenuID == "B8") && !(WinSettings.MenuID == "37") && !(WinSettings.MenuID == "N7") && ((!(WinSettings.MenuID == "T8") && !(WinSettings.MenuID == "R5")) || winFpSpread.ActiveSheetIndex != 3) && !WinSettings.MenuID.Equals("W1") && (!(WinSettings.TTPeriodInfo.LoaiToKhai == "B") || "30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) || !(WinSettings.ToKhaiInfo.KieuKHBS != "3")) && !(winFpSpread.ActiveSheet.SheetName == "KHBS") && !(winFpSpread.ActiveSheet.SheetName == "01-KHBS") && !(winFpSpread.ActiveSheet.SheetName == "01-1/KHBS")) ? "0032" : "0012");
			frmMessage frmMessage3 = new frmMessage(text2, "5", "2", "Thông báo");
			((Form)frmMessage3).ShowDialog();
			if (!frmMessage3.Yes)
			{
				return;
			}
			num = winFpSpread.ActiveSheetIndex;
			text = winFpSpread.ActiveSheet.SheetName;
			WinDoWorkEvent.WinDoWorkDelete(winFpSpread, WinSettings.DicSheet, winFpSpread.ActiveSheetIndex);
			frmHTKKSystems frmHTKKSystems2 = new frmHTKKSystems();
			if (winFpSpread.ActiveSheetIndex == 0 || WinSettings.MenuID == "33" || WinSettings.MenuID == "B9" || WinSettings.MenuID == "P2" || WinSettings.MenuID == "B8" || WinSettings.MenuID == "37" || WinSettings.MenuID == "N7" || ((WinSettings.MenuID == "T8" || WinSettings.MenuID == "R5") && winFpSpread.ActiveSheetIndex == 3) || WinSettings.MenuID.Equals("W1") || (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && !"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID) && WinSettings.ToKhaiInfo.KieuKHBS != "3") || winFpSpread.ActiveSheet.SheetName == "KHBS" || winFpSpread.ActiveSheet.SheetName == "01-KHBS" || winFpSpread.ActiveSheet.SheetName == "01-1/KHBS")
			{
				isDeleteEvent = true;
				((Form)this).Close();
			}
			else
			{
				winFpSpread.ActiveSheetIndex = 0;
				for (int i = 1; i < winFpSpread.Sheets.Count; i++)
				{
					try
					{
						if (WinSettings.ToKhaiInfo.SheetInfos[i] == null)
						{
							winFpSpread.Sheets[i].Visible = false;
							winFpSpread.TabStrip[i].BackColor = Color.White;
						}
					}
					catch (Exception)
					{
					}
				}
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3" && !(text == "01-KHBS") && !(text == "01-1/KHBS") && num != 0)
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
				isDataChangedCallTHKhbs = false;
				isDataChanged = true;
			}
			isDataChanged = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cmdExit_Click(object sender, EventArgs e)
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (isDataChanged && !WinSettings.MenuID.Equals("57"))
			{
				frmMessage frmMessage = new frmMessage("0013", "6", "2", "Cảnh báo");
				((Form)frmMessage).ShowDialog();
				if (frmMessage.Yes)
				{
					cmdSave_Click(sender, e);
					isDataChanged = false;
					isDataChangedCallTHKhbs = false;
				}
				else
				{
					if (frmMessage.Cancel)
					{
						((Control)frmMessage).Hide();
						return;
					}
					isDataChanged = false;
					isDataChangedCallTHKhbs = false;
				}
			}
			((Form)this).Close();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cmdKiemTra_Click(object sender, EventArgs e)
	{
		//IL_0194: Unknown result type (might be due to invalid IL or missing references)
		//IL_0165: Unknown result type (might be due to invalid IL or missing references)
		//IL_013e: Unknown result type (might be due to invalid IL or missing references)
		bool flag = true;
		try
		{
			winFpSpread.SuspendLayout();
			if (!isDataValid || isDataChanged)
			{
				WinDoWorkEvent.WinDoWorkKiemTra(winFpSpread);
				Validate_01_KHBS();
			}
			for (int i = 0; i < winFpSpread.Sheets.Count; i++)
			{
				if (winFpSpread.TabStrip[i].BackColor == Color.Pink)
				{
					flag = false;
					break;
				}
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS != "" && !WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].isValidData)
			{
				flag = false;
			}
			if (!flag)
			{
				string text = HTKKFile.readLogErrorBigData();
				isDataValid = false;
				if (showErrorBigData() && !HTKKFormater.isNullString(text))
				{
					frmMessageForm frmMessageForm = new frmMessageForm("0158", "3", "Chi tiết lỗi", text);
					((Form)frmMessageForm).ShowDialog();
				}
				else
				{
					frmMessage frmMessage = new frmMessage("0158", "3", "1", "Thông báo");
					((Form)frmMessage).ShowDialog();
				}
			}
			else
			{
				isDataValid = true;
				frmMessage frmMessage = new frmMessage("0157", "3", "1", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool showErrorBigData()
	{
		bool flag = false;
		try
		{
			if (WinSettings.MenuID == "31" && (WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[1].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[2].RowsValues.Count > 500))
			{
				flag = true;
			}
			else if (WinSettings.MenuID == "G8" && (WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[1].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[2].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[3].RowsValues.Count > 500))
			{
				flag = true;
			}
			else if (WinSettings.MenuID == "33" && (WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[3].Tables[0].RowsValues.Count > 500))
			{
				flag = true;
			}
			else if (WinSettings.MenuID == "B9" || WinSettings.MenuID == "N7" || WinSettings.MenuID == "L8" || WinSettings.MenuID == "Y6" || WinSettings.MenuID.Equals("W1"))
			{
				flag = true;
			}
			else if (WinSettings.MenuID == "M4")
			{
				int num = 0;
				foreach (HTKKCMCTableInfo table in WinSettings.ToKhaiInfo.SheetInfos[0].Tables)
				{
					num += table.RowsValues.Count;
				}
				if (num > 500)
				{
					flag = true;
				}
			}
			else if ((WinSettings.MenuID == "P1" || WinSettings.MenuID.Equals("Y9")) && WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count > 500)
			{
				flag = true;
			}
			else if ((WinSettings.MenuID == "P6" || WinSettings.MenuID == "X7") && (WinSettings.ToKhaiInfo.SheetInfos[0].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[0].Tables[1].RowsValues.Count > 500))
			{
				flag = true;
			}
			else if (WinSettings.MenuID == "99" && (WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[2].Tables[0].RowsValues.Count > 500))
			{
				flag = true;
			}
			else if (WinSettings.MenuID == "P2" && (WinSettings.ToKhaiInfo.SheetInfos[1].Tables[0].RowsValues.Count > 500 || WinSettings.ToKhaiInfo.SheetInfos[1].Tables[1].RowsValues.Count > 500))
			{
				flag = true;
			}
			else if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
			{
				flag = true;
			}
			else if (WinSettings.ToKhaiInfo.TaiDuLieuLon == "1")
			{
				flag = true;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return flag;
	}

	private void btnThemPhuLuc_Click(object sender, EventArgs e)
	{
		winFpSpread.SuspendLayout();
		frmWinSelectPhuLuc form = new frmWinSelectPhuLuc(winFpSpread);
		isDataChanged = true;
		isDataChangedCallTHKhbs = true;
		FormProcess.AddControl((Form)(object)form, (Control)(object)WinSettings.myWinToKhai.panelChild);
		winFpSpread.ResumeLayout();
	}

	private void cmdImportXML_Click(object sender, EventArgs e)
	{
		//IL_008a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected O, but got Unknown
		//IL_00c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cf: Invalid comparison between Unknown and I4
		//IL_01e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_0412: Unknown result type (might be due to invalid IL or missing references)
		//IL_0358: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			List<HTKKCMCSheetInfo> list = new List<HTKKCMCSheetInfo>();
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[i] != null)
				{
					HTKKCMCSheetInfo item = new HTKKCMCSheetInfo(WinSettings.ToKhaiInfo.SheetInfos[i]);
					list.Add(item);
				}
				else
				{
					HTKKCMCSheetInfo item = new HTKKCMCSheetInfo();
					list.Add(item);
					list[i] = null;
				}
			}
			winFpSpread.SuspendLayout();
			OpenFileDialog val = new OpenFileDialog();
			((FileDialog)val).DefaultExt = "xml";
			((FileDialog)val).Filter = "XML file (*.xml)|*.xml";
			((FileDialog)val).AddExtension = true;
			((FileDialog)val).RestoreDirectory = true;
			((FileDialog)val).Title = "Chọn tờ khai để tải vào chương trình";
			DialogResult val2 = ((CommonDialog)val).ShowDialog();
			if ((int)val2 == 1)
			{
				FileNameLoading = ((FileDialog)val).FileName;
				WinSettings.ToKhaiInfo.RemoveAllSheet();
				WinFnProcess.GetHTKKSheetInfos(WinSettings.ToKhaiInfo, getTemplate: true, WinSettings.TTPeriodInfo);
				WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
				WinSettings.ToKhaiInfo.TaiDuLieuLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuLon");
				if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
				{
					WinSettings.ToKhaiInfo.SheetInfos = WinDoWorkEvent.WinDoWorkImportXmlToSheetInfos(FileNameLoading);
					WinSettings.isReadDataXML = true;
					WinDoWorkEvent.FinishImportXML(winFpSpread, WinSettings.ToKhaiInfo);
					if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
					{
						RevertFromXml();
					}
					frmMessage frmMessage = new frmMessage("0291", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					isDataChanged = true;
					isDataChangedCallTHKhbs = true;
					WinSettings.isReadDataXML = false;
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if (hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.Tables.Count > 0 && hTKKCMCSheetInfo.TaiDuLieuRatLon == "1")
					{
						((Control)panelPage).Visible = true;
					}
					else
					{
						((Control)panelPage).Visible = false;
					}
					return;
				}
				if (WinDoWorkEvent.WinDoWorkImportFormXml(winFpSpread, FileNameLoading))
				{
					if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon != "1")
					{
						string text = HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\log_valid.txt";
						if (File.Exists(text))
						{
							HTKKFile.DeleteFile(text);
						}
						if (!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) && !(!string.IsNullOrEmpty(WinSettings.ToKhaiInfo.XsdFileName) & WinDoWorkEvent.ValidateXmlAgainstXsd(FileNameLoading, WinSettings.ToKhaiInfo.XsdFileName)))
						{
							winFpSpread.ResumeLayout();
							frmMessage frmMessage2 = new frmMessage("0278", "3", "3", "Thông báo");
							((Form)frmMessage2).ShowDialog();
							return;
						}
					}
					WinSettings.isReadDataXML = true;
					frmToKhai_ReLoad();
					if ("~B8~37~33~B9~N7~W1".IndexOf("~" + WinSettings.MenuID + "~") > -1)
					{
						AddFullPLImportXML();
					}
					WinDoWorkEvent.FinishImportXML(winFpSpread, WinSettings.ToKhaiInfo);
					if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
					{
						RevertFromXml();
					}
					frmMessage frmMessage = new frmMessage("0291", "3", "", "Thông báo");
					((Form)frmMessage).ShowDialog();
					isDataChanged = true;
					isDataChangedCallTHKhbs = true;
					WinSettings.isReadDataXML = false;
				}
				else
				{
					WinSettings.ToKhaiInfo.SheetInfos = list;
					for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
					{
						if (list[i] != null)
						{
							WinSettings.ToKhaiInfo.SheetInfos[i] = new HTKKCMCSheetInfo(list[i]);
						}
					}
				}
			}
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos.Count; i++)
			{
				GeneralFunctions.fpSpReindexingSheetTemplate(winFpSpread.Sheets[i], WinSettings.ToKhaiInfo.SheetInfos[i]);
			}
			winFpSpread.ResumeLayout();
		}
		catch (Exception ex)
		{
			winFpSpread.ResumeLayout();
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void InitTimKiem(int sheetIndex)
	{
		//IL_05ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_05d5: Expected O, but got Unknown
		//IL_0612: Unknown result type (might be due to invalid IL or missing references)
		//IL_0619: Expected O, but got Unknown
		//IL_06b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_06bd: Expected O, but got Unknown
		cbbTimKiem.Items.Clear();
		cbbTimKiem.DataSource = null;
		dicTimKiemValue = new Dictionary<string, int>();
		int value = 0;
		int value2 = 0;
		int value3 = 0;
		int num = 0;
		int num2 = 0;
		if (WinSettings.MenuID == "33")
		{
			switch (sheetIndex)
			{
			case 1:
				value = 4;
				value2 = 5;
				value3 = 6;
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			case 2:
				value = 3;
				value2 = 4;
				value3 = 5;
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			case 3:
				value = 4;
				value2 = 5;
				value3 = 6;
				num = 8;
				num2 = 11;
				dicTimKiemValue.Add("Họ và tên NNT", value);
				dicTimKiemValue.Add("Mã số thuế NNT", value2);
				dicTimKiemValue.Add("Họ và tên NPT", value3);
				dicTimKiemValue.Add("Mã số thuế NPT", num);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu NNT", num2);
				break;
			default:
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			}
		}
		else if (WinSettings.MenuID == "B9")
		{
			switch (sheetIndex)
			{
			case 1:
				value = 4;
				value2 = 5;
				value3 = 8;
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/CCCD/SĐDCN Hộ chiếu", value3);
				break;
			case 2:
				value = 4;
				value2 = 5;
				value3 = 8;
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			case 3:
				value = 3;
				value2 = 5;
				value3 = 7;
				num = 11;
				num2 = 15;
				dicTimKiemValue.Add("Họ và tên NNT", value);
				dicTimKiemValue.Add("Mã số thuế NNT", value2);
				dicTimKiemValue.Add("Họ và tên NPT", value3);
				dicTimKiemValue.Add("Mã số thuế NPT", num);
				dicTimKiemValue.Add("Số giấy tờ", num2);
				break;
			default:
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			}
		}
		else if (WinSettings.MenuID == "W1")
		{
			switch (sheetIndex)
			{
			case 1:
				value = 4;
				value2 = 5;
				value3 = 8;
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/CCCD/SĐDCN Hộ chiếu", value3);
				break;
			case 2:
				value = 4;
				value2 = 5;
				value3 = 8;
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			case 3:
				value = 3;
				value2 = 5;
				value3 = 7;
				num = 11;
				num2 = 15;
				dicTimKiemValue.Add("Họ và tên NNT", value);
				dicTimKiemValue.Add("Mã số thuế NNT", value2);
				dicTimKiemValue.Add("Họ và tên NPT", value3);
				dicTimKiemValue.Add("Mã số thuế NPT", num);
				dicTimKiemValue.Add("Số giấy tờ", num2);
				break;
			default:
				dicTimKiemValue.Add("Họ và tên", value);
				dicTimKiemValue.Add("Mã số thuế", value2);
				dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
				break;
			}
		}
		else if (WinSettings.MenuID == "G2")
		{
			value = HTKKFile.ColumnLettersToIndex("F");
			value2 = HTKKFile.ColumnLettersToIndex("I");
			value3 = HTKKFile.ColumnLettersToIndex("O");
			num = HTKKFile.ColumnLettersToIndex("S");
			dicTimKiemValue.Add("Mã DAĐT", value);
			dicTimKiemValue.Add("Tên dự án đầu tư", value2);
			dicTimKiemValue.Add("MST chủ dự án đầu tư", value3);
			dicTimKiemValue.Add("Số văn bản ban hành DAĐT", num);
		}
		else
		{
			value = 4;
			value2 = 5;
			value3 = 6;
			dicTimKiemValue.Add("Họ và tên", value);
			dicTimKiemValue.Add("Mã số thuế", value2);
			dicTimKiemValue.Add("Số CMND/Hộ chiếu", value3);
		}
		cbbTimKiem.Items.Clear();
		foreach (KeyValuePair<string, int> item in dicTimKiemValue)
		{
			cbbTimKiem.Items.Add((object)item.Key);
		}
		((ListControl)cbbTimKiem).SelectedIndex = 0;
		if (!WinSettings.MenuID.Equals("G2"))
		{
			return;
		}
		cbbCQTCuc.Items.Clear();
		cbbCQTChiCuc.Items.Clear();
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		HTKKFile.AddOrUpdateDicStr(dictionary, "", "");
		string text = HTKKConfig.INTERFACE_INI + "\\Catalogue_Tinh_Thanh.xml";
		XmlDocument val = new XmlDocument();
		if (!File.Exists(text))
		{
			return;
		}
		val.Load(text);
		XmlNodeList elementsByTagName = val.GetElementsByTagName("Item");
		foreach (XmlNode item2 in elementsByTagName)
		{
			XmlNode val2 = item2;
			string value4 = ((XmlNode)val2.Attributes["Value"]).Value;
			string[] array = Regex.Split(value4, "###");
			if (array.Length > 3 && array[0] == "1")
			{
				string key = array[3];
				string newValue = array[1];
				HTKKFile.AddOrUpdateDicStr(dictionary, key, newValue);
			}
		}
		cbbCQTCuc.DataSource = (object)new BindingSource((object)dictionary, (string)null);
		((ListControl)cbbCQTCuc).DisplayMember = "Key";
		((ListControl)cbbCQTCuc).ValueMember = "Value";
		((Control)cbbCQTCuc).Refresh();
	}

	private void cmdPrint_Click(object sender, EventArgs e)
	{
		//IL_0419: Unknown result type (might be due to invalid IL or missing references)
		//IL_056e: Unknown result type (might be due to invalid IL or missing references)
		//IL_04e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_05e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_067f: Unknown result type (might be due to invalid IL or missing references)
		//IL_071b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0751: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			WinSettings.isActionPrint = true;
			bool flag = true;
			bool flag2 = false;
			string text = "#B3#B4#B5#A2#A5#A6#A7#C1#C2#C3#C4#C5#";
			if (text.Contains("#" + WinSettings.MenuID + "#"))
			{
				flag2 = true;
			}
			if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && !isClickTongHopKHBS && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && !File.Exists(WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[0].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML)))
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
			}
			else if (isDataChangedCallTHKhbs && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				isClickTongHopSNNC = false;
				btnTongHopKHBS_Click(sender, e);
				isDataChangedCallTHKhbs = false;
			}
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				SetThongTinChungNguoiKyNgayKyKHBS();
			}
			string accessibleDefaultActionDescription = ((Control)winFpSpread).AccessibleDefaultActionDescription;
			((Control)winFpSpread).AccessibleDefaultActionDescription = "NotCheck";
			if (!isDataValid || isDataChanged || flag2)
			{
				WinDoWorkEvent.WinDoWorkKiemTra(winFpSpread);
				if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
				{
					Validate_01_KHBS();
					if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].isValidData)
					{
						flag = false;
						winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].BackColor = Color.Pink;
					}
					if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].isValidData)
					{
						flag = false;
						winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].BackColor = Color.Pink;
					}
				}
			}
			else if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS == "3")
			{
				Validate_01_KHBS();
				if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].isValidData)
				{
					flag = false;
					winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 1].BackColor = Color.Pink;
				}
				if (!WinSettings.ToKhaiInfo.SheetInfos[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].isValidData)
				{
					flag = false;
					winFpSpread.TabStrip[WinSettings.ToKhaiInfo.SheetInfos.Count - 2].BackColor = Color.Pink;
				}
			}
			if (WinSettings.MenuID.Equals("57"))
			{
				object value = null;
				object obj = "CHECK_CHON";
				WinSettings.ToKhaiInfo.dicObject.TryGetValue("CHECK_CHON", out value);
				if (!ConvertHelpers.ConvertToBool(value))
				{
					frmMessage frmMessage = new frmMessage("0449", "3", "3", "Thông báo");
					((Form)frmMessage).ShowDialog();
					WinSettings.isActionPrint = false;
					return;
				}
			}
			if ((WinSettings.MenuID.Equals("78") || WinSettings.MenuID.Equals("A1") || WinSettings.MenuID.Equals("U9")) && ((WinSettings.ToKhaiInfo.SheetInfos[2] == null && WinSettings.ToKhaiInfo.SheetInfos[3] == null) || (WinSettings.ToKhaiInfo.SheetInfos[2] != null && WinSettings.ToKhaiInfo.SheetInfos[3] != null)))
			{
				frmMessage frmMessage2 = new frmMessage("0478", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("C1") && (WinSettings.ToKhaiInfo.SheetInfos[1] == null || (WinSettings.ToKhaiInfo.SheetInfos[2] == null && WinSettings.ToKhaiInfo.SheetInfos[3] == null)))
			{
				frmMessage frmMessage2 = new frmMessage("0498", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if (WinSettings.MenuID.Equals("C2") && (WinSettings.ToKhaiInfo.SheetInfos[1] == null || WinSettings.ToKhaiInfo.SheetInfos[2] == null))
			{
				frmMessage frmMessage2 = new frmMessage("0498", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			if ((WinSettings.MenuID.Equals("C3") || WinSettings.MenuID.Equals("C4") || WinSettings.MenuID.Equals("C5")) && (WinSettings.ToKhaiInfo.SheetInfos[1] == null || WinSettings.ToKhaiInfo.SheetInfos[2] == null))
			{
				frmMessage frmMessage2 = new frmMessage("0499", "3", "3", "Thông báo");
				((Form)frmMessage2).ShowDialog();
				return;
			}
			for (int i = 0; i < winFpSpread.Sheets.Count; i++)
			{
				if (winFpSpread.TabStrip[i].BackColor == Color.Pink)
				{
					flag = false;
					break;
				}
			}
			if (WinSettings.MenuID == "35" && !flag)
			{
				frmMessage frmMessage = new frmMessage("0005", "3", "3", "Thông báo");
				((Form)frmMessage).ShowDialog();
				WinSettings.isActionPrint = false;
				return;
			}
			if (!flag)
			{
				frmMessage frmMessage3 = new frmMessage("0005", "3", "3", "Thông báo");
				((Form)frmMessage3).ShowDialog();
			}
			else
			{
				((Control)pnWinSpread).ResumeLayout();
				WinDoWorkEvent.WinDoWorkPrint(winFpSpread, WinSettings.TTPeriodInfo);
				((ContainerControl)this).ActiveControl = (Control)(object)WinSettings.myWinToKhai.panelChild;
			}
			((Control)winFpSpread).AccessibleDefaultActionDescription = accessibleDefaultActionDescription;
			WinSettings.isActionPrint = false;
		}
		catch (Exception ex)
		{
			WinSettings.isActionPrint = false;
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cbbTimKiem_SelectedValueChanged(object sender, EventArgs e)
	{
		index = 21;
	}

	private void btnTaiBangKe_Click(object sender, EventArgs e)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		try
		{
			frmWinHTKKOpenFile frmWinHTKKOpenFile2 = new frmWinHTKKOpenFile(winFpSpread);
			((Form)frmWinHTKKOpenFile2).FormClosed += new FormClosedEventHandler(FrmopenFile_FormClosed);
			FormProcess.AddControl((Form)(object)frmWinHTKKOpenFile2, (Control)(object)WinSettings.myWinToKhai.panelChild);
			isDataChanged = true;
			isDataChangedCallTHKhbs = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void FrmopenFile_FormClosed(object sender, FormClosedEventArgs e)
	{
		if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].TaiDuLieuRatLon == "1")
		{
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
			if (hTKKCMCSheetInfo != null && hTKKCMCSheetInfo.Tables.Count > 0)
			{
				GeneralFunctions._pageSize = GeneralFunctions.tinhPageSize_Large(hTKKCMCSheetInfo, WinSettings.MenuID);
				GeneralFunctions._totalPage = GeneralFunctions.tinhTotalPage_Large(hTKKCMCSheetInfo, WinSettings.MenuID);
				((Control)labelTotalPage).Text = GeneralFunctions._totalPage.ToString();
				((Control)txtPage).Text = "1";
			}
		}
	}

	private void winFpSpread_SheetTabClick(object sender, SheetTabClickEventArgs e)
	{
		try
		{
			if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && e.SheetTabIndex == winFpSpread.Sheets.Count - 2)
			{
				((Control)btnTaiBangKe).Visible = false;
			}
			if (!(WinSettings.ToKhaiInfo.MenuID == "70"))
			{
				if (WinSettings.ToKhaiInfo.MenuID == "38")
				{
					if (e.SheetTabIndex == 0)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "G5")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3 || e.SheetTabIndex == 4 || e.SheetTabIndex == 5)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "Q8" || WinSettings.ToKhaiInfo.MenuID == "U5")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "R7")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "B9" || WinSettings.ToKhaiInfo.MenuID.Equals("W1"))
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3)
					{
						((Control)pnTimKiem).Visible = true;
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)pnTimKiem).Visible = false;
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "N7")
				{
					if (e.SheetTabIndex == 0 || e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "N2")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "Q4")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "33" || WinSettings.ToKhaiInfo.MenuID == "57")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3)
					{
						((Control)pnTimKiem).Visible = true;
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)pnTimKiem).Visible = false;
						if (e.SheetTabIndex == 3)
						{
							((Control)btnTaiBangKe).Visible = true;
						}
						else
						{
							((Control)btnTaiBangKe).Visible = false;
						}
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "31")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "G8")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "Q9")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "R4" || WinSettings.ToKhaiInfo.MenuID == "X8")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "U7")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "L4")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "79")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "T8" || WinSettings.ToKhaiInfo.MenuID == "R5")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 3)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "03")
				{
					if (e.SheetTabIndex == 12)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "G7")
				{
					if (e.SheetTabIndex == 2 || e.SheetTabIndex == 3 || e.SheetTabIndex == 4 || e.SheetTabIndex == 5)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "D1")
				{
					if (e.SheetTabIndex == 5 || e.SheetTabIndex == 6 || e.SheetTabIndex == 7)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "Z5")
				{
					if (e.SheetTabIndex == 6)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "D4")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "Z3")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "L8")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "Y6")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "C8")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "J3")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "P2")
				{
					if (e.SheetTabIndex == 1)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
				else if (WinSettings.ToKhaiInfo.MenuID == "99")
				{
					if (e.SheetTabIndex == 1 || e.SheetTabIndex == 2 || e.SheetTabIndex == 3 || e.SheetTabIndex == 4 || e.SheetTabIndex == 5)
					{
						((Control)btnTaiBangKe).Visible = true;
					}
					else
					{
						((Control)btnTaiBangKe).Visible = false;
					}
				}
			}
			DisplayF5F7(e.SheetTabIndex);
			InitTimKiem(e.SheetTabIndex);
			ArrangeHorizonPanel(pnButon);
			WinDoWorkEvent.CellChangeBeforClickTab(winFpSpread, winFpSpread.ActiveSheetIndex, cellActiveWhileTab.Row.Index + 1, cellActiveWhileTab.Column.Index + 1);
			if (e.SheetTabIndex == 2)
			{
				WinSpreadProcess.reloadBc26AcPl310(winFpSpread);
			}
			if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && WinSettings.TTPeriodInfo.LoaiToKhai == "B" && e.SheetTabIndex == WinSettings.ToKhaiInfo.SheetInfos.Count - 2)
			{
				SetHeightOfTable_01_KHBS(WinSettings.ToKhaiInfo.SheetInfos.Count - 2);
			}
			showPanelPage(e.SheetTabIndex);
			WinDoWorkEvent.SheetTabChange(winFpSpread, e.SheetTabIndex);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void showPanelPage(int sheetIndex)
	{
		if (WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].TaiDuLieuRatLon == "1")
		{
			HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			GeneralFunctions._pageSize = GeneralFunctions.tinhPageSize_Large(sInfo, WinSettings.MenuID);
			GeneralFunctions._totalPage = GeneralFunctions.tinhTotalPage_Large(sInfo, WinSettings.MenuID);
			((Control)labelTotalPage).Text = GeneralFunctions._totalPage.ToString();
			((Control)txtPage).Text = "1";
			GeneralFunctions._page = 0;
			((Control)panelPage).Visible = true;
			((Control)panelPage).Location = new Point(12, 5);
		}
		else
		{
			((Control)panelPage).Visible = false;
		}
	}

	private void DisplayF5F7(int iSheetIndex)
	{
		try
		{
			bool visible = false;
			bool visible2 = false;
			if (WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables != null)
			{
				if (WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables.Count > 0)
				{
					for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables.Count; i++)
					{
						if (WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables[i].RowsInfo.Cells != null && WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables[i].RowsInfo.Cells.Count > 0)
						{
							visible = true;
							break;
						}
					}
					((Control)pnF5).Visible = visible;
					for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables.Count; i++)
					{
						if (WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables[i].Header != null && WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables[i].Header.StaticCells.Count > 0)
						{
							visible2 = true;
							break;
						}
						if (WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables[i].Footer != null && WinSettings.ToKhaiInfo.SheetInfos[iSheetIndex].Tables[i].Footer.StaticCells.Count > 0)
						{
							visible2 = true;
							break;
						}
					}
					((Control)pnF7).Visible = visible2;
				}
				else
				{
					((Control)pnF5).Visible = false;
					((Control)pnF7).Visible = false;
				}
			}
			else
			{
				((Control)pnF5).Visible = false;
				((Control)pnF7).Visible = false;
			}
			if (((Control)pnF5).Visible && ((Control)pnF7).Visible)
			{
				((Control)pnF5).Location = new Point(0, 0);
				((Control)pnF7).Location = new Point(90, 0);
			}
			else if (((Control)pnF7).Visible)
			{
				((Control)pnF7).Location = new Point(0, 0);
			}
			else if (((Control)pnF5).Visible)
			{
				((Control)pnF5).Location = new Point(0, 0);
			}
			if (WinSettings.MenuID == "100_1" || WinSettings.MenuID == "100_5" || WinSettings.MenuID == "")
			{
				((Control)lblInfo).Text = "Ký hiệu * là bắt buộc nhập";
				((Control)lblInfo).Location = new Point(4, 5);
				((Control)lblInfo).Size = new Size(79, 13);
				((Control)lblInfo).Height = 40;
				((Control)pnF5).Controls.Clear();
				((Control)pnF5).Controls.Add((Control)(object)lblInfo);
				((Control)pnF5).Visible = true;
				((Control)pnF5).Location = new Point(0, 0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_CellChange(object sender, ChangeEventArgs e)
	{
		try
		{
			string text = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
			if (_oldValue == text)
			{
				return;
			}
			string text2 = "";
			isDataChanged = true;
			string selectedValues = string.Empty;
			HTKKCMCCellInfo cStatic = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
			if (winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 2 && winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 3)
			{
				isDataChangedCallTHKhbs = true;
			}
			if (cStatic != null)
			{
				if (cStatic.Controltype == "100")
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						GeneralFunctions.resetCqtTheoTinhStatic(winFpSpread.ActiveSheet, hTKKCMCSheetInfo.StaticCells);
					}
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower())
					{
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "8");
						if (hTKKCMCCellInfo != null && cStatic.CellTemplate.RowIndex <= hTKKCMCCellInfo.CellTemplate.RowIndex)
						{
							GeneralFunctions.resetCqtTheoTinhStatic(winFpSpread.ActiveSheet, hTKKCMCSheetInfo.StaticCells);
						}
					}
				}
				if (cStatic.Controltype == "8")
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower() || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						selectedValues = GeneralFunctions.getMaComboboxSView(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
						GeneralFunctions.setSviewCqtCapChiCucTheoMaCqtCuc_Static(winFpSpread.ActiveSheet, hTKKCMCSheetInfo, selectedValues);
						if (cStatic.CellSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[cStatic.CellSelectedValue.RowIndex - 1, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
						}
						return;
					}
					if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsKey(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text))
					{
						HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
						if (string.IsNullOrEmpty(selectedValues))
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
						}
						else
						{
							if (cStatic.CellSelectedValue != null)
							{
								winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
							}
							if (cStatic.ChildCell != null)
							{
								WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, cStatic.ChildCell.RowIndex, cStatic.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(selectedValues));
								winFpSpread.ActiveSheet.Cells[cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex - 1].Text = "";
							}
						}
					}
				}
				else if (cStatic.Controltype == "9")
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower() || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						selectedValues = GeneralFunctions.getMaComboboxSView(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
						if (cStatic.CellSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[cStatic.CellSelectedValue.RowIndex - 1, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
						}
						return;
					}
					if (cStatic.ParentCell != null && HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text))
					{
						string text3 = "";
						foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc)
						{
							if (item.Value == winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text)
							{
								text3 = item.Key;
							}
						}
						selectedValues = HTKKWinDictionaryDiaDanh.getMaCoQuanThueQuanLy(winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text, winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text);
						if (cStatic.CellSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[cStatic.CellSelectedValue.RowIndex - 1, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
						}
						if (string.IsNullOrEmpty(selectedValues))
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
						}
					}
				}
				cStatic.Value = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
				if (cStatic.Controltype == "14")
				{
					text2 = cStatic.Value;
					string text4 = WinDoWorkEvent.CheckFormatDateOnChange(text2);
					if (!HTKKFormater.isNullString(text2))
					{
						if (HTKKFormater.isNullString(text4))
						{
							isErrCellOld = true;
							cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
						}
						else
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text4;
							cStatic.Value = text4;
						}
					}
				}
			}
			else
			{
				int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
				if (tableIndex >= 0)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
					cStatic = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Header.StaticCells);
					if (cStatic != null)
					{
						cStatic.Value = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
						if (cStatic.Controltype == "72")
						{
							Dictionary<string, string> source = WinSpreadProcess.DMMo();
							KeyValuePair<string, string> current = source.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == cStatic.Value);
							if (!string.IsNullOrEmpty(current.Key))
							{
								winFpSpread.ActiveSheet.Cells[e.Row, cStatic.CellSelectedValue.ColIndex - 1].Text = current.Key;
							}
						}
						if (cStatic.Controltype == "77")
						{
							Dictionary<string, string> source = WinSpreadProcess.GetDMToKhaiPhuLucKHBS();
							KeyValuePair<string, string> current = source.FirstOrDefault((KeyValuePair<string, string> x) => x.Value == cStatic.Value);
							if (!string.IsNullOrEmpty(current.Key))
							{
								winFpSpread.ActiveSheet.Cells[e.Row, cStatic.CellSelectedValue.ColIndex - 1].Text = current.Key;
							}
						}
						if (cStatic.Controltype == "5")
						{
							selectedValues = HTKKWinDictionaryDiaDanh.dicTinhTP[winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text];
							if (cStatic.ChildCell != null)
							{
								WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, cStatic.ChildCell.RowIndex, cStatic.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[selectedValues]);
								winFpSpread.ActiveSheet.Cells[cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex - 1].Text = "";
								HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
								if (staticCell != null)
								{
									winFpSpread.ActiveSheet.Cells[staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1].Text = "";
								}
							}
						}
						if (cStatic.Controltype == "103")
						{
							selectedValues = HTKKWinDictionaryDiaDanh.dicThueTinh[winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text];
							if (cStatic.ChildCell != null)
							{
								string[] values = (from x in HTKKWinDictionaryDiaDanh.thueTinhs
									where x.Ma == selectedValues
									select x.Ten).ToArray();
								WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, cStatic.ChildCell.RowIndex, cStatic.ChildCell.ColIndex, values);
								winFpSpread.ActiveSheet.Cells[cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex - 1].Text = "";
								HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
								if (staticCell != null)
								{
									winFpSpread.ActiveSheet.Cells[staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1].Text = "";
								}
							}
						}
						if (cStatic.Controltype == "100")
						{
							selectedValues = HTKKWinDictionaryDiaDanh.dicTinhTPMoi[winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text];
							if (cStatic.ChildCell != null)
							{
								WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, cStatic.ChildCell.RowIndex, cStatic.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(selectedValues));
								winFpSpread.ActiveSheet.Cells[cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex - 1].Text = "";
								HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
								if (staticCell != null)
								{
									winFpSpread.ActiveSheet.Cells[staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1].Text = "";
								}
							}
							GeneralFunctions.resetCqtTheoTinhStatic(winFpSpread.ActiveSheet, hTKKCMCTableInfo.Header.StaticCells);
						}
						if (cStatic.Controltype == "6")
						{
							if (cStatic.ParentCell != null)
							{
								maTinhTP = winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text;
							}
							selectedValues = HTKKWinDictionaryDiaDanh.getMaQuanHuyen(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, maTinhTP);
							if (cStatic.ChildCell != null)
							{
								WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, cStatic.ChildCell.RowIndex, cStatic.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getPhuongXa(maTinhTP, winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text));
								winFpSpread.ActiveSheet.Cells[cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex - 1].Text = "";
							}
						}
						if (cStatic.Controltype == "7")
						{
							if (cStatic.ParentCell != null)
							{
								maQuanHuyen = winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text;
							}
							selectedValues = HTKKWinDictionaryDiaDanh.getMaPhuongXa(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, maQuanHuyen);
						}
						if (cStatic.Controltype == "102")
						{
							string maCQT = "";
							if (cStatic.ParentCell != null)
							{
								maCQT = winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text;
							}
							selectedValues = HTKKWinDictionaryDiaDanh.coQuanThus.FirstOrDefault((CoQuanThu x) => x.Ten == cStatic.Value && x.MaCQT == maCQT).Ma;
						}
						if (cStatic.Controltype == "101")
						{
							if (cStatic.ParentCell != null)
							{
								maTinhTP = winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text;
							}
							selectedValues = HTKKWinDictionaryDiaDanh.getMaPhuongXaMoi(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, maTinhTP);
						}
						else if (cStatic.Controltype != null && cStatic.Controltype.Length > 3 && cStatic.Controltype.Substring(0, 3) == "DC_")
						{
							foreach (KeyValuePair<string, string> item2 in WinSpreadProcess.GetDMDungChung(cStatic.Controltype))
							{
								if (winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text == item2.Value)
								{
									selectedValues = item2.Key;
									break;
								}
							}
						}
						if (cStatic.Controltype == "8")
						{
							HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
							if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
							{
								selectedValues = GeneralFunctions.getMaComboboxSView(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
								GeneralFunctions.setSviewCqtCapChiCucTheoMaCqtCuc_Static(winFpSpread.ActiveSheet, hTKKCMCSheetInfo, selectedValues);
								if (cStatic.CellSelectedValue != null)
								{
									winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
								}
							}
							else if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsKey(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text))
							{
								HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
								if (string.IsNullOrEmpty(selectedValues))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
								}
								else
								{
									if (cStatic.CellSelectedValue != null)
									{
										winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
									}
									if (cStatic.ChildCell != null)
									{
										WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, cStatic.ChildCell.RowIndex, cStatic.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(selectedValues));
										winFpSpread.ActiveSheet.Cells[cStatic.ChildCell.RowIndex - 1, cStatic.ChildCell.ColIndex - 1].Text = "";
									}
								}
							}
						}
						else if (cStatic.Controltype == "9")
						{
							HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
							if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
							{
								selectedValues = GeneralFunctions.getMaComboboxSView(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
								if (cStatic.CellSelectedValue != null)
								{
									winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
								}
							}
							else if (cStatic.ParentCell != null && HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsValue(winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text))
							{
								string text3 = "";
								foreach (KeyValuePair<string, string> item3 in HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc)
								{
									if (item3.Value == winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text)
									{
										text3 = item3.Key;
									}
								}
								selectedValues = HTKKWinDictionaryDiaDanh.getMaCoQuanThueQuanLy(winFpSpread.ActiveSheet.Cells[cStatic.ParentCell.RowIndex - 1, cStatic.ParentCell.ColIndex - 1].Text, winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text);
								if (cStatic.CellSelectedValue != null)
								{
									winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
								}
								if (string.IsNullOrEmpty(selectedValues))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
								}
							}
						}
						else if (cStatic.Controltype == "56")
						{
							foreach (KeyValuePair<string, string> dMQuocGium in WinSpreadProcess.GetDMQuocGia())
							{
								if (dMQuocGium.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
								{
									selectedValues = dMQuocGium.Key;
									break;
								}
							}
						}
						if (cStatic.CellSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[cStatic.CellSelectedValue.RowIndex - 1, cStatic.CellSelectedValue.ColIndex - 1].Text = selectedValues;
						}
					}
					else
					{
						cStatic = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Footer.StaticCells);
						if (cStatic != null)
						{
							cStatic.Value = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
						}
						HTKKCMCTableCellInfo tableCellInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableCellInfo(e.Row + 1, e.Column + 1);
						if ((e.Row + 1 >= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart) & (e.Row + 1 <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count))
						{
							for (int num = 0; num < hTKKCMCTableInfo.RowsInfo.Cells.Count; num++)
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[num].CellTemplate.ColIndex != e.Column + 1)
								{
									continue;
								}
								if (!HTKKFormater.isNullString(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Formula))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].ResetFormula();
								}
								if (WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].TaiDuLieuRatLon != "1")
								{
									hTKKCMCTableInfo.RowsValues[e.Row + 1 - hTKKCMCTableInfo.Location.DataLocationTemplateRowStart].RowValue[num] = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
								}
								if (tableCellInfo.Controltype == "3" && tableCellInfo.indexColumnSelectedValue != null)
								{
									HTKKWinDictionaryDiaDanh.QuanHe.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
									if (string.IsNullOrEmpty(selectedValues))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
									}
								}
								if (tableCellInfo.Controltype == "4" && tableCellInfo.indexColumnSelectedValue != null)
								{
									HTKKWinDictionaryDiaDanh.QuocTich.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
									if (string.IsNullOrEmpty(selectedValues))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
										if (tableCellInfo.ChildValue != null)
										{
											WinSpreadProcess.WinInsertComboxKey(winFpSpread.ActiveSheet, e.Row + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getTinhTP(selectedValues));
										}
									}
								}
								if (tableCellInfo.Controltype == "5")
								{
									HTKKWinDictionaryDiaDanh.dicTinhTP.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
									if (string.IsNullOrEmpty(selectedValues))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (tableCellInfo.indexColumnSelectedValue != null)
										{
											winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
										}
										if (tableCellInfo.ChildValue != null)
										{
											WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[selectedValues]);
										}
									}
								}
								if (tableCellInfo.Controltype == "100")
								{
									HTKKWinDictionaryDiaDanh.dicTinhTPMoi.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
									if (string.IsNullOrEmpty(selectedValues))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (tableCellInfo.indexColumnSelectedValue != null)
										{
											winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
										}
										if (tableCellInfo.ChildValue != null)
										{
											WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(selectedValues));
										}
									}
									GeneralFunctions.resetCqtTheoTinhDinamic(winFpSpread.ActiveSheet, hTKKCMCTableInfo.RowsInfo, winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1);
								}
								if (tableCellInfo.Controltype == "103")
								{
									HTKKWinDictionaryDiaDanh.dicThueTinh.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
									if (string.IsNullOrEmpty(selectedValues))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (tableCellInfo.indexColumnSelectedValue != null)
										{
											winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
										}
										if (tableCellInfo.ChildValue != null)
										{
											string[] values = (from x in HTKKWinDictionaryDiaDanh.coQuanThus
												where x.MaCQT == selectedValues
												select x.Ten).ToArray();
											WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), values);
										}
									}
								}
								if (tableCellInfo.Controltype == "6")
								{
									if (tableCellInfo.ParentValue != null)
									{
										maTinhTP = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
									}
									maQuanHuyen = HTKKWinDictionaryDiaDanh.getMaQuanHuyen(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, maTinhTP);
									if (string.IsNullOrEmpty(maQuanHuyen))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (tableCellInfo.indexColumnSelectedValue != null)
										{
											winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maQuanHuyen;
										}
										if (tableCellInfo.ChildValue != null)
										{
											WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getPhuongXa(maTinhTP, winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text));
										}
									}
								}
								if (tableCellInfo.Controltype == "7")
								{
									if (tableCellInfo.ParentValue != null)
									{
										maQuanHuyen = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
									}
									string maPhuongXa = HTKKWinDictionaryDiaDanh.getMaPhuongXa(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, maQuanHuyen);
									if (string.IsNullOrEmpty(maPhuongXa))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else if (tableCellInfo.indexColumnSelectedValue != null)
									{
										winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maPhuongXa;
									}
								}
								if (tableCellInfo.Controltype == "101")
								{
									if (tableCellInfo.ParentValue != null)
									{
										maTinhTP = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
									}
									string maPhuongXa = HTKKWinDictionaryDiaDanh.getMaPhuongXaMoi(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, maTinhTP);
									if (string.IsNullOrEmpty(maPhuongXa))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (tableCellInfo.indexColumnSelectedValue != null)
										{
											winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maPhuongXa;
										}
										GeneralFunctions.setSviewCqtCapCucTheoMaXa_Dinamic(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1, hTKKCMCTableInfo.RowsInfo, maPhuongXa);
									}
								}
								if (tableCellInfo.Controltype == "102")
								{
									string maCQT2 = "";
									if (tableCellInfo.ParentValue != null)
									{
										maCQT2 = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
									}
									string ma = HTKKWinDictionaryDiaDanh.coQuanThus.FirstOrDefault((CoQuanThu x) => x.Ten == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text && x.MaCQT == maCQT2).Ma;
									if (string.IsNullOrEmpty(ma))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else if (tableCellInfo.indexColumnSelectedValue != null)
									{
										winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = ma;
									}
								}
								if (tableCellInfo.Controltype == "8")
								{
									if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
									{
										Dictionary<string, string> source = GeneralFunctions.GetDMFromCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
										string tenCqt = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
										if (tenCqt == "Chọn chi cục thuế khu vực")
										{
											return;
										}
										string key = source.First((KeyValuePair<string, string> x) => x.Value == tenCqt).Key;
										if (string.IsNullOrEmpty(key))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											if (tableCellInfo.indexColumnSelectedValue != null)
											{
												winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = key;
											}
											if (tableCellInfo.ChildValue != null)
											{
												GeneralFunctions.setSviewCqtCapChiCucTheoMaCqtCuc_Dinamic(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1, hTKKCMCTableInfo.RowsInfo, key);
											}
										}
									}
									else
									{
										HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.TryGetValue(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text, out selectedValues);
										if (string.IsNullOrEmpty(selectedValues))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											if (tableCellInfo.indexColumnSelectedValue != null)
											{
												winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
											}
											if (tableCellInfo.ChildValue != null)
											{
												WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(selectedValues));
											}
										}
									}
								}
								if (tableCellInfo.Controltype == "9")
								{
									if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
									{
										Dictionary<string, string> source = GeneralFunctions.GetDMFromCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
										string tenCqt2 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
										if (tenCqt2 == "Chọn cơ quan thuế")
										{
											return;
										}
										string key = source.First((KeyValuePair<string, string> x) => x.Value == tenCqt2).Key;
										if (tableCellInfo.indexColumnSelectedValue != null)
										{
											winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = key;
										}
									}
									else
									{
										foreach (KeyValuePair<string, string> item4 in HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy)
										{
											if (item4.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
											{
												selectedValues = item4.Key;
												if (tableCellInfo.indexColumnSelectedValue != null)
												{
													winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
												}
											}
										}
										if (string.IsNullOrEmpty(selectedValues))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
									}
								}
								if (tableCellInfo.Controltype == "20")
								{
									string text5 = "";
									foreach (KeyValuePair<string, string> item5 in WinSpreadProcess.DMKhoangSan("GetDMKhoangSan"))
									{
										if (item5.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
										{
											text5 = item5.Key;
											break;
										}
									}
									if (string.IsNullOrEmpty(text5))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text5))
										{
											string text6 = "";
											text6 = text5;
											if (text5.Split(new char[1] { '_' }).Length > 3)
											{
												text6 = text5.Split(new char[1] { '_' })[2];
											}
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
										}
										if (!string.IsNullOrEmpty(text5) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
										{
											Dictionary<string, string> lstDonVi = new Dictionary<string, string>();
											string text7 = WinSpreadProcess.DMKhoangSanDonVi(ref lstDonVi, "GetDMKhoangSan")[text5];
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text7;
										}
									}
								}
								if (tableCellInfo.Controltype == "62")
								{
									string text5 = "";
									foreach (KeyValuePair<string, string> item6 in WinSpreadProcess.DMKhoangSan("GetDMKSBVMT"))
									{
										if (item6.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
										{
											text5 = item6.Key;
											break;
										}
									}
									if (string.IsNullOrEmpty(text5))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else
									{
										if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text5))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text5;
										}
										if (!string.IsNullOrEmpty(text5) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
										{
											Dictionary<string, string> lstDonVi = new Dictionary<string, string>();
											string text7 = WinSpreadProcess.DMKhoangSanDonVi(ref lstDonVi, "GetDMKSBVMT")[text5];
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text7;
										}
									}
								}
								if (tableCellInfo.Controltype == "23")
								{
									string text5 = "";
									foreach (KeyValuePair<string, string> item7 in WinSpreadProcess.DMTaiNguyen())
									{
										if (item7.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
										{
											text5 = item7.Key;
											break;
										}
									}
									if (string.IsNullOrEmpty(text5))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text5))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text5;
									}
								}
								if (tableCellInfo.Controltype == "22")
								{
									string text8 = "";
									foreach (KeyValuePair<string, string> item8 in WinSpreadProcess.DMPhiLephi())
									{
										if (item8.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
										{
											text8 = item8.Key;
											break;
										}
									}
									if (string.IsNullOrEmpty(text8))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = "";
									}
									else
									{
										if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text8))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text8;
										}
										if (!string.IsNullOrEmpty(text8) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
										{
											string text9 = WinSpreadProcess.DMPhiLephi_TieuMuc()[text8];
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text9;
										}
									}
								}
								if (tableCellInfo.Controltype == "88")
								{
									string text8 = "";
									foreach (KeyValuePair<string, string> item9 in WinSpreadProcess.DMPhiLephiTt63())
									{
										if (item9.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
										{
											text8 = item9.Key;
											break;
										}
									}
									if (string.IsNullOrEmpty(text8))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = "";
									}
									else
									{
										if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text8))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text8;
										}
										if (!string.IsNullOrEmpty(text8) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
										{
											string text9 = WinSpreadProcess.DMPhiLephiTt63_TieuMuc()[text8];
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text9;
										}
									}
								}
								if (tableCellInfo.Controltype == "86")
								{
									string text10 = "";
									Dictionary<string, string> source = WinSpreadProcess.DMPHLPNG();
									foreach (KeyValuePair<string, string> item10 in source)
									{
										if (item10.Value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text)
										{
											text10 = item10.Key;
											break;
										}
									}
									item itemDMPHLPNGById = WinSpreadProcess.GetItemDMPHLPNGById(text10);
									if (string.IsNullOrEmpty(text10))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = "";
									}
									else
									{
										if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text10))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = itemDMPHLPNGById.ma;
										}
										if (!string.IsNullOrEmpty(text10) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = itemDMPHLPNGById.matieumuc;
										}
									}
								}
								if (tableCellInfo.Controltype == "25")
								{
									string text11 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMLoaiTemRuou().ContainsKey(text11) && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text11))
									{
										selectedValues = WinSpreadProcess.DMLoaiTemRuou()[text11];
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = selectedValues;
									}
									if (string.IsNullOrEmpty(selectedValues))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
									}
									else if (!string.IsNullOrEmpty(tableCellInfo.ChildValue))
									{
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text11;
									}
								}
								if (tableCellInfo.Controltype == "28" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text12 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMGhiChuTemRuou().ContainsValue(text12))
									{
										string text13 = "";
										foreach (KeyValuePair<string, string> item11 in WinSpreadProcess.DMGhiChuTemRuou())
										{
											if (item11.Value.Equals(text12))
											{
												text13 = item11.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text13))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text13;
										}
									}
								}
								if (tableCellInfo.Controltype == "73" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text14 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMLoaiTienTe().ContainsValue(text14))
									{
										string text15 = "";
										foreach (KeyValuePair<string, string> item12 in WinSpreadProcess.DMLoaiTienTe())
										{
											if (item12.Value.Equals(text14))
											{
												text15 = item12.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text15))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text15;
										}
									}
								}
								if (tableCellInfo.Controltype == "80" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text16 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMThueSuat().ContainsValue(text16))
									{
										string text17 = "";
										foreach (KeyValuePair<string, string> item13 in WinSpreadProcess.DMThueSuat())
										{
											if (item13.Value.Equals(text16))
											{
												text17 = item13.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text17))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text17;
										}
									}
								}
								if (tableCellInfo.Controltype == "72" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text18 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMMo().ContainsValue(text18))
									{
										string text19 = "";
										foreach (KeyValuePair<string, string> item14 in WinSpreadProcess.DMMo())
										{
											if (item14.Value.Equals(text18))
											{
												text19 = item14.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text19))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text19;
										}
									}
								}
								if (tableCellInfo.Controltype == "81" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text20 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMLoaiDat().ContainsValue(text20))
									{
										string text21 = "";
										foreach (KeyValuePair<string, string> item15 in WinSpreadProcess.DMLoaiDat())
										{
											if (item15.Value.Equals(text20))
											{
												text21 = item15.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text21))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text21;
										}
									}
								}
								if (tableCellInfo.Controltype == "82" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text22 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMHangDat().ContainsValue(text22))
									{
										string text23 = "";
										foreach (KeyValuePair<string, string> item16 in WinSpreadProcess.DMHangDat())
										{
											if (item16.Value.Equals(text22))
											{
												text23 = item16.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text23))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text23;
										}
									}
								}
								if (tableCellInfo.Controltype == "87" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text24 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMNnkdNtnn().ContainsValue(text24))
									{
										string text25 = "";
										foreach (KeyValuePair<string, string> item17 in WinSpreadProcess.DMNnkdNtnn())
										{
											if (item17.Value.Equals(text24))
											{
												text25 = item17.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text25))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text25;
										}
									}
								}
								if (tableCellInfo.Controltype == "89" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text14 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMLoaiTienDk().ContainsValue(text14))
									{
										string text15 = "";
										foreach (KeyValuePair<string, string> item18 in WinSpreadProcess.DMNnkdNtnn())
										{
											if (item18.Value.Equals(text14))
											{
												text15 = item18.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text15))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text15;
										}
									}
								}
								if (tableCellInfo.Controltype != null && tableCellInfo.Controltype.Length > 3 && tableCellInfo.Controltype.Substring(0, 3) == "DC_" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									string text26 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.DMLoaiTienDk().ContainsValue(text26))
									{
										string text27 = "";
										foreach (KeyValuePair<string, string> item19 in WinSpreadProcess.GetDMDungChung(tableCellInfo.Controltype))
										{
											if (item19.Value.Equals(text26))
											{
												text27 = item19.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text27))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
										}
									}
								}
								string text28;
								string text29;
								if (tableCellInfo.Controltype == "35" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									if (WinSpreadProcess.GetDanhMuc("GetDMThueBVMT").ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item20 in WinSpreadProcess.GetDanhMuc("GetDMThueBVMT"))
										{
											if (item20.Value.Equals(text28))
											{
												text29 = item20.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "50" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									Dictionary<string, string> khoanNopNSSHoanThue = HTKKWinDictionaryDiaDanh.KhoanNopNSSHoanThue;
									if (khoanNopNSSHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item21 in khoanNopNSSHoanThue)
										{
											if (item21.Value.Equals(text28))
											{
												text29 = item21.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								Dictionary<string, string> kyPhatSinhHoanThue;
								if (tableCellInfo.Controltype == "51" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = HTKKWinDictionaryDiaDanh.KyPhatSinhHoanThue;
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item22 in kyPhatSinhHoanThue)
										{
											if (item22.Value.Equals(text28))
											{
												text29 = item22.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "90" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetMaTieuMuc();
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item23 in kyPhatSinhHoanThue)
										{
											if (item23.Value.Equals(text28))
											{
												text29 = item23.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "91" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMNganhCoThuNhapCap();
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item24 in kyPhatSinhHoanThue)
										{
											if (item24.Value.Equals(text28))
											{
												text29 = item24.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "92" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMChuong();
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item25 in kyPhatSinhHoanThue)
										{
											if (item25.Value.Equals(text28))
											{
												text29 = item25.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "93" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMNganHang();
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item26 in kyPhatSinhHoanThue)
										{
											if (item26.Value.Equals(text28))
											{
												text29 = item26.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "106" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMDynamic("DM_NganHang_Moi.xml");
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item27 in kyPhatSinhHoanThue)
										{
											if (item27.Value.Equals(text28))
											{
												text29 = item27.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "107" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMDynamic("DM_Kho_Bac.xml");
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item28 in kyPhatSinhHoanThue)
										{
											if (item28.Value.Equals(text28))
											{
												text29 = item28.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "108" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMDynamic("DM_CQHQ.xml");
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item29 in kyPhatSinhHoanThue)
										{
											if (item29.Value.Equals(text28))
											{
												text29 = item29.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										else
										{
											winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
										}
									}
								}
								if (tableCellInfo.Controltype == "98" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									kyPhatSinhHoanThue = WinSpreadProcess.GetDMNganHangKBNN();
									if (kyPhatSinhHoanThue.ContainsValue(text28))
									{
										text29 = "";
										foreach (KeyValuePair<string, string> item30 in kyPhatSinhHoanThue)
										{
											if (item30.Value.Equals(text28))
											{
												text29 = item30.Key;
												break;
											}
										}
										if (string.IsNullOrEmpty(text29))
										{
											winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
										}
										winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
									}
								}
								if (tableCellInfo.Controltype == null || tableCellInfo.Controltype.Length <= 3 || !(tableCellInfo.Controltype.Substring(0, 3) == "DC_") || string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
								{
									continue;
								}
								text28 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
								kyPhatSinhHoanThue = WinSpreadProcess.GetDMDungChung(tableCellInfo.Controltype);
								if (WinSettings.MenuID == "N8" && (tableCellInfo.Controltype == "DC_08" || tableCellInfo.Controltype == "DC_13"))
								{
									kyPhatSinhHoanThue = GeneralFunctions.GetDMFromCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1);
								}
								if (!kyPhatSinhHoanThue.ContainsValue(text28))
								{
									continue;
								}
								text29 = "";
								foreach (KeyValuePair<string, string> item31 in kyPhatSinhHoanThue)
								{
									if (item31.Value.Equals(text28))
									{
										text29 = item31.Key;
										break;
									}
								}
								if (string.IsNullOrEmpty(text29))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = "";
								}
								else
								{
									winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
								}
							}
						}
					}
				}
			}
			WinDoWorkEvent.CellChange(winFpSpread, winFpSpread.ActiveSheetIndex, e.Row + 1, e.Column + 1);
			cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
			if (winFpSpread.ActiveSheet.SheetName == "01-1/KHBS")
			{
				CellChange_Event_01_1_KHBS(e.Column, e.Row);
			}
			if (winFpSpread.ActiveSheet.SheetName == "01-KHBS")
			{
				CellChange_Event_01_KHBS(e.Column, e.Row);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_editchange(object sender, EditorNotifyEventArgs e)
	{
		winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Formula = "";
		string text = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
		MaskCellType maskCellType = new MaskCellType();
		maskCellType.Mask = "00/00/0000";
		maskCellType.MaskChar = '.';
		isDataChanged = true;
		HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
		if (winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 2 && winFpSpread.ActiveSheetIndex != winFpSpread.Sheets.Count - 3)
		{
			isDataChangedCallTHKhbs = true;
		}
		if (staticCell != null)
		{
			if (staticCell.Controltype == "16")
			{
				if (string.IsNullOrEmpty(ConvertHelpers.ConvertToString(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value)))
				{
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = (string.IsNullOrEmpty(staticCell.DefaultValue) ? "0" : staticCell.DefaultValue);
				}
			}
			else if (staticCell.Controltype == "14")
			{
				text = text.Replace("//", "");
				string text2 = ((text.Count((char f) => f == '/') == 2) ? (text.Split(new char[1] { '/' })[2] + text.Split(new char[1] { '/' })[1] + text.Split(new char[1] { '/' })[0]) : "");
				text2 = text2.Replace(".", "");
				if (ConvertHelpers.ConvertToInt(text2) == 0 && !text2.Equals("0"))
				{
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value = "";
				}
			}
		}
		else
		{
			int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
			if (tableIndex >= 0)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
				staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Header.StaticCells);
				if (staticCell != null)
				{
					if (staticCell.Controltype == "16")
					{
						if (string.IsNullOrEmpty(ConvertHelpers.ConvertToString(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value)))
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = (string.IsNullOrEmpty(staticCell.DefaultValue) ? "0" : staticCell.DefaultValue);
						}
					}
					else if (staticCell.Controltype == "14")
					{
						string text2 = ((text.Count((char f) => f == '/') == 2) ? (text.Split(new char[1] { '/' })[2] + text.Split(new char[1] { '/' })[1] + text.Split(new char[1] { '/' })[0]) : "");
						text2 = text2.Replace(".", "");
						if (ConvertHelpers.ConvertToInt(text2) == 0 && !text2.Equals("0"))
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value = "";
						}
					}
				}
				else
				{
					staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Footer.StaticCells);
					if (staticCell != null)
					{
						if (staticCell.Controltype == "16")
						{
							if (string.IsNullOrEmpty(ConvertHelpers.ConvertToString(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value)))
							{
								staticCell.Value = (string.IsNullOrEmpty(staticCell.DefaultValue) ? "0" : staticCell.DefaultValue);
							}
						}
						else if (staticCell.Controltype == "14")
						{
							string text2 = ((text.Count((char f) => f == '/') == 2) ? (text.Split(new char[1] { '/' })[2] + text.Split(new char[1] { '/' })[1] + text.Split(new char[1] { '/' })[0]) : "");
							text2 = text2.Replace(".", "");
							if (ConvertHelpers.ConvertToInt(text2) == 0 && !text2.Equals("0"))
							{
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value = "";
							}
						}
					}
					if ((e.Row + 1 >= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart) & (e.Row + 1 <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count))
					{
						for (int num = 0; num < hTKKCMCTableInfo.RowsInfo.Cells.Count; num++)
						{
							if (hTKKCMCTableInfo.RowsInfo.Cells[num].CellTemplate.ColIndex != e.Column + 1)
							{
								continue;
							}
							if (hTKKCMCTableInfo.RowsInfo.Cells[num].Controltype == "16")
							{
								if (string.IsNullOrEmpty(ConvertHelpers.ConvertToString(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value)))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsInfo.Cells[num].DefaultValue) ? "0" : hTKKCMCTableInfo.RowsInfo.Cells[num].DefaultValue);
								}
							}
							else if (hTKKCMCTableInfo.RowsInfo.Cells[num].Controltype == "14")
							{
								string text2 = ((text.Count((char f) => f == '/') == 2) ? (text.Split(new char[1] { '/' })[2] + text.Split(new char[1] { '/' })[1] + text.Split(new char[1] { '/' })[0]) : "");
								text2 = text2.Replace(".", "");
								if (ConvertHelpers.ConvertToInt(text2) == 0 && !text2.Equals("0"))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value = "";
								}
							}
						}
					}
				}
			}
		}
		cellChange(e.Row, e.Column);
		WinDoWorkEvent.CellEditchange(winFpSpread, winFpSpread.ActiveSheetIndex, e.Row + 1, e.Column + 1);
	}

	private void winFpSpread_ComboCloseUp(object sender, EditorNotifyEventArgs e)
	{
		if (string.IsNullOrEmpty(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text))
		{
			return;
		}
		try
		{
			cellChange(e.Row, e.Column);
			if (WinSettings.MenuID == "G9")
			{
				winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveRow.Index, 0);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_ComboSelChange(object sender, EditorNotifyEventArgs e)
	{
		//IL_0521: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_03d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_08ce: Unknown result type (might be due to invalid IL or missing references)
		if (string.IsNullOrEmpty(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text))
		{
			return;
		}
		try
		{
			string value = "";
			selectValueOld.TryGetValue(e.Row.ToString() + e.Column, out value);
			if (value == winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text && WinSettings.TTPeriodInfo.menuID != "91" && WinSettings.TTPeriodInfo.menuID != "79")
			{
				return;
			}
			HTKKFile.AddOrUpdateDicStr(selectValueOld, e.Row.ToString() + e.Column, winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text);
			HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
			if (staticCell != null)
			{
				return;
			}
			HTKKCMCTableCellInfo tableCellInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableCellInfo(e.Row + 1, e.Column + 1);
			if (tableCellInfo.Controltype == "39")
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string value2 = sheetView.Cells[e.Row, e.Column].Value.ToString();
				string[] array = new string[4] { "01_TTDB_I", "01_TTDB_III_A", "01_TTDB_III_B", "01_TTDB_III_C" };
				if (Convert.ToBoolean(sheetView.Cells[33, HTKKFile.ColumnLettersToIndex("N") - 1].Value))
				{
					sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex("AD") - 1].Text = "0";
				}
				string[] array2 = array;
				foreach (string tableName in array2)
				{
					HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
					HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
					List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
					Predicate<HTKKCMCTableInfo> match = (HTKKCMCTableInfo table) => table.TableName == tableName;
					HTKKCMCTableInfo hTKKCMCTableInfo = tables.Find(match);
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
					string empty = string.Empty;
					string empty2 = string.Empty;
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					for (int num2 = dataLocationTemplateRowStart; num2 <= num; num2++)
					{
						empty = sheetView.Cells[num2 - 1, e.Column].Text;
						if (empty.Equals(value2) && num2 - 1 != e.Row && e.Row >= dataLocationTemplateRowStart && e.Row <= num)
						{
							sheetView.Cells[e.Row, e.Column].Text = "";
							sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
							frmMessage frmMessage = new frmMessage("0261", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							break;
						}
					}
				}
			}
			if (tableCellInfo.Controltype == "58")
			{
				HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
				SheetView sheetView = winFpSpread.Sheets[0];
				string value2 = sheetView.Cells[e.Row, e.Column].Value.ToString();
				HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				int num3 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				string empty3 = string.Empty;
				for (int num2 = dataLocationTemplateRowStart2; num2 <= num3; num2++)
				{
					empty3 = sheetView.Cells[num2 - 1, e.Column].Text;
					if (empty3.Equals(value2) && num2 - 1 != e.Row)
					{
						frmMessage frmMessage = new frmMessage("0445", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						sheetView.Cells[e.Row, e.Column].Text = "";
						sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
						sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = "";
						break;
					}
				}
			}
			if (tableCellInfo.Controltype == "43")
			{
				HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
				SheetView sheetView = winFpSpread.Sheets[0];
				string value2 = sheetView.Cells[e.Row, e.Column].Value.ToString();
				HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo.Tables[0];
				int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				int num3 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo2.RowsValues.Count - 1;
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				string empty3 = string.Empty;
				for (int num2 = dataLocationTemplateRowStart2; num2 <= num3; num2++)
				{
					empty3 = sheetView.Cells[num2 - 1, e.Column].Text;
					if (empty3.Equals(value2) && num2 - 1 != e.Row)
					{
						frmMessage frmMessage = new frmMessage("0445", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						sheetView.Cells[e.Row, e.Column].Text = "";
						sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
						sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = "";
						break;
					}
				}
			}
			if (tableCellInfo.Controltype == "42")
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string value2 = sheetView.Cells[e.Row, e.Column].Value.ToString();
				string[] array = new string[1] { "01_TTDB_II" };
				string[] array2 = array;
				foreach (string tableName2 in array2)
				{
					HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
					HTKKCMCSheetInfo hTKKCMCSheetInfo = toKhaiInfo.SheetInfos[0];
					List<HTKKCMCTableInfo> tables2 = hTKKCMCSheetInfo.Tables;
					Predicate<HTKKCMCTableInfo> match2 = (HTKKCMCTableInfo table) => table.TableName == tableName2;
					HTKKCMCTableInfo hTKKCMCTableInfo = tables2.Find(match2);
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					int num = dataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count - 1;
					string empty = string.Empty;
					string empty2 = string.Empty;
					Dictionary<string, string> dictionary = new Dictionary<string, string>();
					for (int num2 = dataLocationTemplateRowStart; num2 <= num; num2++)
					{
						empty = sheetView.Cells[num2 - 1, e.Column].Text;
						if (empty.Equals(value2) && num2 - 1 != e.Row && e.Row >= dataLocationTemplateRowStart && e.Row <= num)
						{
							frmMessage frmMessage = new frmMessage("0261", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							sheetView.Cells[e.Row, e.Column].Text = "";
							sheetView.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
							break;
						}
					}
				}
			}
			if (!(tableCellInfo.Controltype == "64"))
			{
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_EnterCell(object sender, EnterCellEventArgs e)
	{
		try
		{
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
			{
				return;
			}
			if (winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType != null && winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType.ToString() != "TaxComboBoxCellType")
			{
				InputMap inputMap = null;
				inputMap = winFpSpread.GetInputMap(InputMapMode.WhenFocused);
				inputMap.Put(new Keystroke((Keys)38, (Keys)0), SpreadActions.MoveToPreviousRow);
				inputMap.Put(new Keystroke((Keys)40, (Keys)0), SpreadActions.MoveToNextRow);
				inputMap.Put(new Keystroke((Keys)37, (Keys)0), SpreadActions.MoveToPreviousColumn);
				inputMap.Put(new Keystroke((Keys)39, (Keys)0), SpreadActions.MoveToNextColumn);
			}
			HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
			if (staticCell != null)
			{
				string empty = string.Empty;
				if (staticCell.Controltype == "8")
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower() || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						string maXa = GeneralFunctions.timMaChildenFromRowCol_Static(winFpSpread.ActiveSheet, hTKKCMCSheetInfo.StaticCells, "101");
						Dictionary<string, string> dic = GeneralFunctions.loadCqtCapCucTheoDiaBanXa(maXa);
						GeneralFunctions.InsertCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, dic);
						return;
					}
				}
				if (staticCell.Controltype == "9")
				{
					if (staticCell.ParentCell != null)
					{
						string text = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
						if (!string.IsNullOrEmpty(text))
						{
							HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
							if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower() || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
							{
								string maXa = GeneralFunctions.timMaChildenFromRowCol_Static(winFpSpread.ActiveSheet, hTKKCMCSheetInfo.StaticCells, "101");
								Dictionary<string, string> dic2 = GeneralFunctions.loadCqtCapChiCucTheoDiaBanXa(text, maXa);
								GeneralFunctions.InsertCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, dic2);
								return;
							}
							Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(text);
							string[] array = new string[coQuanThueQuanLybyMaCoQuanThueCapCuc.Count];
							string[] array2 = new string[coQuanThueQuanLybyMaCoQuanThueCapCuc.Count];
							int num = 0;
							foreach (KeyValuePair<string, string> item in coQuanThueQuanLybyMaCoQuanThueCapCuc)
							{
								string[] array4;
								string[] array3 = (array4 = array);
								int num2 = num;
								nint num3 = num2;
								array3[num2] = array4[num3] + item.Value;
								string[] array5 = (array4 = array2);
								int num4 = num;
								num3 = num4;
								array5[num4] = array4[num3] + item.Key;
								num++;
							}
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = array;
							taxComboBoxCellType.ItemData = array2;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				else if (staticCell.Controltype == "6")
				{
					if (staticCell.ParentCell != null)
					{
						string text = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
						if (!string.IsNullOrEmpty(text))
						{
							string[] items = HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[text];
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = items;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				else if (staticCell.Controltype == "7")
				{
					if (staticCell.ParentCell != null)
					{
						string text = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
						if (!string.IsNullOrEmpty(text))
						{
							string[] items = HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen[text];
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = items;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				else if (staticCell.Controltype == "101")
				{
					if (staticCell.ParentCell != null)
					{
						string text = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
						if (!string.IsNullOrEmpty(text))
						{
							string[] items = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(text);
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = items;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				else if (staticCell.Controltype == "102" && staticCell.ParentCell != null)
				{
					string parentSelectedValue = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
					if (!string.IsNullOrEmpty(parentSelectedValue))
					{
						string[] items = (from x in HTKKWinDictionaryDiaDanh.coQuanThus
							where x.MaCQT == parentSelectedValue
							select x.Ten).ToArray();
						TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
						taxComboBoxCellType.Items = items;
						taxComboBoxCellType.ListWidth = 500;
						winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
					}
					else
					{
						TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
						taxComboBoxCellType.Items = new string[1] { "" };
						winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
					}
				}
			}
			else
			{
				HTKKCMCTableCellInfo tableCellInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableCellInfo(e.Row + 1, e.Column + 1);
				if (tableCellInfo.Controltype == "6")
				{
					if (!string.IsNullOrEmpty(tableCellInfo.ParentValue))
					{
						string text = winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
						if (!string.IsNullOrEmpty(text))
						{
							string[] items = HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[text];
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = items;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				else if (tableCellInfo.Controltype == "7")
				{
					if (!string.IsNullOrEmpty(tableCellInfo.ParentValue))
					{
						string text = winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
						if (!string.IsNullOrEmpty(text))
						{
							string[] items = HTKKWinDictionaryDiaDanh.dicPhuongXaByQuanHuyen[text];
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = items;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				else if (tableCellInfo.Controltype == "8")
				{
					string maXa = GeneralFunctions.timMaChildenFromRowCol_Dinamic(winFpSpread.ActiveSheet, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, "101");
					Dictionary<string, string> dic3 = GeneralFunctions.loadCqtCapCucTheoDiaBanXa(maXa);
					GeneralFunctions.InsertCombobox(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, dic3, deleteText: false);
				}
				else if (tableCellInfo.Controltype == "9")
				{
					if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						string maXa = GeneralFunctions.timMaChildenFromRowCol_Dinamic(winFpSpread.ActiveSheet, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, "101");
						string maCqt = GeneralFunctions.timMaChildenFromRowCol_Dinamic(winFpSpread.ActiveSheet, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex], winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, "8");
						Dictionary<string, string> dic3 = GeneralFunctions.loadCqtCapChiCucTheoDiaBanXa(maCqt, maXa);
						GeneralFunctions.InsertCombobox(winFpSpread.ActiveSheet, winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1, dic3, deleteText: false);
						return;
					}
					if (!string.IsNullOrEmpty(tableCellInfo.ParentValue))
					{
						string text = winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
						string text2 = winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text;
						string text3 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
						if (!string.IsNullOrEmpty(text) && !text.Equals("0"))
						{
							Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(text);
							string[] array = new string[coQuanThueQuanLybyMaCoQuanThueCapCuc.Count];
							string[] array2 = new string[coQuanThueQuanLybyMaCoQuanThueCapCuc.Count];
							int num = 0;
							foreach (KeyValuePair<string, string> item2 in coQuanThueQuanLybyMaCoQuanThueCapCuc)
							{
								string[] array4;
								string[] array6 = (array4 = array);
								int num5 = num;
								nint num3 = num5;
								array6[num5] = array4[num3] + item2.Value;
								string[] array7 = (array4 = array2);
								int num6 = num;
								num3 = num6;
								array7[num6] = array4[num3] + item2.Key;
								num++;
							}
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = array;
							taxComboBoxCellType.ItemData = array2;
							taxComboBoxCellType.ListWidth = 500;
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
							WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, coQuanThueQuanLybyMaCoQuanThueCapCuc);
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text3;
							winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
						}
						else
						{
							TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
							taxComboBoxCellType.Items = new string[1] { "" };
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
						}
					}
				}
				int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
				if (tableIndex >= 0)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
					staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Header.StaticCells);
					if (staticCell != null)
					{
						staticCell.Value = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
						if (staticCell.Controltype == "8")
						{
							HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
							if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
							{
								string maXa = GeneralFunctions.timMaChildenFromRowCol_Static(winFpSpread.ActiveSheet, hTKKCMCTableInfo.Header.StaticCells, "101");
								Dictionary<string, string> dic = GeneralFunctions.loadCqtCapCucTheoDiaBanXa(maXa);
								GeneralFunctions.InsertCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, dic);
								return;
							}
						}
						if (staticCell.Controltype == "9")
						{
							if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
							{
								if (staticCell.ParentCell != null)
								{
									string text = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
									if (!string.IsNullOrEmpty(text))
									{
										string maXa = GeneralFunctions.timMaChildenFromRowCol_Static(winFpSpread.ActiveSheet, hTKKCMCTableInfo.Header.StaticCells, "101");
										Dictionary<string, string> dic2 = GeneralFunctions.loadCqtCapChiCucTheoDiaBanXa(text, maXa);
										GeneralFunctions.InsertCombobox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, dic2);
									}
								}
								return;
							}
							if (staticCell.ParentCell != null)
							{
								string text = winFpSpread.ActiveSheet.Cells[e.Row, staticCell.ParentCell.ColIndex - 1].Text;
								string text2 = winFpSpread.ActiveSheet.Cells[e.Row, staticCell.CellSelectedValue.ColIndex - 1].Text;
								string text3 = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
								if (!string.IsNullOrEmpty(text) && !text.Equals("0"))
								{
									Dictionary<string, string> coQuanThueQuanLybyMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(text);
									string[] array = new string[coQuanThueQuanLybyMaCoQuanThueCapCuc.Count];
									string[] array2 = new string[coQuanThueQuanLybyMaCoQuanThueCapCuc.Count];
									int num = 0;
									foreach (KeyValuePair<string, string> item3 in coQuanThueQuanLybyMaCoQuanThueCapCuc)
									{
										string[] array4;
										string[] array8 = (array4 = array);
										int num7 = num;
										nint num3 = num7;
										array8[num7] = array4[num3] + item3.Value;
										string[] array9 = (array4 = array2);
										int num8 = num;
										num3 = num8;
										array9[num8] = array4[num3] + item3.Key;
										num++;
									}
									TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
									taxComboBoxCellType.Items = array;
									taxComboBoxCellType.ItemData = array2;
									taxComboBoxCellType.ListWidth = 500;
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
									WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, coQuanThueQuanLybyMaCoQuanThueCapCuc);
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text3;
									winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
								}
								else
								{
									TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
									taxComboBoxCellType.Items = new string[1] { "" };
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
								}
							}
						}
						else if (staticCell.Controltype == "101" && staticCell.ParentCell != null)
						{
							string text = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
							if (!string.IsNullOrEmpty(text))
							{
								string[] items = HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(text);
								TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
								taxComboBoxCellType.Items = items;
								taxComboBoxCellType.ListWidth = 500;
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
							}
							else
							{
								TaxComboBoxCellType taxComboBoxCellType = new TaxComboBoxCellType();
								taxComboBoxCellType.Items = new string[1] { "" };
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = taxComboBoxCellType;
								if (staticCell.CellSelectedValue != null)
								{
									winFpSpread.ActiveSheet.Cells[staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1].Text = "";
								}
							}
						}
					}
				}
			}
			if (WinSettings.MenuID == "71")
			{
				InputMap inputMap = null;
				if (e.Column == HTKKFile.ColumnLettersToIndex("K") - 1)
				{
					inputMap = winFpSpread.GetInputMap(InputMapMode.WhenFocused);
					inputMap.Put(new Keystroke((Keys)9, (Keys)0), SpreadActions.MoveToNextColumnFirstRow);
				}
			}
			if (WinSettings.MenuID == "B8" || WinSettings.MenuID == "L2" || WinSettings.MenuID == "G5")
			{
				WinDoWorkEvent.EnterCell(winFpSpread, winFpSpread.ActiveSheetIndex, e.Row + 1, e.Column + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_ComboCloseDown(object sender, EditorNotifyEventArgs e)
	{
	}

	private void winFpSpread_ButtonClicked(object sender, EditorNotifyEventArgs e)
	{
		try
		{
			ButtonCellType buttonCellType = new ButtonCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			if ((object)winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType.GetType() == buttonCellType.GetType() || (object)winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType.GetType() == checkBoxCellType.GetType())
			{
				WinDoWorkEvent.ButtonClick(winFpSpread, e.Row, e.Column);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btnTimKiem_Click(object sender, EventArgs e)
	{
		//IL_06a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0919: Unknown result type (might be due to invalid IL or missing references)
		//IL_08b7: Unknown result type (might be due to invalid IL or missing references)
		int activeSheetIndex;
		int num;
		if (WinSettings.MenuID.Equals("G2"))
		{
			string text = "";
			string text2 = "";
			string text3 = "";
			string text4 = "";
			num = 0;
			activeSheetIndex = winFpSpread.ActiveSheetIndex;
			num = WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex].Tables[0].RowsValues.Count;
			totalRow = startRowTimKiemDADT + num;
			text = ((ListControl)cbbCQTCuc).SelectedValue.ToString();
			if (!text.Equals(""))
			{
				text2 = ((ListControl)cbbCQTChiCuc).SelectedValue.ToString();
			}
			text4 = ((Control)txtTimKiem).Text.Trim();
			for (int i = 0; i < num; i++)
			{
				winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = true;
			}
			if (!text.Equals("") && !text2.Equals(""))
			{
				for (int i = 0; i < num; i++)
				{
					string text5 = "";
					string text6 = "";
					text5 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, HTKKFile.ColumnLettersToIndex("W") - 1].Text;
					text6 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, HTKKFile.ColumnLettersToIndex("X") - 1].Text;
					text3 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, dicTimKiemValue[((Control)cbbTimKiem).Text] - 1].Text;
					if (!text.Equals(text5) || !text2.Equals(text6))
					{
						winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
					}
					else if (!text4.Equals("") && !text3.Contains(text4))
					{
						winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
					}
				}
				return;
			}
			if (!text.Equals(""))
			{
				for (int i = 0; i < num; i++)
				{
					string text5 = "";
					string text6 = "";
					text5 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, HTKKFile.ColumnLettersToIndex("W") - 1].Text;
					text6 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, HTKKFile.ColumnLettersToIndex("X") - 1].Text;
					text3 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, dicTimKiemValue[((Control)cbbTimKiem).Text] - 1].Text;
					if (!text.Equals(text5))
					{
						winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
					}
					else if (!text4.Equals("") && !text3.Contains(text4))
					{
						winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
					}
				}
				return;
			}
			if (!text2.Equals(""))
			{
				for (int i = 0; i < num; i++)
				{
					string text5 = "";
					string text6 = "";
					text5 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, HTKKFile.ColumnLettersToIndex("W") - 1].Text;
					text6 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, HTKKFile.ColumnLettersToIndex("X") - 1].Text;
					text3 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, dicTimKiemValue[((Control)cbbTimKiem).Text] - 1].Text;
					if (!text2.Equals(text6))
					{
						winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
					}
					else if (!text4.Equals("") && !text3.Contains(text4))
					{
						winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
					}
				}
				return;
			}
			for (int i = 0; i < num; i++)
			{
				text3 = winFpSpread.Sheets[activeSheetIndex].Cells[startRowTimKiemDADT + i, dicTimKiemValue[((Control)cbbTimKiem).Text] - 1].Text;
				if (!text4.Equals("") && !text3.Contains(text4))
				{
					winFpSpread.Sheets[activeSheetIndex].Rows[startRowTimKiemDADT + i].Visible = false;
				}
			}
			return;
		}
		if (HTKKFormater.isNullString(((Control)txtTimKiem).Text))
		{
			frmMessage frmMessage = new frmMessage("0160", "3", "1", "Thông báo");
			((Form)frmMessage).ShowDialog();
			return;
		}
		num = 0;
		activeSheetIndex = winFpSpread.ActiveSheetIndex;
		num = WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex].Tables[0].RowsValues.Count;
		totalRow = startRowTimKiem + num;
		int foundRowIndex = 0;
		int foundColumnIndex = 0;
		string text7 = "";
		winFpSpread.Sheets[activeSheetIndex].ClearSelection();
		text7 = winFpSpread.Search(activeSheetIndex, ((Control)txtTimKiem).Text.Trim(), caseSensitive: false, exactMatch: false, alternateSearch: false, useWildcards: false, index, dicTimKiemValue[((Control)cbbTimKiem).Text] - 1, index + totalRow, dicTimKiemValue[((Control)cbbTimKiem).Text] - 1, ref foundRowIndex, ref foundColumnIndex);
		if (foundRowIndex > -1)
		{
			ar = winFpSpread.Sheets[activeSheetIndex].Rows[foundRowIndex];
			index = foundRowIndex + 1;
			winFpSpread.Sheets[activeSheetIndex].SetActiveCell(foundRowIndex, WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex].Tables[0].iColumnSTT - 1);
			winFpSpread.Sheets[activeSheetIndex].AddSelection(foundRowIndex, WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex].Tables[0].iColumnSTT - 1, 1, WinSettings.ToKhaiInfo.SheetInfos[activeSheetIndex].Tables[0].RowsValues[0].RowValue.Count);
			winFpSpread.SetViewportTopRow(0, foundRowIndex);
		}
		else if (index > 21)
		{
			frmMessage frmMessage2 = new frmMessage("0459", "5", "2", "Thông báo");
			((Form)frmMessage2).ShowDialog();
			if (frmMessage2.Yes)
			{
				winFpSpread.Sheets[activeSheetIndex].ClearSelection();
				index = 21;
				btnTimKiem_Click(sender, e);
			}
		}
		else
		{
			frmMessage frmMessage2 = new frmMessage("0160", "3", "", "Thông báo");
			((Form)frmMessage2).ShowDialog();
			index = 21;
		}
	}

	private void LoadButton()
	{
		string path = "Button.XML";
		if (File.Exists(path))
		{
			XmlDataDocument xmlDataDocument = new XmlDataDocument();
			FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read);
			((XmlDocument)xmlDataDocument).Load((Stream)fileStream);
			XmlNodeList elementsByTagName = ((XmlDocument)xmlDataDocument).GetElementsByTagName("Button");
			for (int i = 0; i < elementsByTagName.Count; i++)
			{
				if (!(((XmlNode)elementsByTagName[i].Attributes["ID"]).Value == WinSettings.ToKhaiInfo.MenuID))
				{
					continue;
				}
				string text = ((elementsByTagName[i].Attributes["ThemPhuLuc"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["ThemPhuLuc"]).Value);
				if (text == "1")
				{
					((Control)btnThemPhuLuc).Visible = true;
				}
				else
				{
					((Control)btnThemPhuLuc).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["Ghi"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["Ghi"]).Value);
				if (text == "1")
				{
					((Control)cmdSave).Visible = true;
				}
				else
				{
					((Control)cmdSave).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["Xoa"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["Xoa"]).Value);
				if (text == "1")
				{
					((Control)cmdDelete).Visible = true;
				}
				else
				{
					((Control)cmdDelete).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["TaiToKhai"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["TaiToKhai"]).Value);
				if (text == "1")
				{
					((Control)cmdLoadToKhai).Visible = true;
				}
				else
				{
					((Control)cmdLoadToKhai).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["NhapLai"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["NhapLai"]).Value);
				if (text == "1")
				{
					((Control)cmdClear).Visible = true;
				}
				else
				{
					((Control)cmdClear).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["In"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["In"]).Value);
				if (text == "1")
				{
					((Control)cmdPrint).Visible = true;
				}
				else
				{
					((Control)cmdPrint).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["KiemTra"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["KiemTra"]).Value);
				if (text == "1")
				{
					((Control)cmdKiemTra).Visible = true;
				}
				else
				{
					((Control)cmdKiemTra).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["KetXuat"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["KetXuat"]).Value);
				if (text == "1")
				{
					((Control)cmdExport).Visible = true;
				}
				else
				{
					((Control)cmdExport).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["NhapTuXML"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["NhapTuXML"]).Value);
				if (text == "1")
				{
					((Control)cmdImportXML).Visible = true;
				}
				else
				{
					((Control)cmdImportXML).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["Dong"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["Dong"]).Value);
				if (text == "1")
				{
					((Control)cmdExit).Visible = true;
				}
				else
				{
					((Control)cmdExit).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["TaiBangKe"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["TaiBangKe"]).Value);
				if (text == "1")
				{
					((Control)btnTaiBangKe).Visible = true;
				}
				else
				{
					((Control)btnTaiBangKe).Visible = false;
				}
				text = ((elementsByTagName[i].Attributes["TaiBangKe"] == null) ? "" : ((XmlNode)elementsByTagName[i].Attributes["TaiBangKe"]).Value);
				if (text == "1")
				{
					((Control)btnTaiBangKe).Visible = true;
				}
				else
				{
					((Control)btnTaiBangKe).Visible = false;
				}
				if (WinSettings.TTPeriodInfo.LoaiToKhai == "B" && WinSettings.ToKhaiInfo.KieuKHBS != "" && "I5~K1~K6~K8~K9~M7~N1~N9~O8~O9~".IndexOf(WinSettings.MenuID) < 0)
				{
					((Control)btnTongHopKHBS).Visible = true;
					if (!"30,I5,K9,K1,M7,K6,K8,73,55,71,31,P2,X6".Contains(WinSettings.MenuID))
					{
						((Control)btnThemPhuLuc).Visible = false;
					}
				}
				else
				{
					((Control)btnTongHopKHBS).Visible = false;
				}
			}
			if (WinSettings.ToKhaiInfo.MenuID == "N7" && WinSettings.TTPeriodInfo.LoaiToKhai == "B")
			{
				((Control)btnThemPhuLuc).Visible = false;
				((Control)cmdSave).Visible = false;
				((Control)cmdDelete).Visible = false;
				((Control)cmdLoadToKhai).Visible = false;
				((Control)cmdClear).Visible = false;
				((Control)cmdPrint).Visible = false;
				((Control)cmdKiemTra).Visible = false;
				((Control)cmdExport).Visible = false;
				((Control)cmdImportXML).Visible = false;
				((Control)btnTaiBangKe).Visible = false;
			}
		}
		if (WinSettings.MenuID == "30" && WinSettings.TTPeriodInfo.KieuKy == "Q")
		{
			((Control)btnCustom).Visible = true;
		}
		if (WinSettings.ToKhaiInfo.MenuID == "95")
		{
			((Control)btnTaiBangKe).Visible = true;
		}
		ArrangeHorizonPanel(pnButon);
	}

	private void ArrangeHorizonPanel(Panel pn)
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected O, but got Unknown
		//IL_0187: Unknown result type (might be due to invalid IL or missing references)
		//IL_018d: Expected O, but got Unknown
		//IL_02af: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b5: Expected O, but got Unknown
		int num = 0;
		ControlProperties item = default(ControlProperties);
		if (listControlProperties.Count == 0)
		{
			foreach (Control item2 in (ArrangedElementCollection)((Control)pn).Controls)
			{
				Control val = item2;
				item.Name = val.Name;
				item.X = val.Location.X;
				item.Visible = val.Visible;
				listControlProperties.Add(item);
				num++;
			}
			for (int i = 0; i < listControlProperties.Count - 1; i++)
			{
				for (int j = i + 1; j < listControlProperties.Count; j++)
				{
					if (listControlProperties[i].X > listControlProperties[j].X)
					{
						item = listControlProperties[i];
						listControlProperties[i] = listControlProperties[j];
						listControlProperties[j] = item;
					}
				}
			}
		}
		else
		{
			for (int i = 0; i < listControlProperties.Count - 1; i++)
			{
				foreach (Control item3 in (ArrangedElementCollection)((Control)pn).Controls)
				{
					Control val = item3;
					if (val.Name == listControlProperties[i].Name)
					{
						item = listControlProperties[i];
						item.Visible = val.Visible;
						listControlProperties[i] = item;
						if (val.Visible)
						{
							num++;
						}
					}
				}
			}
		}
		if (num == 0)
		{
			((Control)pn).Visible = false;
			return;
		}
		((Control)pn).Visible = true;
		int num2 = ((Control)pn).Width;
		int num3 = 0;
		for (int i = listControlProperties.Count - 1; i >= 0; i--)
		{
			if (listControlProperties[i].Visible)
			{
				foreach (Control item4 in (ArrangedElementCollection)((Control)pn).Controls)
				{
					Control val = item4;
					if (val.Name == listControlProperties[i].Name)
					{
						num3 = 2 + val.Width;
						num2 -= num3;
						val.Location = new Point(num2, val.Location.Y);
					}
				}
			}
		}
	}

	private void TimKiem(int iColumn)
	{
		//IL_01fa: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			while (index <= totalRow)
			{
				if (winFpSpread.ActiveSheet.Cells[index, iColumn].Text.ToLower().IndexOf(((Control)txtTimKiem).Text.ToLower()) != -1)
				{
					if (WinSettings.MenuID.Equals("57"))
					{
						winFpSpread.ActiveSheetIndex = 0;
						winFpSpread.ActiveSheet.SetActiveCell(index, HTKKFile.ColumnLettersToIndex("T") - 1);
					}
					else
					{
						winFpSpread.ActiveSheet.SetActiveCell(index, iColumn);
					}
					ar = winFpSpread.ActiveSheet.ActiveRow;
					winFpSpread.ActiveSheet.ActiveRow.BackColor = Color.DarkCyan;
					result.Add(((Control)txtTimKiem).Text);
					index = winFpSpread.ActiveSheet.ActiveRow.Index + 1;
					break;
				}
				index++;
				if (index == totalRow && winFpSpread.ActiveSheet.Cells[index, iColumn].Text.Contains(((Control)txtTimKiem).Text))
				{
					index = 21;
				}
				else if (index == totalRow && !winFpSpread.ActiveSheet.Cells[index, iColumn].Text.Contains(((Control)txtTimKiem).Text))
				{
					if (result.Count == 0)
					{
						frmMessage frmMessage = new frmMessage("0160", "3", "", "Thông báo");
						((Form)frmMessage).ShowDialog();
						index = 21;
						break;
					}
					if (result.Count != 0)
					{
						result.Clear();
						index = 21;
					}
				}
			}
			if (index >= totalRow)
			{
				index = 21;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void cmdClear_Click(object sender, EventArgs e)
	{
		winFpSpread.SuspendLayout();
		if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon != "1")
		{
			WinDoWorkEvent.UpdateObjectValue(winFpSpread, new Dictionary<int, int>());
		}
		else
		{
			((Control)labelTotalPage).Text = "1";
			GeneralFunctions._totalPage = 1;
			GeneralFunctions._page = 0;
		}
		HTKKFile.AddOrUpdateDicStr(GeneralFunctions.dicVar, "key_action", "CLEAR");
		WinSpreadProcess.ClearData(winFpSpread, winFpSpread.ActiveSheetIndex, isShowMess: true);
		if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
		{
			((Control)txtPage).Text = "1";
		}
		isDataChanged = true;
		isDataChangedCallTHKhbs = true;
		selectValueOld = new Dictionary<string, string>();
		GeneralFunctions.dicVar.Remove("key_action");
		winFpSpread.ResumeLayout();
		Default_01_1_KHBS();
	}

	private void btnTongHopKHBS_Click(object sender, EventArgs e)
	{
		//IL_0f8a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0259: Unknown result type (might be due to invalid IL or missing references)
		//IL_0e3b: Unknown result type (might be due to invalid IL or missing references)
		int num = 0;
		int num2 = 0;
		try
		{
			winFpSpread.SuspendLayout();
			WinSettings.ToKhaiInfo.TaiDuLieuRatLon = WinDoWorkEvent.checkTaiDuLieuLon(WinSettings.MenuID, "TaiDuLieuRatLon");
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1" && WinSettings.MenuID == "P2")
			{
				WinDoWorkEvent.SheetTabChange(winFpSpread, 0);
			}
			if (WinSettings.ToKhaiInfo.KieuKHBS != "3")
			{
				int iCountDieuChinh = 0;
				int iCountDieuChinh2 = 0;
				int iActiveSheet = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
				int[] array = GeneralFunctions.ArrayCellId("BI_5");
				HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet];
				WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet] = hTKKCMCSheetInfo;
				WinSpreadProcess.RemoveDataRows(winFpSpread.Sheets[iActiveSheet], hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart, hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1, hTKKCMCSheetInfo);
				WinSpreadProcess.RemoveDataRows(winFpSpread.Sheets[iActiveSheet], hTKKCMCSheetInfo.Tables[1].Location.DataLocationTemplateRowStart, hTKKCMCSheetInfo.Tables[1].RowsValues.Count - 1, hTKKCMCSheetInfo);
				HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
				hTKKCMCRowValue.RowValue = new List<string> { "", "", "", "0", "0", "0" };
				HTKKCMCRowValue rowValue = hTKKCMCRowValue;
				WinSpreadProcess.SetValueRow(winFpSpread.Sheets[iActiveSheet], hTKKCMCSheetInfo.Tables[0], 0, rowValue);
				WinSpreadProcess.SetValueRow(winFpSpread.Sheets[iActiveSheet], hTKKCMCSheetInfo.Tables[1], 0, rowValue);
				hTKKCMCSheetInfo.Tables[0].RowsValues.Clear();
				hTKKCMCSheetInfo.Tables[1].RowsValues.Clear();
				if (((Control)(Button)sender).Name != "btnTongHopKHBS")
				{
					WinDoWorkEvent.SetUpDataKhbs(winFpSpread, WinSettings.TTPeriodInfo, isClickTongHopSNNC: false, WinSettings.HTKKHeaderInfo);
				}
				else
				{
					WinDoWorkEvent.SetUpDataKhbs(winFpSpread, WinSettings.TTPeriodInfo, isClickTongHopSNNC: true, WinSettings.HTKKHeaderInfo);
				}
				for (num = 0; num < WinSettings.KHBS.Count; num++)
				{
					for (num2 = 0; num2 < WinSettings.KHBS_LanTruoc.StaticCells.Count; num2++)
					{
						if (WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.ColIndex == WinSettings.KHBS[num].Cell.ColIndex && WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.RowIndex == WinSettings.KHBS[num].Cell.RowIndex)
						{
							double num3 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[0].Cells[WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.RowIndex - 1, WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.ColIndex - 1].Value);
							double num4 = ConvertHelpers.ConvertToDouble(WinSettings.ToKhaiInfo_BoSungLanTruoc.SheetInfos[0].StaticCells[num2].Value);
							if (WinSettings.KHBS[num].DieuChinhTang == "1")
							{
								if (num3 - num4 < 0.0)
								{
									TongHopKHBS(hTKKCMCSheetInfo.Tables[0], ref iCountDieuChinh, num3, num4, hTKKCMCSheetInfo, iActiveSheet, num, 0);
								}
								else if (num3 - num4 > 0.0)
								{
									TongHopKHBS(hTKKCMCSheetInfo.Tables[1], ref iCountDieuChinh2, num3, num4, hTKKCMCSheetInfo, iActiveSheet, num, 1);
								}
							}
							else if (num3 - num4 < 0.0)
							{
								TongHopKHBS(hTKKCMCSheetInfo.Tables[1], ref iCountDieuChinh2, num3, num4, hTKKCMCSheetInfo, iActiveSheet, num, 1);
							}
							else if (num3 - num4 > 0.0)
							{
								TongHopKHBS(hTKKCMCSheetInfo.Tables[0], ref iCountDieuChinh, num3, num4, hTKKCMCSheetInfo, iActiveSheet, num, 0);
							}
						}
						double num5 = 0.0;
						double num6 = 0.0;
						int num7 = 0;
						double num8 = 0.0;
						double num9 = 0.0;
						int num10 = 0;
						if (WinSettings.ToKhaiInfo.KieuKHBS.Equals("1"))
						{
							switch (num2)
							{
							case 23:
								num5 = ConvertHelpers.ConvertToDouble(WinSettings.ToKhaiInfo_BoSungLanTruoc.SheetInfos[0].StaticCells[num2].Value);
								num6 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[0].Cells[WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.RowIndex - 1, WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.ColIndex - 1].Value);
								winFpSpread.Sheets[iActiveSheet].Cells[GeneralFunctions.ArrayCellId("BJ_5")[1], GeneralFunctions.ArrayCellId("BJ_5")[0]].Value = num5.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[GeneralFunctions.ArrayCellId("BK_5")[1], GeneralFunctions.ArrayCellId("BK_5")[0]].Value = num6.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[GeneralFunctions.ArrayCellId("BL_5")[1], GeneralFunctions.ArrayCellId("BL_5")[0]].Value = (num6 - num5).ToString();
								break;
							case 26:
								num8 = ConvertHelpers.ConvertToDouble(WinSettings.ToKhaiInfo_BoSungLanTruoc.SheetInfos[0].StaticCells[num2].Value);
								num9 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[0].Cells[WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.RowIndex - 1, WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.ColIndex - 1].Value);
								winFpSpread.Sheets[iActiveSheet].Cells[GeneralFunctions.ArrayCellId("BJ_6")[1], GeneralFunctions.ArrayCellId("BJ_6")[0]].Value = num8.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[GeneralFunctions.ArrayCellId("BK_6")[1], GeneralFunctions.ArrayCellId("BK_6")[0]].Value = num9.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[GeneralFunctions.ArrayCellId("BL_6")[1], GeneralFunctions.ArrayCellId("BL_6")[0]].Value = (num9 - num8).ToString();
								break;
							}
						}
						else if (WinSettings.ToKhaiInfo.KieuKHBS.Equals("2"))
						{
							string path = WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[18].Path;
							FarPoint.Win.Spread.Cell cell = winFpSpread.Sheets[iActiveSheet].Cells[WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[18].CellTemplate.RowIndex - 1, WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[18].CellTemplate.ColIndex - 1];
							if (GeneralFunctions.ArrayCellId(path)[0] == WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.ColIndex - 1 && GeneralFunctions.ArrayCellId(path)[1] == WinSettings.KHBS_LanTruoc.StaticCells[num2].CellTemplate.RowIndex - 1)
							{
								double num11 = ConvertHelpers.ConvertToDouble(WinSettings.ToKhaiInfo_BoSungLanTruoc.SheetInfos[0].StaticCells[num2].Value);
								double num12 = ConvertHelpers.ConvertToDouble(winFpSpread.Sheets[0].Cells[GeneralFunctions.ArrayCellId(path)[1], GeneralFunctions.ArrayCellId(path)[0]].Value);
								WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[15].Value = num11.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[15].CellTemplate.RowIndex - 1, WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[15].CellTemplate.ColIndex - 1].Value = num11;
								WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[16].Value = num12.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[16].CellTemplate.RowIndex - 1, WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[16].CellTemplate.ColIndex - 1].Value = num12;
								double num13 = num12 - num11;
								WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[18].Value = num13.ToString();
								winFpSpread.Sheets[iActiveSheet].Cells[WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[18].CellTemplate.RowIndex - 1, WinSettings.ToKhaiInfo.SheetInfos[iActiveSheet].StaticCells[18].CellTemplate.ColIndex - 1].Value = num13;
							}
						}
					}
				}
				int num14 = 0;
				foreach (HTKKCMCTableInfo table in hTKKCMCSheetInfo.Tables)
				{
					if (table.RowsValues.Count == 0)
					{
						List<string> list = new List<string>();
						HTKKCMCRowInfo rowsInfo = table.RowsInfo;
						for (int i = 0; i < rowsInfo.Cells.Count; i++)
						{
							list.Add(rowsInfo.Cells[i].DefaultValue);
						}
						table.RowsValues.Add(new HTKKCMCRowValue
						{
							RowValue = list,
							RowDisplay = "1"
						});
					}
					if (table.RowsValues.Count > 1)
					{
						WinSpreadProcess.InsertTableRowData(winFpSpread.Sheets[iActiveSheet], table.Location.DataLocationTemplateRowStart, -1, table.Location.DataLocationTemplateRowStart, -1, table.RowsValues.Count - 1, -1);
						hTKKCMCSheetInfo.IncrementStaticCell(table.Location.DataLocationTemplateRowStart, table.RowsValues.Count - 1);
						foreach (HTKKCMCTableInfo table2 in hTKKCMCSheetInfo.Tables)
						{
							if (table2.Location.DataLocationTemplateRowStart > table.Location.DataLocationTemplateRowStart)
							{
								table2.Increment(table.RowsValues.Count - 1);
							}
						}
						for (int j = 0; j < table.RowsValues.Count; j++)
						{
							WinSpreadProcess.SetValueRow(winFpSpread.Sheets[iActiveSheet], table, j, table.RowsValues[j]);
						}
						WinSpreadProcess.fpSpReindexingTableTemplate(winFpSpread.Sheets[iActiveSheet], hTKKCMCSheetInfo.Tables[num14]);
					}
					else
					{
						WinSpreadProcess.SetValueRow(winFpSpread.Sheets[iActiveSheet], table, 0, table.RowsValues[0]);
					}
					num14++;
				}
				isClickTongHopKHBS = true;
				isDataChanged = true;
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0222", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
			else
			{
				int sheetIndex = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
				SheetView sView = winFpSpread.Sheets[sheetIndex];
				HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
				int sheetIndex2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
				SheetView sheetView = winFpSpread.Sheets[sheetIndex2];
				HTKKCMCSheetInfo sInfo2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex2];
				TongHopKhbsMoi();
				TinhSoThuePhaiNopKhauTruHoan_01_KHBS();
				dicTieuMucKHBS = WinSpreadProcess.GetDMTieuMucKHBS(winFpSpread);
				CollectionDynamicPhanBo_01_KHBS();
				SuaCotHuyenPhanBo_PL_01_KHBS(sheetView, _mapMct);
				AnHienSoThueThuHoiHoanTienNopCham_01_KHBS();
				TinhSoNgayNopCham();
				TinhSoTienPhatNopCham();
				ChangeCellTypeNhieuLoaiTien(sView, sInfo);
				ChangeCellTypeNhieuLoaiTien(sheetView, sInfo2);
				AnHienSoNgayVaSoTienNopCham_01_KHBS();
				GeneralFunctions.SaveAllCellTypeOfSheet(sView, sInfo, sheetIndex);
				GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, sInfo2, sheetIndex2);
				SetHeightOfTable_01_KHBS(WinSettings.ToKhaiInfo.SheetInfos.Count - 2);
				isClickTongHopKHBS = true;
				isDataChangedCallTHKhbs = false;
				winFpSpread.ResumeLayout();
				frmMessage frmMessage = new frmMessage("0222", "3", "", "Thông báo");
				((Form)frmMessage).ShowDialog();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void Default_01_1_KHBS()
	{
		try
		{
			if (WinSettings.ToKhaiInfo.KieuKHBS == "3" && WinSettings.TTPeriodInfo.LoaiToKhai == "B")
			{
				int num = HTKKFile.ColumnLettersToIndex("B");
				int num2 = HTKKFile.ColumnLettersToIndex("H");
				int num3 = HTKKFile.ColumnLettersToIndex("P");
				int num4 = HTKKFile.ColumnLettersToIndex("BS");
				int num5 = HTKKFile.ColumnLettersToIndex("AD");
				int num6 = HTKKFile.ColumnLettersToIndex("AH");
				int num7 = HTKKFile.ColumnLettersToIndex("BV");
				int num8 = HTKKFile.ColumnLettersToIndex("BT");
				int sheetIndex = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
				SheetView sheetView = winFpSpread.Sheets[sheetIndex];
				HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
				int sheetIndex2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
				SheetView sheetView2 = winFpSpread.Sheets[sheetIndex2];
				HTKKCMCSheetInfo sInfo2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex2];
				dicToKhaiPhuLucKHBS = WinSpreadProcess.GetDMToKhaiPhuLucKHBS();
				dicTieuMucKHBS = WinSpreadProcess.GetDMTieuMucKHBS(winFpSpread);
				if (WinSettings.ToKhaiInfo.SheetInfos.Count <= 3)
				{
					((Control)btnThemPhuLuc).Visible = false;
				}
				else if (WinSettings.MenuID == "B9" || WinSettings.MenuID == "B8" || WinSettings.ToKhaiInfo.MenuID.Equals("W1"))
				{
					((Control)btnThemPhuLuc).Visible = false;
				}
				else
				{
					((Control)btnThemPhuLuc).Visible = true;
				}
				_mapMct = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
				_mapMctTieuMuc = WinFnProcess.GetMCT01("MapMCT_TieuMuc.xml", WinSettings.ToKhaiInfo.MenuID);
				UpdateXmlFileMapMCTSpecial(ref _mapMct, ref _mapMctTieuMuc);
				SetHeaderKHBSMoi(sheetView, sheetView2);
				SetHeightOfRow_01_KHBS();
				LockUnlock_01_1_KHBS();
				SuaCotHuyenPhanBo_PL_01_KHBS(sheetView2, _mapMct);
				ReSetControltypeAndConfig();
				SetDefaultLoaiTienStatic();
				ChangeCellTypeNhieuLoaiTien(sheetView, sInfo);
				ChangeCellTypeNhieuLoaiTien(sheetView2, sInfo2);
				AnHienSoThueThuHoiHoanTienNopCham_01_KHBS();
				TinhSoNgayNhanDuocTienHoanThue_01_KHBS(sheetView2, is_edit: false);
				TinhSoTienChamNop_01_KHBS(sheetView2, is_edit: false);
				AnHienSoNgayVaSoTienNopCham_01_KHBS();
				GeneralFunctions.SaveAllCellTypeOfSheet(sheetView, sInfo, sheetIndex);
				GeneralFunctions.SaveAllCellTypeOfSheet(sheetView2, sInfo2, sheetIndex2);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetHeightOfRow_01_KHBS()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sView = winFpSpread.Sheets[num];
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num].Tables;
			HTKKCMCTableInfo hTKKCMCTableInfo = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/ChiTiet");
			HTKKCMCTableInfo hTKKCMCTableInfo2 = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_II/ChiTiet");
			HTKKCMCTableInfo hTKKCMCTableInfo3 = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_III/ChiTiet");
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num2 = 0; num2 < hTKKCMCTableInfo.RowsValues.Count; num2++)
			{
				GeneralFunctions.SetRowHeight(sView, dataLocationTemplateRowStart + num2, 30f);
			}
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int num2 = 0; num2 < hTKKCMCTableInfo2.RowsValues.Count; num2++)
			{
				GeneralFunctions.SetRowHeight(sView, dataLocationTemplateRowStart2 + num2, 30f);
			}
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			for (int num2 = 0; num2 < hTKKCMCTableInfo3.RowsValues.Count; num2++)
			{
				GeneralFunctions.SetRowHeight(sView, dataLocationTemplateRowStart3 + num2, 30f);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetHeaderKHBSMoi(SheetView sView_01_1_KHBS, SheetView sView_01_KHBS)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("O_6");
			int[] array2 = GeneralFunctions.ArrayCellId("O_10");
			int[] array3 = GeneralFunctions.ArrayCellId("O_12");
			int[] array4 = GeneralFunctions.ArrayCellId("O_14");
			int[] array5 = GeneralFunctions.ArrayCellId("O_16");
			int[] array6 = GeneralFunctions.ArrayCellId("M_74");
			int[] array7 = GeneralFunctions.ArrayCellId("M_76");
			int[] array8 = GeneralFunctions.ArrayCellId("X_74");
			int[] array9 = GeneralFunctions.ArrayCellId("X_76");
			int[] array10 = GeneralFunctions.ArrayCellId("L_6");
			int[] array11 = GeneralFunctions.ArrayCellId("L_8");
			int[] array12 = GeneralFunctions.ArrayCellId("L_10");
			int[] array13 = GeneralFunctions.ArrayCellId("L_12");
			int[] array14 = GeneralFunctions.ArrayCellId("L_14");
			int[] array15 = GeneralFunctions.ArrayCellId("L_16");
			int[] array16 = GeneralFunctions.ArrayCellId("L_18");
			int[] array17 = GeneralFunctions.ArrayCellId("L_20");
			int[] array18 = GeneralFunctions.ArrayCellId("L_22");
			int[] array19 = GeneralFunctions.ArrayCellId("L_24");
			int[] array20 = GeneralFunctions.ArrayCellId("J_137");
			int[] array21 = GeneralFunctions.ArrayCellId("J_139");
			int[] array22 = GeneralFunctions.ArrayCellId("U_137");
			int[] array23 = GeneralFunctions.ArrayCellId("U_139");
			string value = "";
			if (WinSettings.TTinChung.kieuKy == "Y")
			{
				value = ((!string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuThang)) ? ("Năm " + WinSettings.TTinChung.nam + " (Từ tháng " + WinSettings.TTinChung.kyKKhaiTuThang + " Đến tháng " + WinSettings.TTinChung.kyKKhaiDenThang + ")") : (string.IsNullOrEmpty(WinSettings.TTinChung.kyKKhaiTuNgay) ? ("Năm " + WinSettings.TTinChung.nam) : ("Năm " + WinSettings.TTinChung.nam + " (Từ ngày " + WinSettings.TTinChung.kyKKhaiTuNgay + " Đến ngày " + WinSettings.TTinChung.kyKKhaiDenNgay + ")")));
			}
			if (WinSettings.TTinChung.kieuKy == "M")
			{
				value = "Tháng " + WinSettings.TTinChung.thang + " " + "Năm".ToLower() + " " + WinSettings.TTinChung.nam;
			}
			if (WinSettings.TTinChung.kieuKy == "Q")
			{
				value = "Quý " + WinSettings.TTinChung.quy + " " + "Năm".ToLower() + " " + WinSettings.TTinChung.nam;
			}
			if (WinSettings.TTinChung.kieuKy == "D")
			{
				value = "Từng lần phát sinh ngày " + WinSettings.TTinChung.ngay + " " + "Tháng".ToLower() + " " + WinSettings.TTinChung.thang + " " + "Năm".ToLower() + " " + WinSettings.TTinChung.nam;
			}
			if (WinSettings.TTinChung.kieuKy == "D_D")
			{
				value = "Từ ngày " + WinSettings.TTinChung.kyKKhaiTuNgay + " Đến ngày " + WinSettings.TTinChung.kyKKhaiDenNgay;
			}
			if (WinSettings.TTinChung.kieuKy == "K")
			{
				value = "Từ ngày " + WinSettings.TTinChung.kyKKhaiTuNgay + " Đến ngày " + WinSettings.TTinChung.kyKKhaiDenNgay;
			}
			sView_01_1_KHBS.Cells[array[1], array[0]].Value = ((_mapMct != null && _mapMct.Attributes["MauToKhai"] != null) ? ((XmlNode)_mapMct.Attributes["MauToKhai"]).Value : "");
			sView_01_1_KHBS.Cells[array2[1], array2[0]].Value = value;
			sView_01_1_KHBS.Cells[array3[1], array3[0]].Value = WinSettings.TTinChung.soLan;
			sView_01_1_KHBS.Cells[array4[1], array4[0]].Value = WinSettings.TTinChung.tenNNT;
			sView_01_1_KHBS.Cells[array5[1], array5[0]].Value = WinSettings.TTinChung.mst;
			sView_01_KHBS.Cells[array10[1], array10[0]].Value = ((_mapMct != null && _mapMct.Attributes["MauToKhai"] != null) ? ((XmlNode)_mapMct.Attributes["MauToKhai"]).Value : "");
			sView_01_KHBS.Cells[array12[1], array12[0]].Value = value;
			sView_01_KHBS.Cells[array13[1], array13[0]].Value = WinSettings.TTinChung.soLan;
			sView_01_KHBS.Cells[array14[1], array14[0]].Value = WinSettings.TTinChung.tenNNT;
			sView_01_KHBS.Cells[array15[1], array15[0]].Value = WinSettings.TTinChung.mst;
			sView_01_KHBS.Cells[array16[1], array16[0]].Value = WinSettings.TTinChung.tenDLyThue;
			sView_01_KHBS.Cells[array17[1], array17[0]].Value = WinSettings.TTinChung.mstDLyThue;
			sView_01_KHBS.Cells[array18[1], array18[0]].Value = WinSettings.TTinChung.soHDongDLyThue;
			sView_01_KHBS.Cells[array19[1], array19[0]].Value = WinSettings.TTinChung.ngayKyHDDLyThue;
			SetThongTinChungNguoiKyNgayKyKHBS();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetThongTinChungNguoiKyNgayKyKHBS()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView2 = winFpSpread.Sheets[num2];
			SheetView sheetView3 = winFpSpread.Sheets[0];
			string pathTenNVienDLyThue = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue/NVienDLy/tenNVienDLyThue";
			string pathNguoiKy = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/nguoiKy";
			string pathChungChiHNghe = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/DLyThue/NVienDLy/cchiHNghe";
			string pathNgayKy = "HSoThueDTu/HSoKhaiThue/TTinChung/TTinTKhaiThue/TKhaiThue/ngayKy";
			HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathTenNVienDLyThue));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = WinSettings.ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNguoiKy));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = WinSettings.ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathChungChiHNghe));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = WinSettings.ToKhaiInfo.SheetInfos[0].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathNgayKy));
			int[] array = GeneralFunctions.ArrayCellId("N_18");
			int[] array2 = GeneralFunctions.ArrayCellId("N_20");
			int[] array3 = GeneralFunctions.ArrayCellId("R_18");
			int[] array4 = GeneralFunctions.ArrayCellId("R_20");
			int[] array5 = GeneralFunctions.ArrayCellId("O_39");
			int[] array6 = GeneralFunctions.ArrayCellId("O_40");
			int[] array7 = GeneralFunctions.ArrayCellId("U_39");
			int[] array8 = GeneralFunctions.ArrayCellId("U_40");
			string value = ((hTKKCMCCellInfo == null) ? "" : sheetView3.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim());
			string text = ((hTKKCMCCellInfo2 == null) ? "" : sheetView3.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim());
			string value2 = ((hTKKCMCCellInfo3 == null) ? "" : sheetView3.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim());
			string value3 = ((hTKKCMCCellInfo4 == null) ? "" : sheetView3.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Text.Trim());
			if (string.IsNullOrEmpty(WinSettings.HTKKHeaderInfo.nguoiKy) && !string.IsNullOrEmpty(text))
			{
				WinSettings.HTKKHeaderInfo.nguoiKy = text;
				WinSettings.TTinChung.nguoiKy = text;
			}
			sheetView.Cells[array[1], array[0]].Value = value;
			sheetView.Cells[array2[1], array2[0]].Value = value2;
			sheetView.Cells[array3[1], array3[0]].Value = text;
			sheetView.Cells[array4[1], array4[0]].Value = value3;
			sheetView2.Cells[array5[1], array5[0]].Value = value;
			sheetView2.Cells[array6[1], array6[0]].Value = value2;
			sheetView2.Cells[array7[1], array7[0]].Value = text;
			sheetView2.Cells[array8[1], array8[0]].Value = value3;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void TongHopKhbsMoi()
	{
		//IL_03ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_03f4: Expected O, but got Unknown
		try
		{
			int colStt = HTKKFile.ColumnLettersToIndex("B");
			int colCt = HTKKFile.ColumnLettersToIndex("H");
			int colSheetIndex_01_1_KHBS = HTKKFile.ColumnLettersToIndex("BS");
			int num = HTKKFile.ColumnLettersToIndex("AD");
			int num2 = HTKKFile.ColumnLettersToIndex("AH");
			int num3 = HTKKFile.ColumnLettersToIndex("BV");
			int num4 = HTKKFile.ColumnLettersToIndex("BN");
			int colMaHoSo = HTKKFile.ColumnLettersToIndex("BT");
			int num5 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num5];
			SheetView sheetView = winFpSpread.Sheets[num5];
			int num6 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView2 = winFpSpread.Sheets[num6];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = WinSettings.ToKhaiInfo.SheetInfos[num6];
			if (WinSettings.MenuID == "O1")
			{
				int num7 = GeneralFunctions.SoNgayNopCham(WinSettings.TTPeriodInfo, WinSettings.HTKKHeaderInfo);
				int[] array = GeneralFunctions.ArrayCellId("AA_36");
				int num8 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
				sheetView2.Cells[array[1], array[0]].Value = num7;
			}
			int num9 = ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? (hTKKCMCSheetInfo.Tables.Count - 4) : (hTKKCMCSheetInfo.Tables.Count - 1));
			for (int num10 = num9 - 1; num10 >= 0; num10--)
			{
				if (num10 < hTKKCMCSheetInfo.Tables.Count - 1 && hTKKCMCSheetInfo.Tables[num10] != null)
				{
					int tableTemplateRowStart = hTKKCMCSheetInfo.Tables[num10].Location.TableTemplateRowStart;
					GeneralFunctions.RemoveTableF8New(sheetView, hTKKCMCSheetInfo, num10, setActiveCell: true);
				}
			}
			if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count > 1)
			{
				int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
				GeneralFunctions.RemoveDataRows(sheetView, 0, dataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[0].RowsValues.Count - 1, hTKKCMCSheetInfo);
			}
			if (hTKKCMCSheetInfo.Tables[0].RowsValues.Count > 0)
			{
				int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[0].Location.DataLocationTemplateRowStart;
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCSheetInfo.Tables[0].RowsInfo.Cells)
				{
					sheetView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
			dicToKhaiPhuLucKHBS = WinSpreadProcess.GetDMToKhaiPhuLucKHBS();
			dicTieuMucKHBS = WinSpreadProcess.GetDMTieuMucKHBS(winFpSpread);
			SetDefaultLoaiTienStatic();
			SetThongTinChungNguoiKyNgayKyKHBS();
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num5].Tables;
			int num11 = 0;
			int num12 = 0;
			DanhDauNeuCoPhanBo(sheetView, _mapMct);
			string loaiThue = "";
			foreach (KeyValuePair<string, string> dicToKhaiPhuLucKHB in dicToKhaiPhuLucKHBS)
			{
				num12 = Convert.ToInt32(dicToKhaiPhuLucKHB.Key);
				HTKKCMCSheetInfo kHBS_LanTruoc = WinFnProcess.ReadSheet_LanTruoc(num12);
				int nextRow = 0;
				if (num11 > 0)
				{
					GeneralFunctions.InsertTableF7(sheetView, hTKKCMCSheetInfo, num11 - 1);
				}
				SetSheetNameToTableHeader(sheetView, tables[num11], dicToKhaiPhuLucKHB, colCt, colSheetIndex_01_1_KHBS, colStt, num11 + 1, colMaHoSo);
				foreach (XmlNode childNode in _mapMct.ChildNodes)
				{
					XmlNode val = childNode;
					int num13 = ((val.Attributes["Dynamic"] != null) ? Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value) : (-1));
					int num14 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num5 : num6));
					int num15 = ((val.Attributes["SheetIndex"] != null) ? Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value) : (-1));
					string text = ((val.Attributes["LoaiThue"] != null) ? ((XmlNode)val.Attributes["LoaiThue"]).Value : "");
					string value = ((val.Attributes["MCT"] != null) ? ((XmlNode)val.Attributes["MCT"]).Value.ToString().Trim() : "");
					string text2 = ((val.Attributes["MCTPath"] != null) ? ((XmlNode)val.Attributes["MCTPath"]).Value.ToString().Trim() : "");
					string text3 = ((val.Attributes["Manual"] != null) ? ((XmlNode)val.Attributes["Manual"]).Value.ToString().Trim() : "");
					string value2 = ((val.Attributes["TenTinhPath"] != null) ? ((XmlNode)val.Attributes["TenTinhPath"]).Value.ToString().Trim() : "");
					if (!string.IsNullOrEmpty(value2))
					{
						continue;
					}
					if (!string.IsNullOrEmpty(text))
					{
						loaiThue = text;
					}
					if (num13 >= 0 && num14 >= 0 && !string.IsNullOrEmpty(value) && !(text3 == "1") && num12 == num15)
					{
						if (num13 == 1 && num14 == num5)
						{
							nextRow = CollectionDynamicHeader01_1_KHBS(val, num11, nextRow, kHBS_LanTruoc);
							nextRow = CollectionDynamic01_1_KHBS(val, num11, nextRow, kHBS_LanTruoc);
						}
						if (num13 == 0 && num14 == num5)
						{
							nextRow = CollectionStatic01_1_KHBS(val, num11, nextRow, kHBS_LanTruoc);
						}
					}
				}
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num5].Tables[num11]);
				num11++;
			}
			CheckNeuNhapThuCong_01_1_KHBS();
			IEnumerable<HTKKCMCTableInfo> source = WinSettings.ToKhaiInfo.SheetInfos[num5].Tables.Where((HTKKCMCTableInfo x) => x.TablePath == "PL01_1_KHBS/Muc_A/DSachHSo/BKeHSo");
			RemoveTableF7IfNull(sheetView, hTKKCMCSheetInfo, source.ToList());
			GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num5]);
			LockUnlock_01_1_KHBS();
			TinhTongCt7_8_9(loaiThue);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void DanhDauNeuCoPhanBo(SheetView sView_01_1_KHBS, XmlNode _mapMct)
	{
		try
		{
			int[] array = GeneralFunctions.ArrayCellId("BP_38");
			string text = ((_mapMct.Attributes["CoPhanBo"] != null) ? ((XmlNode)_mapMct.Attributes["CoPhanBo"]).Value.ToString().Trim() : "");
			sView_01_1_KHBS.Cells[array[1], array[0]].Value = ((!string.IsNullOrEmpty(text) && text == "1") ? 1 : 0);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckNeuNhapThuCong_01_1_KHBS()
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("BN");
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num2].Tables;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num2];
			SheetView sView = winFpSpread.Sheets[num2];
			int num3 = 0;
			foreach (KeyValuePair<string, string> dicToKhaiPhuLucKHB in dicToKhaiPhuLucKHBS)
			{
				Dictionary<string, string> dMMaChiTieuTheoPhuLucKHBS = GetDMMaChiTieuTheoPhuLucKHBS(int.Parse(dicToKhaiPhuLucKHB.Key));
				HTKKCMCTableInfo hTKKCMCTableInfo = tables[num3];
				if (dMMaChiTieuTheoPhuLucKHBS.Count > 0 && hTKKCMCTableInfo.RowsValues.Count > 0)
				{
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
					{
						SetDefaultManualRow_01_1_KHBS(sView, num3, dataLocationTemplateRowStart + i);
					}
					InsertComboboxCt3_01_1_KHBS(num3);
				}
				num3++;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void DeletePhuLucRong_01_1_KHBS(SheetView sView, List<HTKKCMCSheetInfo> sInfos, int sheetIndex)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("H");
			List<HTKKCMCTableInfo> tables = sInfos[sheetIndex].Tables;
			int num2 = ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? (tables.Count - 3) : tables.Count);
			for (int num3 = num2 - 1; num3 >= 0; num3--)
			{
				if (tables[num3].RowsValues.Count > 0)
				{
					int dataLocationTemplateRowStart = tables[num3].Location.DataLocationTemplateRowStart;
					string value = sView.Cells[dataLocationTemplateRowStart - 1, num - 1].Text.Trim();
					if (string.IsNullOrEmpty(value))
					{
						GeneralFunctions.DeleteTableF8(sView, sInfos[sheetIndex], num3, dataLocationTemplateRowStart - 1, 0);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetHeightOfTable_01_KHBS(int sheetIndex)
	{
		try
		{
			TextCellType textCellType = new TextCellType();
			SheetView sheetView = winFpSpread.Sheets[sheetIndex];
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables;
			foreach (HTKKCMCTableInfo item in tables)
			{
				int num3 = item.Location.DataLocationTemplateRowStart - 1;
				for (int i = 0; i < item.RowsValues.Count; i++)
				{
					float preferredRowHeight = sheetView.GetPreferredRowHeight(num3 + i);
					if (item.TableName == "KHBS_01_TABLE_2")
					{
						GeneralFunctions.SetRowHeight(sheetView, num3 + i + 1, preferredRowHeight);
					}
					else
					{
						GeneralFunctions.SetRowHeight(sheetView, num3 + i, preferredRowHeight);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhTongCt7_8_9(string loaiThue)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				int num2 = HTKKFile.ColumnLettersToIndex("AD");
				int num3 = HTKKFile.ColumnLettersToIndex("AH");
				int num4 = HTKKFile.ColumnLettersToIndex("CA");
				int num5 = HTKKFile.ColumnLettersToIndex("BV");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(num5, num5);
				Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
				dictionary2.Add(num3, num3);
				dictionary2.Add(num4, num4);
				dictionary2.Add(num2, num2);
				int num6 = -1;
				if (loaiThue == "0")
				{
					num6 = hTKKCMCSheetInfo.Tables.Count - 4;
					GeneralFunctions.CopyTableToTableGroup_F7_F5(sView, hTKKCMCSheetInfo, "KHBS_01_1_TABLE_1", "KHBS_01_1_TABLE_2", dictionary, dictionary2, usingCellType: true);
				}
				if (loaiThue == "1")
				{
					num6 = hTKKCMCSheetInfo.Tables.Count - 3;
					GeneralFunctions.CopyTableToTableGroup_F7_F5(sView, hTKKCMCSheetInfo, "KHBS_01_1_TABLE_1", "KHBS_01_1_TABLE_3", dictionary, dictionary2, usingCellType: true);
				}
				if (loaiThue == "2")
				{
					num6 = hTKKCMCSheetInfo.Tables.Count - 2;
					GeneralFunctions.CopyTableToTableGroup_F7_F5(sView, hTKKCMCSheetInfo, "KHBS_01_1_TABLE_1", "KHBS_01_1_TABLE_4", dictionary, dictionary2, usingCellType: true);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetSheetNameToTableHeader(SheetView sView, HTKKCMCTableInfo table, KeyValuePair<string, string> header, int colCt2, int colSheetIndex_01_1_KHBS, int colStt, int stt, int colMaHoSo)
	{
		try
		{
			int rowIndex = table.Header.StaticCells[1].CellTemplate.RowIndex;
			sView.Cells[rowIndex - 1, colStt - 1].Text = stt.ToString();
			sView.Cells[rowIndex - 1, colCt2 - 1].Text = header.Value;
			sView.Cells[rowIndex - 1, colSheetIndex_01_1_KHBS - 1].Text = header.Key;
			sView.Cells[rowIndex - 1, colMaHoSo - 1].Text = GetMaHoSo_01_1_KHBS(Convert.ToInt32(header.Key), header.Value);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public string GetMaHoSo_01_1_KHBS(int sheetIndex, string sheetName)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		try
		{
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			if (val == null)
			{
				return "";
			}
			XmlNode val3 = ((IEnumerable)val).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["SheetName"] != null && ((XmlNode)x.Attributes["SheetName"]).Value == sheetName));
			if (val3 != null && val3.Attributes["MaHoSo"] != null)
			{
				return ((XmlNode)val3.Attributes["MaHoSo"]).Value;
			}
			if (sheetIndex == 0)
			{
				return (val.Count > 0 && val[0].Attributes["MaHoSo"] != null) ? ((XmlNode)val[0].Attributes["MaHoSo"]).Value : "";
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	public string GetSheetNameOrSheetIndexFromMaHoSo_01_1_KHBS(string maHoSo)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_001c: Expected O, but got Unknown
		try
		{
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			if (val == null)
			{
				return "";
			}
			XmlNode val3 = ((IEnumerable)val).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MaHoSo"] != null && ((XmlNode)x.Attributes["MaHoSo"]).Value == maHoSo));
			if (val3 != null && val3.Attributes["SheetName"] != null)
			{
				for (int num = 0; num < val.Count; num++)
				{
					if (((XmlNode)val3.Attributes["SheetName"]).Value == ((XmlNode)val[0].Attributes["SheetName"]).Value)
					{
						return "0";
					}
				}
				return ((XmlNode)val3.Attributes["SheetName"]).Value;
			}
			return "";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public void UpdateCtForDoiTac(string loaiChiTieu, int row_01_1_KHBS)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("B");
			int num2 = HTKKFile.ColumnLettersToIndex("H");
			int num3 = HTKKFile.ColumnLettersToIndex("BY");
			int num4 = HTKKFile.ColumnLettersToIndex("P");
			int num5 = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("V");
			int num7 = HTKKFile.ColumnLettersToIndex("Z");
			int num8 = HTKKFile.ColumnLettersToIndex("AD");
			int num9 = HTKKFile.ColumnLettersToIndex("AH");
			int num10 = HTKKFile.ColumnLettersToIndex("CA");
			int num11 = HTKKFile.ColumnLettersToIndex("BR");
			int num12 = HTKKFile.ColumnLettersToIndex("BS");
			int num13 = HTKKFile.ColumnLettersToIndex("BT");
			int num14 = HTKKFile.ColumnLettersToIndex("G");
			int num15 = HTKKFile.ColumnLettersToIndex("BV");
			int num16 = HTKKFile.ColumnLettersToIndex("BX");
			int num17 = HTKKFile.ColumnLettersToIndex("BW");
			int num18 = HTKKFile.ColumnLettersToIndex("BD");
			int num19 = HTKKFile.ColumnLettersToIndex("BN");
			int num20 = HTKKFile.ColumnLettersToIndex("BO");
			int num21 = HTKKFile.ColumnLettersToIndex("BM");
			int num22 = HTKKFile.ColumnLettersToIndex("BE");
			int num23 = HTKKFile.ColumnLettersToIndex("BF");
			int num24 = HTKKFile.ColumnLettersToIndex("BG");
			int num25 = HTKKFile.ColumnLettersToIndex("BH");
			int num26 = HTKKFile.ColumnLettersToIndex("BI");
			int num27 = HTKKFile.ColumnLettersToIndex("BJ");
			int num28 = HTKKFile.ColumnLettersToIndex("BK");
			int num29 = HTKKFile.ColumnLettersToIndex("BL");
			int num30 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num30];
			sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Value = "";
			sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Value = "";
			sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Value = "";
			sheetView.Cells[row_01_1_KHBS - 1, num27 - 1].Value = "";
			sheetView.Cells[row_01_1_KHBS - 1, num28 - 1].Value = "";
			string text = sheetView.Cells[row_01_1_KHBS - 1, num2 - 1].Text.ToString().Trim();
			object value = sheetView.Cells[row_01_1_KHBS - 1, num29 - 1].Value;
			string text2 = ((value == null) ? "" : value.ToString().Trim());
			if (text2 != null && !string.IsNullOrEmpty(text2) && text.StartsWith(text2))
			{
				text = text.Substring(text2.Length + " - ".Length, text.Length - (" - ".Length + text2.Length));
			}
			if (loaiChiTieu == "1")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Value = text;
				string text3 = sheetView.Cells[row_01_1_KHBS - 1, num16 - 1].Text.Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value = text3;
				string text4 = sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Text.Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value = GeneralFunctions.ReplaceFromStart(text4, text3 + " - ");
				sheetView.Cells[row_01_1_KHBS - 1, num29 - 1].Value = text2;
			}
			if (loaiChiTieu == "2")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num16 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num2 - 1].Value;
			}
			if (loaiChiTieu == "3")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num27 - 1].Value = text;
				string text3 = sheetView.Cells[row_01_1_KHBS - 1, num16 - 1].Text.Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value = text3;
				string text4 = sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Text.Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value = GeneralFunctions.ReplaceFromStart(text4, text3 + " - ");
				sheetView.Cells[row_01_1_KHBS - 1, num29 - 1].Value = text2;
			}
			if (loaiChiTieu == "4")
			{
				string text3 = sheetView.Cells[row_01_1_KHBS - 1, num16 - 1].Text.Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value = text3;
				string text4 = sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Text.Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value = GeneralFunctions.ReplaceFromStart(text4, text3 + " - ");
				sheetView.Cells[row_01_1_KHBS - 1, num29 - 1].Value = GeneralFunctions.FormatKey2(text2, text, " - ");
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private int CollectionDynamic01_1_KHBS(XmlNode mapCt, int ToTableindex, int nextRow, HTKKCMCSheetInfo KHBS_LanTruoc)
	{
		string text = "";
		string text2 = "";
		try
		{
			if (mapCt == null)
			{
				return nextRow;
			}
			int sheetIndex = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			text = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string tablePath = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			text2 = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string text3 = ((mapCt.Attributes["TCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim());
			string maTCTPath = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string text4 = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string text5 = ((mapCt.Attributes["Caption"] == null) ? "" : ((XmlNode)mapCt.Attributes["Caption"]).Value.ToString().Trim());
			string DVTPath = ((mapCt.Attributes["DVTPath"] != null) ? ((XmlNode)mapCt.Attributes["DVTPath"]).Value.ToString().Trim() : "");
			string maDVTPath = ((mapCt.Attributes["MaDVTPath"] != null) ? ((XmlNode)mapCt.Attributes["MaDVTPath"]).Value.ToString().Trim() : "");
			string text6 = ((mapCt.Attributes["LoaiChiTieu"] != null) ? ((XmlNode)mapCt.Attributes["LoaiChiTieu"]).Value.ToString().Trim() : "");
			string text7 = ((mapCt.Attributes["Manual"] != null) ? ((XmlNode)mapCt.Attributes["Manual"]).Value.ToString().Trim() : "");
			string prefixPath = ((mapCt.Attributes["PrefixPath"] != null) ? ((XmlNode)mapCt.Attributes["PrefixPath"]).Value.ToString().Trim() : "");
			string text8 = ((mapCt.Attributes["Prefix"] != null) ? ((XmlNode)mapCt.Attributes["Prefix"]).Value.ToString().Trim() : "");
			int num = ((mapCt.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("BY");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("Z");
			int col = HTKKFile.ColumnLettersToIndex("AD");
			int num9 = HTKKFile.ColumnLettersToIndex("AH");
			int num10 = HTKKFile.ColumnLettersToIndex("CA");
			int num11 = HTKKFile.ColumnLettersToIndex("BR");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("BS");
			int colSheetName = HTKKFile.ColumnLettersToIndex("BT");
			int num12 = HTKKFile.ColumnLettersToIndex("G");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("BV");
			int num13 = HTKKFile.ColumnLettersToIndex("BX");
			int num14 = HTKKFile.ColumnLettersToIndex("BW");
			int num15 = HTKKFile.ColumnLettersToIndex("BD");
			int num16 = HTKKFile.ColumnLettersToIndex("BN");
			int num17 = HTKKFile.ColumnLettersToIndex("BO");
			int num18 = HTKKFile.ColumnLettersToIndex("BM");
			int num19 = HTKKFile.ColumnLettersToIndex("BQ");
			int num20 = HTKKFile.ColumnLettersToIndex("BL");
			int num21 = HTKKFile.ColumnLettersToIndex("CE");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			List<HTKKCMCTableInfo> list = KHBS_LanTruoc?.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
			SheetView sheetView = ((hTKKCMCSheetInfo == null) ? null : winFpSpread.Sheets[sheetIndex]);
			List<HTKKCMCTableInfo> list2 = ((hTKKCMCSheetInfo == null) ? new List<HTKKCMCTableInfo>() : hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList());
			int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView2 = winFpSpread.Sheets[sheetIndex_KHBS];
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS].Tables[ToTableindex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			ICellType cellType = sheetView2.Cells[0, 0].CellType;
			for (int num22 = 0; num22 < list2.Count(); num22++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo2 = list2[num22];
				int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				List<HTKKCMCTableCellInfo> list3 = new List<HTKKCMCTableCellInfo>();
				List<int> list4 = new List<int>();
				List<string> list5 = GeneralFunctions.SplitString(text3, "###");
				foreach (string path in list5)
				{
					List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo2.RowsInfo.Cells;
					Func<HTKKCMCTableCellInfo, bool> predicate = (HTKKCMCTableCellInfo x) => x.Path == path;
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = cells.FirstOrDefault(predicate);
					int item = hTKKCMCTableInfo2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == path);
					list3.Add(hTKKCMCTableCellInfo);
					list4.Add(item);
					if (hTKKCMCTableCellInfo == null && hTKKCMCTableInfo2.Header.StaticCells.Count == 0)
					{
						HTKKFile.Log("xem lai path TCT: " + path);
					}
				}
				if (list3.Count == 0 || (list3.Count > 0 && list3[0] == null))
				{
					continue;
				}
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = hTKKCMCTableInfo2.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
				int num23 = hTKKCMCTableInfo2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo3 = hTKKCMCTableInfo2.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
				int num24 = hTKKCMCTableInfo2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo4 = hTKKCMCTableInfo2.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
				int num25 = hTKKCMCTableInfo2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo5 = hTKKCMCTableInfo2.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
				int num26 = hTKKCMCTableInfo2.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo6 = hTKKCMCTableInfo2.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
				if (hTKKCMCTableCellInfo3 == null)
				{
					HTKKFile.Log("xem lai path MCT: " + MCTPath);
				}
				TextCellType textCellType = new TextCellType();
				for (int num27 = 0; num27 < hTKKCMCTableInfo2.RowsValues.Count; num27++)
				{
					if (sheetView == null || CheckTruongHopDacBietPhanBo_01_KHBSS_InView(sheetView, sheetIndex, dataLocationTemplateRowStart2 + num27))
					{
						continue;
					}
					List<string> list6 = new List<string>();
					foreach (HTKKCMCTableCellInfo item2 in list3)
					{
						string text9 = "";
						if (item2 != null)
						{
							object value = sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, item2.CellTemplate.ColIndex - 1].Value;
							text9 = FormatValueDieuChinh(item2, value);
						}
						list6.Add(text9.Trim());
					}
					string text10 = ((hTKKCMCTableCellInfo2 != null) ? Convert.ToString(sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Value).Trim() : "");
					text10 = text10.Trim();
					if (GeneralFunctions.CheckInListEmpty(list6))
					{
						continue;
					}
					object obj = null;
					if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCTableCellInfo6 != null)
					{
						obj = sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo6.CellTemplate.ColIndex - 1].Value;
						if (obj != null)
						{
							if (!list5.Contains(prefixPath))
							{
								list5.Add(prefixPath);
							}
							list6.Add(obj.ToString().Trim());
						}
					}
					string value2 = "";
					List<HTKKCMCTableInfo> list7 = ((KHBS_LanTruoc == null) ? null : GeneralFunctions.GetTableByKeyBody(KHBS_LanTruoc.Tables, list5, list6, hTKKCMCTableInfo2.TablePath));
					HTKKCMCTableInfo hTKKCMCTableInfo3 = ((list7 != null && list7.Count > 0) ? list7[0] : null);
					int num28 = hTKKCMCTableInfo3?.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MCTPath) ?? (-1);
					if (hTKKCMCTableInfo3 != null)
					{
						List<HTKKCMCRowValue> rowValueInTable = GeneralFunctions.GetRowValueInTable(hTKKCMCTableInfo3, list5, list6);
						value2 = ((rowValueInTable != null && num28 >= 0) ? rowValueInTable[0].RowValue[num28] : "");
						if (hTKKCMCTableInfo3 != null && num28 >= 0)
						{
							value2 = FormatValueLanTruoc(hTKKCMCTableInfo3.RowsInfo.Cells[num28], value2);
						}
						else if (hTKKCMCTableCellInfo3 != null)
						{
							value2 = FormatValueLanTruoc(hTKKCMCTableCellInfo3, value2);
						}
					}
					else
					{
						value2 = FormatValueLanTruoc(hTKKCMCTableCellInfo3, value2);
					}
					object obj2 = null;
					string text11 = "";
					if (hTKKCMCTableCellInfo3 != null)
					{
						obj2 = ((!(hTKKCMCTableCellInfo3.Controltype == "16")) ? sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1].Text.Trim() : sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1].Value);
					}
					text11 = FormatValueDieuChinh(hTKKCMCTableCellInfo3, obj2);
					text11 = text11.Trim();
					value2 = value2.Trim();
					if ((string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text11)) || (value2 != null && value2.ToLower().Trim() == text11.ToLower().Trim()) || (value2 != null && text11 != null && GeneralFunctions.IsNumeric(value2) && GeneralFunctions.IsNumeric(text11) && Convert.ToDouble(value2) == Convert.ToDouble(text11)))
					{
						continue;
					}
					if (nextRow > 0)
					{
						GeneralFunctions.InsertTableRow(sheetView2, dataLocationTemplateRowStart + nextRow - 1, -1, dataLocationTemplateRowStart + nextRow - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS], ToTableindex);
					}
					ThongTinThemChoKHBSMoi(sheetView2, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, text, colLoaiThue, text4);
					string text12 = GeneralFunctions.ListToString(list6, " - ");
					if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCTableCellInfo6 != null && obj != null && !string.IsNullOrEmpty(obj.ToString()))
					{
						text12 = text12.Replace(" - " + obj.ToString(), "");
					}
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text12;
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num4 - 1].Value = text10;
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num5 - 1].Value = text5;
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num13 - 1].Value = text2;
					GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView, dataLocationTemplateRowStart2 + num27, hTKKCMCTableCellInfo3.CellTemplate.ColIndex);
					FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView2, num6, num7, num8, col);
					if (hTKKCMCTableCellInfo3.Controltype == "16")
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = (GeneralFunctions.IsNumeric(value2) ? Convert.ToDouble(value2) : 0.0);
					}
					else
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = value2;
					}
					if (hTKKCMCTableCellInfo3.Controltype == "16")
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = (GeneralFunctions.IsNumeric(obj2.ToString()) ? obj2 : ((object)0));
					}
					else
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = obj2;
					}
					string maTieuMuc = "";
					XmlNode mapTieuMuc = null;
					int num29 = FindIndexTieuMuc(hTKKCMCTableInfo2, sheetIndex, tablePath, MCTPath, ref mapTieuMuc, ref maTieuMuc);
					if (string.IsNullOrEmpty(maTieuMuc))
					{
						maTieuMuc = TimTieuMucDynamicInTable(mapTieuMuc, sheetView, hTKKCMCTableInfo2, dataLocationTemplateRowStart2 + num27);
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = maTieuMuc;
					}
					else
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = maTieuMuc;
					}
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num21 - 1].Value = num;
					if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num9 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo4.CellTemplate.ColIndex - 1].Value;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num10 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo5.CellTemplate.ColIndex - 1].Value;
					}
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num14 - 1].Value = hTKKCMCTableCellInfo3.Controltype;
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num15 - 1].Value = text6;
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num16 - 1].Value = (string.IsNullOrEmpty(text7) ? "0" : text7);
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num8 - 1].Value = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num18 - 1].Value;
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = "";
					string text13 = ((!string.IsNullOrEmpty(text8)) ? text8 : ((hTKKCMCTableCellInfo6 != null) ? Convert.ToString(sheetView.Cells[dataLocationTemplateRowStart2 + num27 - 1, hTKKCMCTableCellInfo6.CellTemplate.ColIndex - 1].Value) : ""));
					if (!string.IsNullOrEmpty(text13))
					{
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = text13;
					}
					if (!string.IsNullOrEmpty(text13) && text6 != "2")
					{
						string text14 = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Text.Trim();
						if (!text14.StartsWith(text13))
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text13, text14, " - ");
						}
					}
					UpdateCtForDoiTac(text6, dataLocationTemplateRowStart + nextRow);
					sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num19 - 1].Value = Functions.FomartKey(tablePath, MCTPath);
					if (!string.IsNullOrEmpty(text4))
					{
						ChangeCellTypeHoachToan(dataLocationTemplateRowStart + nextRow, num7);
						SetLoaiTienStatic(sheetView, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
					}
					GeneralFunctions.SetRowHeight(sheetView2, dataLocationTemplateRowStart + nextRow);
					nextRow++;
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (HTKKCMCTableInfo item3 in list)
				{
					try
					{
						if (item3 == null || item3.RowsValues.Count <= 0)
						{
							continue;
						}
						List<HTKKCMCTableCellInfo> list3 = new List<HTKKCMCTableCellInfo>();
						List<int> list4 = new List<int>();
						List<string> list5 = GeneralFunctions.SplitString(text3, "###");
						foreach (string path2 in list5)
						{
							List<HTKKCMCTableCellInfo> cells2 = item3.RowsInfo.Cells;
							Func<HTKKCMCTableCellInfo, bool> predicate2 = (HTKKCMCTableCellInfo x) => x.Path == path2;
							HTKKCMCTableCellInfo hTKKCMCTableCellInfo = cells2.FirstOrDefault(predicate2);
							int item = item3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == path2);
							list3.Add(hTKKCMCTableCellInfo);
							list4.Add(item);
							if (hTKKCMCTableCellInfo == null && item3.Header.StaticCells.Count == 0)
							{
								HTKKFile.Log("xem lai path TCT: " + path2);
							}
						}
						if (list3.Count == 0 || (list3.Count > 0 && list3[0] == null))
						{
							continue;
						}
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo7 = item3.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
						int num30 = item3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo8 = item3.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
						int num31 = item3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo9 = item3.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
						int num32 = item3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo10 = item3.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
						int num33 = item3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo11 = item3.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
						int num34 = item3.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
						foreach (HTKKCMCRowValue rowsValue in item3.RowsValues)
						{
							List<string> list8 = new List<string>();
							foreach (int item4 in list4)
							{
								string text15 = ((item4 < 0) ? "" : rowsValue.RowValue[item4]);
								text15 = text15.Trim();
								if (item3 != null && num31 >= 0)
								{
									text15 = FormatValueLanTruoc(item3.RowsInfo.Cells[item4], text15);
								}
								list8.Add(text15);
							}
							if (GeneralFunctions.CheckInListEmpty(list8))
							{
								continue;
							}
							string text16 = GeneralFunctions.ListToString(list8, " - ");
							text16 = text16.Trim();
							string text17 = ((num30 >= 0 && rowsValue != null && rowsValue.RowValue.Count > num30) ? rowsValue.RowValue[num30] : "");
							text17 = text17.Trim();
							if (string.IsNullOrEmpty(text16))
							{
								continue;
							}
							int num35 = -1;
							if (!string.IsNullOrEmpty(prefixPath) && num34 >= 0)
							{
								string text18 = ((num34 < 0) ? "" : rowsValue.RowValue[num34]);
								text18 = text18.Trim();
								if (text18 != null)
								{
									if (!list5.Contains(prefixPath))
									{
										list5.Add(prefixPath);
									}
									list8.Add(text18);
								}
							}
							List<HTKKCMCTableInfo> list9 = ((hTKKCMCSheetInfo == null) ? null : DetectTableByKeyBodyBySview(sheetView, list2, list5, list8, item3.TablePath));
							HTKKCMCTableInfo hTKKCMCTableInfo2 = ((list9 != null && list9.Count > 0) ? list9[0] : null);
							if (!string.IsNullOrEmpty(text17) && hTKKCMCTableInfo2 != null && sheetView != null)
							{
								List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView, hTKKCMCTableInfo2, hTKKCMCTableCellInfo7.CellTemplate.ColIndex, text17);
								if (rowIndexInSViewByCompare.Count == 1)
								{
									num35 = rowIndexInSViewByCompare[0];
								}
							}
							if (string.IsNullOrEmpty(text17) && hTKKCMCTableInfo2 != null && sheetView != null)
							{
								List<int> list10 = new List<int>();
								foreach (HTKKCMCTableCellInfo item5 in list3)
								{
									if (item5 != null)
									{
										list10.Add(item5.CellTemplate.ColIndex);
									}
								}
								List<string> list11 = new List<string>();
								foreach (int item6 in list4)
								{
									if (rowsValue != null && rowsValue.RowValue.Count > item6)
									{
										list11.Add(rowsValue.RowValue[item6]);
									}
								}
								List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo2, list10, list8);
								if (rowIndexInSViewByCompares.Count == 1)
								{
									num35 = rowIndexInSViewByCompares[0];
								}
							}
							string text13 = "";
							if (!string.IsNullOrEmpty(text8))
							{
								text13 = text8;
								text16 = text16.Replace(" - " + text13, "");
							}
							else
							{
								text13 = ((rowsValue != null && rowsValue.RowValue != null && rowsValue.RowValue.Count > 0 && num34 >= 0) ? rowsValue.RowValue[num34] : "");
							}
							text13 = text13.Trim();
							if (!string.IsNullOrEmpty(text13) && !text16.StartsWith(text13))
							{
								text16 = GeneralFunctions.FormatKey2(text13, text16, " - ");
							}
							text16 = text16.Trim();
							if ((string.IsNullOrEmpty(text17) || checkExitsTrongsView(sheetView2, hTKKCMCTableInfo, num4, text17, num5, text5)) && (!string.IsNullOrEmpty(text17) || checkExitsTrongsView(sheetView2, hTKKCMCTableInfo, num3, text16, num5, text5)))
							{
								continue;
							}
							object obj2 = null;
							string text11 = "";
							if (sheetView != null && hTKKCMCTableCellInfo8 != null)
							{
								obj2 = ((!(hTKKCMCTableCellInfo8.Controltype == "16")) ? ((num35 < 0) ? "" : sheetView.Cells[num35 - 1, hTKKCMCTableCellInfo8.CellTemplate.ColIndex - 1].Text.Trim()) : ((num35 < 0) ? "0" : sheetView.Cells[num35 - 1, hTKKCMCTableCellInfo8.CellTemplate.ColIndex - 1].Value));
							}
							text11 = FormatValueDieuChinh(hTKKCMCTableCellInfo8, obj2);
							string value2 = ((rowsValue != null && rowsValue.RowValue != null && rowsValue.RowValue.Count > 0 && num31 >= 0) ? rowsValue.RowValue[num31] : "");
							if (item3 != null && num31 >= 0 && item3.RowsInfo.Cells.Count > num31)
							{
								value2 = FormatValueLanTruoc(item3.RowsInfo.Cells[num31], value2);
							}
							text11 = text11.Trim();
							value2 = value2.Trim();
							if ((string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(text11)) || value2.ToLower().Trim() == text11.ToLower().Trim() || (GeneralFunctions.IsNumeric(value2) && GeneralFunctions.IsNumeric(text11) && Convert.ToDouble(value2) == Convert.ToDouble(text11)))
							{
								continue;
							}
							if (nextRow > 0)
							{
								GeneralFunctions.InsertTableRow(sheetView2, dataLocationTemplateRowStart + nextRow - 1, -1, dataLocationTemplateRowStart + nextRow - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS], ToTableindex);
							}
							ThongTinThemChoKHBSMoi(sheetView2, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, text, colLoaiThue, text4);
							if (!string.IsNullOrEmpty(text8))
							{
								text13 = text8;
								text16 = text16.Replace(" - " + text13, "");
							}
							else if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCTableCellInfo11 != null && !string.IsNullOrEmpty(text13))
							{
								text16 = text16.Replace(" - " + text13, "");
							}
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text16;
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num4 - 1].Value = text17;
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num5 - 1].Value = text5;
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num13 - 1].Value = text2;
							if (sheetView != null && hTKKCMCTableInfo2 != null && num35 > 0)
							{
								GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView, (num35 > 0) ? num35 : item3.Location.DataLocationTemplateRowStart, hTKKCMCTableCellInfo8.CellTemplate.ColIndex);
								FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView2, num6, num7, num8, col);
							}
							else
							{
								GeneralFunctions.FormatCol4567KHBSBySinfo(hTKKCMCTableCellInfo8, sheetView2, dataLocationTemplateRowStart + nextRow, num6, num7, num8, col);
							}
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num14 - 1].Value = hTKKCMCTableCellInfo8.Controltype;
							if (hTKKCMCTableCellInfo8.Controltype == "16")
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = ((value2 != null && GeneralFunctions.IsNumeric(value2)) ? Convert.ToDouble(value2) : 0.0);
							}
							else
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = value2;
							}
							if (hTKKCMCTableCellInfo8.Controltype == "16")
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = ((obj2 != null && GeneralFunctions.IsNumeric(obj2.ToString())) ? obj2 : ((object)0));
							}
							else
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = obj2;
							}
							string maTieuMuc = "";
							XmlNode mapTieuMuc = null;
							int num29 = FindIndexTieuMuc(item3, sheetIndex, tablePath, MCTPath, ref mapTieuMuc, ref maTieuMuc);
							if (string.IsNullOrEmpty(maTieuMuc) && num29 >= 0)
							{
								if (num29 >= 0)
								{
									sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = rowsValue.RowValue[num29];
								}
							}
							else
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = maTieuMuc;
							}
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num21 - 1].Value = num;
							if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num9 - 1].Value = rowsValue.RowValue[num32];
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num10 - 1].Value = rowsValue.RowValue[num33];
							}
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num15 - 1].Value = text6;
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num16 - 1].Value = (string.IsNullOrEmpty(text7) ? "0" : text7);
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num8 - 1].Value = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num18 - 1].Value;
							if (!string.IsNullOrEmpty(text4))
							{
								ChangeCellTypeHoachToan(dataLocationTemplateRowStart + nextRow, num7);
								if (sheetView != null)
								{
									SetLoaiTienStatic(sheetView, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
								}
							}
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = "";
							if (!string.IsNullOrEmpty(text13))
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = text13;
							}
							if (!string.IsNullOrEmpty(text13) && text6 != "2")
							{
								string text14 = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Text.Trim();
								if (!text14.StartsWith(text13))
								{
									sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text13, text14, " - ");
								}
							}
							UpdateCtForDoiTac(text6, dataLocationTemplateRowStart + nextRow);
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num19 - 1].Value = Functions.FomartKey(tablePath, MCTPath);
							GeneralFunctions.SetRowHeight(sheetView2, dataLocationTemplateRowStart + nextRow);
							nextRow++;
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name + " lỗi ở sheet " + text + " MCT " + text2, ex);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name + " lỗi ở sheet " + text + " MCT " + text2, ex);
		}
		return nextRow;
	}

	private int CollectionDynamicHeader01_1_KHBS(XmlNode mapCt, int ToTableindex, int nextRow, HTKKCMCSheetInfo KHBS_LanTruoc)
	{
		string text = "";
		string text2 = "";
		try
		{
			if (mapCt == null)
			{
				return nextRow;
			}
			int sheetIndex = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			text = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string tablePath = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			text2 = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string text3 = ((mapCt.Attributes["TCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim());
			string maTCTPath = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string text4 = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string text5 = ((mapCt.Attributes["Caption"] == null) ? "" : ((XmlNode)mapCt.Attributes["Caption"]).Value.ToString().Trim());
			string DVTPath = ((mapCt.Attributes["DVTPath"] != null) ? ((XmlNode)mapCt.Attributes["DVTPath"]).Value.ToString().Trim() : "");
			string maDVTPath = ((mapCt.Attributes["MaDVTPath"] != null) ? ((XmlNode)mapCt.Attributes["MaDVTPath"]).Value.ToString().Trim() : "");
			string text6 = ((mapCt.Attributes["LoaiChiTieu"] != null) ? ((XmlNode)mapCt.Attributes["LoaiChiTieu"]).Value.ToString().Trim() : "");
			string text7 = ((mapCt.Attributes["Manual"] != null) ? ((XmlNode)mapCt.Attributes["Manual"]).Value.ToString().Trim() : "");
			string prefixPath = ((mapCt.Attributes["PrefixPath"] != null) ? ((XmlNode)mapCt.Attributes["PrefixPath"]).Value.ToString().Trim() : "");
			string text8 = ((mapCt.Attributes["Prefix"] != null) ? ((XmlNode)mapCt.Attributes["Prefix"]).Value.ToString().Trim() : "");
			int num = ((mapCt.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("BY");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("Z");
			int col = HTKKFile.ColumnLettersToIndex("AD");
			int num9 = HTKKFile.ColumnLettersToIndex("AH");
			int num10 = HTKKFile.ColumnLettersToIndex("CA");
			int num11 = HTKKFile.ColumnLettersToIndex("BR");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("BS");
			int colSheetName = HTKKFile.ColumnLettersToIndex("BT");
			int num12 = HTKKFile.ColumnLettersToIndex("G");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("BV");
			int num13 = HTKKFile.ColumnLettersToIndex("BX");
			int num14 = HTKKFile.ColumnLettersToIndex("BW");
			int num15 = HTKKFile.ColumnLettersToIndex("BD");
			int num16 = HTKKFile.ColumnLettersToIndex("BN");
			int num17 = HTKKFile.ColumnLettersToIndex("BO");
			int num18 = HTKKFile.ColumnLettersToIndex("BM");
			int num19 = HTKKFile.ColumnLettersToIndex("BQ");
			int num20 = HTKKFile.ColumnLettersToIndex("BL");
			int num21 = HTKKFile.ColumnLettersToIndex("CE");
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			List<HTKKCMCTableInfo> list = KHBS_LanTruoc?.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
			SheetView sheetView = ((hTKKCMCSheetInfo == null) ? null : winFpSpread.Sheets[sheetIndex]);
			List<HTKKCMCTableInfo> list2 = hTKKCMCSheetInfo?.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
			if (list2 != null)
			{
				for (int num22 = 0; num22 < list2.Count(); num22++)
				{
					try
					{
						HTKKCMCTableInfo hTKKCMCTableInfo = list2[num22];
						if (hTKKCMCTableInfo == null || hTKKCMCTableInfo.Header == null || hTKKCMCTableInfo.Header.StaticCells == null || hTKKCMCTableInfo.Header.StaticCells.Count == 0)
						{
							continue;
						}
						List<HTKKCMCCellInfo> list3 = new List<HTKKCMCCellInfo>();
						List<int> list4 = new List<int>();
						List<string> list5 = GeneralFunctions.SplitString(text3, "###");
						foreach (string path in list5)
						{
							List<HTKKCMCCellInfo> staticCells = hTKKCMCTableInfo.Header.StaticCells;
							Func<HTKKCMCCellInfo, bool> predicate = (HTKKCMCCellInfo x) => x.Path == path;
							HTKKCMCCellInfo hTKKCMCCellInfo = staticCells.FirstOrDefault(predicate);
							int item = hTKKCMCTableInfo.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == path);
							list3.Add(hTKKCMCCellInfo);
							list4.Add(item);
							if (hTKKCMCCellInfo == null && hTKKCMCTableInfo.Header.StaticCells.Count <= 0)
							{
							}
						}
						if (list3.Count == 0 || (list3.Count > 0 && list3[0] == null))
						{
							continue;
						}
						List<string> list6 = new List<string>();
						if (sheetView != null)
						{
							foreach (HTKKCMCCellInfo item4 in list3)
							{
								object value = sheetView.Cells[item4.CellTemplate.RowIndex - 1, item4.CellTemplate.ColIndex - 1].Value;
								string item2 = FormatValueDieuChinhStatic(item4, value);
								list6.Add(item2);
							}
						}
						object obj = null;
						HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == prefixPath);
						int num23 = hTKKCMCTableInfo.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == prefixPath);
						if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCCellInfo2 != null)
						{
							obj = sheetView?.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
							if (obj != null)
							{
								if (!list5.Contains(prefixPath))
								{
									list5.Add(prefixPath);
								}
								list6.Add(obj.ToString());
							}
						}
						HTKKCMCTableInfo tableByHeader = GeneralFunctions.GetTableByHeader(list, list5, list6);
						int num24 = tableByHeader?.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == MCTPath) ?? (-1);
						HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTCTPath);
						int num25 = hTKKCMCTableInfo.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maTCTPath);
						HTKKCMCCellInfo cellMCTToKhaiDieuChinh = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
						int num26 = hTKKCMCTableInfo.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == MCTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
						int num27 = hTKKCMCTableInfo.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
						int num28 = hTKKCMCTableInfo.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
						string value2 = ((sheetView != null && hTKKCMCCellInfo3 != null) ? Convert.ToString(sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value) : "");
						int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
						SheetView sheetView2 = winFpSpread.Sheets[sheetIndex_KHBS];
						HTKKCMCTableInfo hTKKCMCTableInfo2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS].Tables[ToTableindex];
						int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
						ICellType cellType = sheetView2.Cells[0, 0].CellType;
						TextCellType textCellType = new TextCellType();
						string text9 = "";
						if (tableByHeader != null)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo6 = tableByHeader.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == cellMCTToKhaiDieuChinh.Path);
							if (hTKKCMCCellInfo6 != null)
							{
								text9 = FormatValueLanTruocStatic(cellMCTToKhaiDieuChinh, hTKKCMCCellInfo6.Value);
							}
						}
						else
						{
							text9 = FormatValueLanTruocStatic(cellMCTToKhaiDieuChinh, text9);
						}
						object obj2 = null;
						string text10 = "";
						if (sheetView != null)
						{
							obj2 = ((!(cellMCTToKhaiDieuChinh.Controltype == "16")) ? sheetView.Cells[cellMCTToKhaiDieuChinh.CellTemplate.RowIndex - 1, cellMCTToKhaiDieuChinh.CellTemplate.ColIndex - 1].Text.Trim() : sheetView.Cells[cellMCTToKhaiDieuChinh.CellTemplate.RowIndex - 1, cellMCTToKhaiDieuChinh.CellTemplate.ColIndex - 1].Value);
						}
						text10 = FormatValueDieuChinhStatic(cellMCTToKhaiDieuChinh, obj2);
						if ((string.IsNullOrEmpty(text9) && string.IsNullOrEmpty(text10)) || (text9 != null && text9.ToLower().Trim() == text10.ToLower().Trim()) || (text9 != null && text10 != null && GeneralFunctions.IsNumeric(text9) && GeneralFunctions.IsNumeric(text10) && Convert.ToDouble(text9) == Convert.ToDouble(text10)))
						{
							continue;
						}
						if (nextRow > 0)
						{
							GeneralFunctions.InsertTableRow(sheetView2, dataLocationTemplateRowStart + nextRow - 1, -1, dataLocationTemplateRowStart + nextRow - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS], ToTableindex);
						}
						ThongTinThemChoKHBSMoi(sheetView2, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, text, colLoaiThue, text4);
						string text11 = GeneralFunctions.ListToString(list6, " - ");
						if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCCellInfo2 != null && !string.IsNullOrEmpty(obj.ToString()))
						{
							text11 = text11.Replace(" - " + obj.ToString(), "");
						}
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text11;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num4 - 1].Value = value2;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num5 - 1].Value = text5;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num13 - 1].Value = text2;
						GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView, cellMCTToKhaiDieuChinh.CellTemplate.RowIndex, cellMCTToKhaiDieuChinh.CellTemplate.ColIndex);
						FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView2, num6, num7, num8, col);
						if (cellMCTToKhaiDieuChinh.Controltype == "16")
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = (GeneralFunctions.IsNumeric(text9) ? Convert.ToDouble(text9) : 0.0);
						}
						else
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = text9;
						}
						if (cellMCTToKhaiDieuChinh.Controltype == "16")
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = (GeneralFunctions.IsNumeric(obj2.ToString()) ? obj2 : ((object)0));
						}
						else
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = obj2;
						}
						XmlNode mapTieuMuc = null;
						string value3 = FindTieuMucInTableHeader(tableByHeader, sheetIndex, tablePath, MCTPath, ref mapTieuMuc);
						if (string.IsNullOrEmpty(value3))
						{
							value3 = FindTieuMucInSview(sheetView, hTKKCMCTableInfo, ref mapTieuMuc);
						}
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = value3;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num21 - 1].Value = num;
						if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num9 - 1].Value = sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num10 - 1].Value = sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value;
						}
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num14 - 1].Value = cellMCTToKhaiDieuChinh.Controltype;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num15 - 1].Value = text6;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num16 - 1].Value = (string.IsNullOrEmpty(text7) ? "0" : text7);
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num8 - 1].Value = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num18 - 1].Value;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = "";
						string text12 = ((!string.IsNullOrEmpty(text8)) ? text8 : ((hTKKCMCCellInfo2 != null) ? Convert.ToString(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value) : ""));
						if (!string.IsNullOrEmpty(text12))
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = text12;
						}
						if (!string.IsNullOrEmpty(text12) && text6 != "2")
						{
							string text13 = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Text.Trim();
							if (!text13.StartsWith(text12))
							{
								sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text12, text13, " - ");
							}
						}
						UpdateCtForDoiTac(text6, dataLocationTemplateRowStart + nextRow);
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num19 - 1].Value = Functions.FomartKey(tablePath, MCTPath);
						if (!string.IsNullOrEmpty(text4))
						{
							ChangeCellTypeHoachToan(dataLocationTemplateRowStart + nextRow, num7);
							if (sheetView != null)
							{
								SetLoaiTienStatic(sheetView, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
							}
						}
						GeneralFunctions.SetRowHeight(sheetView2, dataLocationTemplateRowStart + nextRow);
						nextRow++;
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name + "header lỗi ở sheet " + text + " MCT " + text2, ex);
					}
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (HTKKCMCTableInfo item5 in list)
				{
					try
					{
						if (item5 == null || item5.Header == null || item5.Header.StaticCells == null || item5.Header.StaticCells.Count <= 0)
						{
							continue;
						}
						List<HTKKCMCCellInfo> list7 = new List<HTKKCMCCellInfo>();
						List<int> list8 = new List<int>();
						List<string> list5 = GeneralFunctions.SplitString(text3, "###");
						foreach (string path2 in list5)
						{
							List<HTKKCMCCellInfo> staticCells2 = item5.Header.StaticCells;
							Func<HTKKCMCCellInfo, bool> predicate2 = (HTKKCMCCellInfo x) => x.Path == path2;
							HTKKCMCCellInfo hTKKCMCCellInfo7 = staticCells2.FirstOrDefault(predicate2);
							int item3 = item5.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == path2);
							list7.Add(hTKKCMCCellInfo7);
							list8.Add(item3);
							if (hTKKCMCCellInfo7 == null && item5.Header.StaticCells.Count <= 0)
							{
							}
						}
						if (list7.Count == 0 || (list7.Count > 0 && list7[0] == null))
						{
							continue;
						}
						List<string> list9 = new List<string>();
						foreach (HTKKCMCCellInfo item6 in list7)
						{
							string value4 = item6.Value;
							list9.Add(value4);
						}
						if (GeneralFunctions.CheckInListEmpty(list9))
						{
							continue;
						}
						string valueCompare = GeneralFunctions.ListToString(list9, " - ");
						int num24 = item5?.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == MCTPath) ?? (-1);
						HTKKCMCCellInfo hTKKCMCCellInfo8 = item5.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTCTPath);
						int num29 = item5.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maTCTPath);
						HTKKCMCCellInfo cellMCT_KHBS_LanTruoc = item5.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
						int num30 = item5.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == MCTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo9 = item5.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
						int num31 = item5.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo10 = item5.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
						int num32 = item5.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo11 = item5.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == prefixPath);
						int num33 = item5.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == prefixPath);
						string text14 = ((hTKKCMCCellInfo8 != null) ? hTKKCMCCellInfo8.Value : "");
						if (cellMCT_KHBS_LanTruoc == null && item5.Header.StaticCells.Count > 0)
						{
							continue;
						}
						int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
						SheetView sheetView2 = winFpSpread.Sheets[sheetIndex_KHBS];
						HTKKCMCTableInfo hTKKCMCTableInfo2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS].Tables[ToTableindex];
						int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
						ICellType cellType = sheetView2.Cells[0, 0].CellType;
						TextCellType textCellType = new TextCellType();
						if ((string.IsNullOrEmpty(text14) || checkExitsTrongsView(sheetView2, hTKKCMCTableInfo2, num4, text14, num5, text5)) && (!string.IsNullOrEmpty(text14) || checkExitsTrongsView(sheetView2, hTKKCMCTableInfo2, num3, valueCompare, num5, text5)))
						{
							continue;
						}
						string text9 = "";
						if (cellMCT_KHBS_LanTruoc != null)
						{
							text9 = FormatValueLanTruocStatic(cellMCT_KHBS_LanTruoc, cellMCT_KHBS_LanTruoc.Value);
						}
						object obj2 = null;
						string text10 = "";
						string text15 = "";
						if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCCellInfo11 != null)
						{
							text15 = hTKKCMCCellInfo11.Value;
							if (text15 != null)
							{
								if (!list5.Contains(prefixPath))
								{
									list5.Add(prefixPath);
								}
								list9.Add(text15);
							}
						}
						HTKKCMCTableInfo hTKKCMCTableInfo3 = ((sheetView == null) ? null : DetectIndexTableHeaderBySview(sheetView, list2, list5, list9, item5.TablePath));
						if (hTKKCMCTableInfo3 != null)
						{
							HTKKCMCCellInfo hTKKCMCCellInfo12 = hTKKCMCTableInfo3?.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == cellMCT_KHBS_LanTruoc.Path);
							if (sheetView != null)
							{
								obj2 = ((!(cellMCT_KHBS_LanTruoc.Controltype == "16")) ? sheetView.Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Text.Trim() : sheetView.Cells[hTKKCMCCellInfo12.CellTemplate.RowIndex - 1, hTKKCMCCellInfo12.CellTemplate.ColIndex - 1].Value);
							}
						}
						text10 = FormatValueDieuChinhStatic(cellMCT_KHBS_LanTruoc, obj2);
						if ((string.IsNullOrEmpty(text9) && string.IsNullOrEmpty(text10)) || (text9 != null && text9.ToLower().Trim() == text10.ToLower().Trim()) || (text9 != null && text10 != null && GeneralFunctions.IsNumeric(text9) && GeneralFunctions.IsNumeric(text10) && Convert.ToDouble(text9) == Convert.ToDouble(text10)))
						{
							continue;
						}
						if (nextRow > 0)
						{
							GeneralFunctions.InsertTableRow(sheetView2, dataLocationTemplateRowStart + nextRow - 1, -1, dataLocationTemplateRowStart + nextRow - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS], ToTableindex);
						}
						ThongTinThemChoKHBSMoi(sheetView2, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, text, colLoaiThue, text4);
						string text11 = GeneralFunctions.ListToString(list9, " - ");
						if (!string.IsNullOrEmpty(prefixPath) && hTKKCMCCellInfo11 != null && !string.IsNullOrEmpty(text15))
						{
							text11 = text11.Replace(" - " + text15, "");
						}
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text11;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num4 - 1].Value = text14;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num5 - 1].Value = text5;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num13 - 1].Value = text2;
						if (sheetView != null && hTKKCMCTableInfo3 != null)
						{
							if (GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView, cellMCT_KHBS_LanTruoc.CellTemplate.RowIndex, cellMCT_KHBS_LanTruoc.CellTemplate.ColIndex))
							{
								FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView2, num6, num7, num8, col);
							}
							else
							{
								GeneralFunctions.FormatCol4567KHBSBySinfoStatic(cellMCT_KHBS_LanTruoc, sheetView2, dataLocationTemplateRowStart + nextRow, num6, num7, num8, col);
							}
						}
						else
						{
							GeneralFunctions.FormatCol4567KHBSBySinfoStatic(cellMCT_KHBS_LanTruoc, sheetView2, dataLocationTemplateRowStart + nextRow, num6, num7, num8, col);
						}
						if (cellMCT_KHBS_LanTruoc.Controltype == "16")
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = (GeneralFunctions.IsNumeric(text9) ? Convert.ToDouble(text9) : 0.0);
						}
						else
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = text9;
						}
						if (cellMCT_KHBS_LanTruoc.Controltype == "16")
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = 0;
						}
						else
						{
							sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = "";
						}
						XmlNode mapTieuMuc = null;
						string value3 = FindTieuMucInTableHeader(item5, sheetIndex, tablePath, MCTPath, ref mapTieuMuc);
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = value3;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num21 - 1].Value = num;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num14 - 1].Value = cellMCT_KHBS_LanTruoc.Controltype;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num15 - 1].Value = text6;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num16 - 1].Value = (string.IsNullOrEmpty(text7) ? "0" : text7);
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num8 - 1].Value = sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num18 - 1].Value;
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = "";
						UpdateCtForDoiTac(text6, dataLocationTemplateRowStart + nextRow);
						sheetView2.Cells[dataLocationTemplateRowStart + nextRow - 1, num19 - 1].Value = Functions.FomartKey(tablePath, MCTPath);
						if (!string.IsNullOrEmpty(text4))
						{
							ChangeCellTypeHoachToan(dataLocationTemplateRowStart + nextRow, num7);
							if (sheetView != null)
							{
								SetLoaiTienStatic(sheetView, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
							}
						}
						GeneralFunctions.SetRowHeight(sheetView2, dataLocationTemplateRowStart + nextRow);
						nextRow++;
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name + "header lỗi ở sheet " + text + " MCT " + text2, ex);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name + "header lỗi ở sheet " + text + " MCT " + text2, ex);
		}
		return nextRow;
	}

	public List<HTKKCMCTableInfo> DetectTableByKeyBodyBySview(SheetView sView, List<HTKKCMCTableInfo> tables, List<string> paths, List<string> values, string tablePath)
	{
		try
		{
			List<HTKKCMCTableInfo> list = new List<HTKKCMCTableInfo>();
			int num = 0;
			List<HTKKCMCTableInfo> list2 = tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
			if (list2 != null && list2.Count() == 1 && (list2[0].Header == null || list2[0].Header.Length == 0))
			{
				return list2.ToList();
			}
			foreach (HTKKCMCTableInfo table in tables)
			{
				if (!string.IsNullOrEmpty(tablePath) && table.TablePath != tablePath)
				{
					continue;
				}
				int num2 = 0;
				int i;
				for (i = 0; i < paths.Count; i++)
				{
					if (!string.IsNullOrEmpty(values[i]))
					{
						int num3 = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == paths[i]);
						if (num3 < 0)
						{
							continue;
						}
						string controltype = table.RowsInfo.Cells[num3].Controltype;
						for (int num4 = 0; num4 < table.RowsValues.Count; num4++)
						{
							if (controltype == "14")
							{
								if (table.RowsValues[num4].RowValue[num3] == values[i].Replace("/", ""))
								{
									num2++;
									break;
								}
							}
							else if (table.RowsValues[num4].RowValue[num3] == values[i])
							{
								num2++;
								break;
							}
						}
					}
					else
					{
						num2++;
					}
				}
				if (num2 > 0 && num2 == paths.Count && num >= 0 && num < tables.Count)
				{
					list.Add(tables[num]);
				}
				num++;
			}
			if (list.Count > 0)
			{
				return list;
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public HTKKCMCTableInfo DetectIndexTableHeaderBySview(SheetView sView, List<HTKKCMCTableInfo> tables, List<string> paths, List<string> values, string tablePath)
	{
		try
		{
			int num = 0;
			List<HTKKCMCTableInfo> list = tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList();
			if (list != null && list.Count() == 1 && list[0].Header != null && list[0].Header.Length == 0)
			{
				return list.First();
			}
			foreach (HTKKCMCTableInfo table in tables)
			{
				if (!string.IsNullOrEmpty(tablePath) && table.TablePath != tablePath)
				{
					continue;
				}
				int num2 = 0;
				int i;
				for (i = 0; i < paths.Count; i++)
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == paths[i]);
					if (hTKKCMCCellInfo == null)
					{
						continue;
					}
					string text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
					if (hTKKCMCCellInfo.Controltype == "16")
					{
						if (!string.IsNullOrEmpty(values[i]) && text == values[i])
						{
							num2++;
						}
						else if (string.IsNullOrEmpty(values[i]) && !string.IsNullOrEmpty(text) && Convert.ToDouble(text) == 0.0)
						{
							num2++;
						}
					}
					else if (text == values[i])
					{
						num2++;
					}
				}
				if (num2 > 0 && num2 == paths.Count)
				{
					break;
				}
				num++;
			}
			if (num >= 0 && num < tables.Count)
			{
				return tables[num];
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return null;
	}

	public string FormatValueDieuChinh(HTKKCMCTableCellInfo cellInfo, object value)
	{
		try
		{
			string text = "";
			return (cellInfo != null && cellInfo.Controltype == "16") ? ((value != null && value.ToString().Trim() != "" && cellInfo.DefaultValue != null && cellInfo.DefaultValue.IndexOf(",") >= 0) ? Convert.ToDouble(value).ToString().Trim() : ((value != null && !string.IsNullOrEmpty(value.ToString().Trim()) && Convert.ToDouble(value) != 0.0) ? ((cellInfo.DefaultValue != null && cellInfo.DefaultValue.IndexOf(",") < 0) ? value.ToString().Trim().Replace(".0", "") : Convert.ToDouble(value).ToString().Trim()) : "0")) : ((cellInfo != null && cellInfo.Controltype == "14" && value != null && value.ToString().Trim() != "" && value.ToString().IndexOf("/") < 0) ? GeneralFunctions.FormartDateStr(value.ToString().Trim()) : ((cellInfo != null && cellInfo.Controltype == "2" && value != null && value.ToString().Trim() != "") ? (Convert.ToBoolean(value) ? "True" : "False") : ((value == null) ? "" : value.ToString().Trim())));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public string FormatValueDieuChinhStatic(HTKKCMCCellInfo cellInfo, object value)
	{
		try
		{
			string text = "";
			return (cellInfo != null && cellInfo.Controltype == "16") ? ((value != null && value.ToString().Trim() != "" && cellInfo.DefaultValue != null && cellInfo.DefaultValue.IndexOf(",") >= 0) ? Convert.ToDouble(value).ToString().Trim() : ((value != null && !string.IsNullOrEmpty(value.ToString().Trim()) && Convert.ToDouble(value) != 0.0) ? ((cellInfo.DefaultValue != null && cellInfo.DefaultValue.IndexOf(",") < 0) ? value.ToString().Trim().Replace(".0", "") : Convert.ToDouble(value).ToString().Trim()) : "0")) : ((cellInfo != null && cellInfo.Controltype == "14" && value != null && value.ToString().Trim() != "" && value.ToString().IndexOf("/") < 0) ? GeneralFunctions.FormartDateStr(value.ToString().Trim()) : ((cellInfo != null && cellInfo.Controltype == "2" && value != null && value.ToString().Trim() != "") ? (Convert.ToBoolean(value) ? "True" : "False") : ((value == null) ? "" : value.ToString().Trim())));
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public string FormatValueLanTruoc(HTKKCMCTableCellInfo cellInfo, string value)
	{
		try
		{
			value = value.Trim();
			if (cellInfo.Controltype == "2")
			{
				string text = value;
			}
			if (cellInfo != null && cellInfo.Controltype == "14" && !string.IsNullOrEmpty(value))
			{
				value = GeneralFunctions.FormartDateStr(value);
			}
			if ((cellInfo != null && cellInfo.Controltype == "16" && string.IsNullOrEmpty(value)) || (cellInfo.Controltype == "16" && !string.IsNullOrEmpty(value) && Convert.ToDouble(value) == 0.0))
			{
				value = "0";
			}
			return value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public string FormatValueLanTruocStatic(HTKKCMCCellInfo cellInfo, string value)
	{
		try
		{
			value = value.Trim();
			if (cellInfo != null && cellInfo.Controltype == "14" && !string.IsNullOrEmpty(value))
			{
				value = GeneralFunctions.FormartDateStr(value);
			}
			if ((cellInfo != null && cellInfo.Controltype == "16" && string.IsNullOrEmpty(value)) || (cellInfo.Controltype == "16" && !string.IsNullOrEmpty(value) && Convert.ToDouble(value) == 0.0))
			{
				value = "0";
			}
			return value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public int FindIndexTieuMuc(HTKKCMCTableInfo tableToKhaiDieuChinh, int sheetIndex, string tablePath, string MCTPath, ref XmlNode mapTieuMuc, ref string maTieuMuc)
	{
		try
		{
			int num = -1;
			if (tableToKhaiDieuChinh == null)
			{
				return num;
			}
			maTieuMuc = "";
			mapTieuMuc = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath)));
			if (mapTieuMuc != null)
			{
				maTieuMuc = ((mapTieuMuc.Attributes["MaTieuMuc"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMuc"]).Value.Trim() : "");
			}
			if (string.IsNullOrEmpty(maTieuMuc))
			{
				string maTieuMucPath = ((mapTieuMuc != null && mapTieuMuc.Attributes["MaTieuMucPath"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMucPath"]).Value : "");
				if (!string.IsNullOrEmpty(maTieuMucPath))
				{
					num = tableToKhaiDieuChinh.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTieuMucPath);
				}
			}
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1;
		}
	}

	public string FindTieuMucInTableHeader(HTKKCMCTableInfo table, int sheetIndex, string tablePath, string MCTPath, ref XmlNode mapTieuMuc)
	{
		try
		{
			string value = "";
			mapTieuMuc = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath)));
			if (mapTieuMuc != null)
			{
				value = ((mapTieuMuc.Attributes["MaTieuMuc"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMuc"]).Value.Trim() : "");
			}
			if (string.IsNullOrEmpty(value) && table != null)
			{
				string maTieuMucPath = ((mapTieuMuc != null && mapTieuMuc.Attributes["MaTieuMucPath"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMucPath"]).Value : "");
				if (!string.IsNullOrEmpty(maTieuMucPath))
				{
					HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTieuMucPath);
					if (hTKKCMCCellInfo != null)
					{
						value = hTKKCMCCellInfo.Value;
					}
				}
			}
			return value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public string FindTieuMucInSview(SheetView sView, HTKKCMCTableInfo table, ref XmlNode mapTieuMuc)
	{
		try
		{
			string text = "";
			if (mapTieuMuc == null)
			{
				return text;
			}
			string maTieuMucPath = ((mapTieuMuc != null && mapTieuMuc.Attributes["MaTieuMucPath"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMucPath"]).Value : "");
			if (!string.IsNullOrEmpty(maTieuMucPath))
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = table.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTieuMucPath);
				if (hTKKCMCCellInfo != null)
				{
					text = sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				}
			}
			return text;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	public ICellType GetCellType(int sheetIndex_KHBS, int row_KHBS, int col_KHBS)
	{
		try
		{
			List<string> list = new List<string>();
			list.Add(sheetIndex_KHBS.ToString());
			list.Add(row_KHBS.ToString());
			list.Add(col_KHBS.ToString());
			string key = Functions.FormatKey(list, "___");
			if (ConvertHelpers._dicCellType_01_1_KHBS.ContainsKey(key))
			{
				return ConvertHelpers._dicCellType_01_1_KHBS[key].iCellType;
			}
			return null;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void FormatCol4567(int sheetIndex_KHBS, int row_KHBS, int col_KHBS, SheetView sView_KHBS, int col4, int col5, int col6, int col7)
	{
		try
		{
			NumberCellType numberCellType = new NumberCellType();
			numberCellType.DecimalPlaces = 0;
			CurrencyCellType currencyCellType = new CurrencyCellType();
			TextCellType textCellType = new TextCellType();
			CheckBoxCellType checkBoxCellType = new CheckBoxCellType();
			MaskCellType maskCellType = new MaskCellType();
			List<string> list = new List<string>();
			list.Add(sheetIndex_KHBS.ToString());
			list.Add(row_KHBS.ToString());
			list.Add(col_KHBS.ToString());
			string key = Functions.FormatKey(list, "___");
			if (!ConvertHelpers._dicCellType_01_1_KHBS.ContainsKey(key))
			{
				return;
			}
			ICellType iCellType = ConvertHelpers._dicCellType_01_1_KHBS[key].iCellType;
			object obj = sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value;
			object obj2 = sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value;
			if ((object)iCellType.GetType() == numberCellType.GetType() || (object)iCellType.GetType() == currencyCellType.GetType())
			{
				if (obj != null)
				{
					obj = obj.ToString().Replace(",", ".");
				}
				if (obj2 != null)
				{
					obj2 = obj2.ToString().Replace(",", ".");
				}
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = ((obj != null && !string.IsNullOrEmpty(obj.ToString().Trim()) && GeneralFunctions.IsNumeric(obj.ToString())) ? Convert.ToDouble(obj) : 0.0);
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = ((obj2 != null && !string.IsNullOrEmpty(obj2.ToString().Trim()) && GeneralFunctions.IsNumeric(obj2.ToString())) ? Convert.ToDouble(obj2) : 0.0);
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].CellType = iCellType;
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].CellType = iCellType;
				sView_KHBS.Cells[row_KHBS - 1, col6 - 1].CellType = iCellType;
				sView_KHBS.Cells[row_KHBS - 1, col7 - 1].CellType = iCellType;
				return;
			}
			sView_KHBS.Cells[row_KHBS - 1, col4 - 1].CellType = iCellType;
			sView_KHBS.Cells[row_KHBS - 1, col5 - 1].CellType = iCellType;
			if ((object)iCellType.GetType() == checkBoxCellType.GetType() && (sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Text.Trim() == "1" || sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Text.Trim().ToLower() == "true"))
			{
				sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = true;
			}
			if ((object)iCellType.GetType() == checkBoxCellType.GetType() && (sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Text.Trim() == "1" || sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Text.Trim().ToLower() == "true"))
			{
				sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = true;
			}
			if ((object)iCellType.GetType() == maskCellType.GetType())
			{
				if (sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value != null)
				{
					string text = sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value.ToString().Trim();
					if (text.Length == 10 && text.IndexOf("-") >= 0)
					{
						string value = DateTime.ParseExact(text, "yyyy-MM-dd", null).ToString("ddMMyyyy");
						sView_KHBS.Cells[row_KHBS - 1, col4 - 1].Value = value;
					}
				}
				if (sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value != null)
				{
					string text2 = sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value.ToString().Trim();
					if (text2.Length == 10 && text2.IndexOf("-") >= 0)
					{
						string value2 = DateTime.ParseExact(text2, "yyyy-MM-dd", null).ToString("ddMMyyyy");
						sView_KHBS.Cells[row_KHBS - 1, col5 - 1].Value = value2;
					}
				}
			}
			sView_KHBS.Cells[row_KHBS - 1, col6 - 1].CellType = numberCellType;
			sView_KHBS.Cells[row_KHBS - 1, col7 - 1].CellType = numberCellType;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public string TimTieuMucDynamicInTable(XmlNode mapTieuMuc, SheetView sView, HTKKCMCTableInfo table, int row)
	{
		try
		{
			if (mapTieuMuc == null)
			{
				return "";
			}
			string maTieuMucPath = ((mapTieuMuc.Attributes["MaTieuMucPath"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMucPath"]).Value : "");
			if (string.IsNullOrEmpty(maTieuMucPath))
			{
				return "";
			}
			HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTieuMucPath);
			if (hTKKCMCTableCellInfo != null)
			{
				return sView.Cells[row - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	public string TimTieuMucStatic(XmlNode mapTieuMuc, HTKKCMCSheetInfo sInfo, SheetView sView)
	{
		try
		{
			if (mapTieuMuc == null)
			{
				return "";
			}
			string maTieuMucPath = ((mapTieuMuc.Attributes["MaTieuMucPath"] == null) ? "" : ((XmlNode)mapTieuMuc.Attributes["MaTieuMucPath"]).Value.Trim());
			if (string.IsNullOrEmpty(maTieuMucPath))
			{
				return "";
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(maTieuMucPath));
			if (hTKKCMCCellInfo != null)
			{
				return sView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	public string TimTieuMucStaticLanTruoc(XmlNode mapTieuMuc, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (mapTieuMuc == null)
			{
				return "";
			}
			string maTieuMucPath = ((mapTieuMuc.Attributes["MaTieuMucPath"] != null) ? ((XmlNode)mapTieuMuc.Attributes["MaTieuMucPath"]).Value : "");
			if (string.IsNullOrEmpty(maTieuMucPath))
			{
				return "";
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTieuMucPath);
			if (hTKKCMCCellInfo != null)
			{
				return hTKKCMCCellInfo.Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	private int CollectionStatic01_1_KHBS(XmlNode mapCt, int toTableindex, int nextRow, HTKKCMCSheetInfo KHBS_LanTruoc)
	{
		try
		{
			if (mapCt == null)
			{
				return nextRow;
			}
			int sheetIndex = ((mapCt.Attributes["SheetIndex"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value) : (-1));
			string sheetname = ((mapCt.Attributes["SheetName"] != null) ? ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim() : "");
			string text = ((mapCt.Attributes["MCT"] != null) ? ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim() : "");
			string MCTPath = ((mapCt.Attributes["MCTPath"] != null) ? ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim() : "");
			string text2 = ((mapCt.Attributes["Caption"] != null) ? ((XmlNode)mapCt.Attributes["Caption"]).Value.ToString().Trim() : "");
			string text3 = ((mapCt.Attributes["LoaiThue"] != null) ? ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim() : "");
			string DVTPath = ((mapCt.Attributes["DVTPath"] != null) ? ((XmlNode)mapCt.Attributes["DVTPath"]).Value.ToString().Trim() : "");
			string maDVTPath = ((mapCt.Attributes["MaDVTPath"] != null) ? ((XmlNode)mapCt.Attributes["MaDVTPath"]).Value.ToString().Trim() : "");
			string text4 = ((mapCt.Attributes["LoaiChiTieu"] != null) ? ((XmlNode)mapCt.Attributes["LoaiChiTieu"]).Value.ToString().Trim() : "");
			string text5 = ((mapCt.Attributes["Manual"] != null) ? ((XmlNode)mapCt.Attributes["Manual"]).Value.ToString().Trim() : "");
			int num = ((mapCt.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
			string value = ((mapCt.Attributes["TCTPath"] != null) ? ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim() : "");
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("BY");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int num6 = HTKKFile.ColumnLettersToIndex("R");
			int num7 = HTKKFile.ColumnLettersToIndex("V");
			int num8 = HTKKFile.ColumnLettersToIndex("Z");
			int col = HTKKFile.ColumnLettersToIndex("AD");
			int num9 = HTKKFile.ColumnLettersToIndex("BR");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("BS");
			int colSheetName = HTKKFile.ColumnLettersToIndex("BT");
			int num10 = HTKKFile.ColumnLettersToIndex("G");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("BV");
			int num11 = HTKKFile.ColumnLettersToIndex("BX");
			int num12 = HTKKFile.ColumnLettersToIndex("BW");
			int num13 = HTKKFile.ColumnLettersToIndex("AH");
			int num14 = HTKKFile.ColumnLettersToIndex("CA");
			int num15 = HTKKFile.ColumnLettersToIndex("BD");
			int num16 = HTKKFile.ColumnLettersToIndex("BN");
			int num17 = HTKKFile.ColumnLettersToIndex("BO");
			int num18 = HTKKFile.ColumnLettersToIndex("BM");
			int num19 = HTKKFile.ColumnLettersToIndex("BQ");
			int num20 = HTKKFile.ColumnLettersToIndex("CE");
			int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex_KHBS];
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS].Tables[toTableindex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			ICellType cellType = sheetView.Cells[0, 0].CellType;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			SheetView sheetView2 = ((hTKKCMCSheetInfo == null) ? null : winFpSpread.Sheets[sheetIndex]);
			if (hTKKCMCSheetInfo != null && sheetView2 != null)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
				if (hTKKCMCCellInfo == null)
				{
					return nextRow;
				}
				HTKKCMCCellInfo hTKKCMCCellInfo2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == DVTPath);
				int num21 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(DVTPath) && x.Path == DVTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo3 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maDVTPath);
				int num22 = hTKKCMCTableInfo.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => !string.IsNullOrEmpty(maDVTPath) && x.Path == maDVTPath);
				string value2 = ((hTKKCMCCellInfo2 == null) ? "" : sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim());
				string value3 = ((hTKKCMCCellInfo3 == null) ? "" : sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Text.Trim());
				object obj = null;
				string text6 = "";
				string text7 = "";
				HTKKCMCCellInfo hTKKCMCCellInfo4 = null;
				XmlNode val = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] == null)));
				if (val != null)
				{
					text7 = ((XmlNode)val.Attributes["MaTieuMuc"]).Value.Trim();
					if (string.IsNullOrEmpty(text7))
					{
						text7 = TimTieuMucStatic(val, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], sheetView2);
					}
				}
				string text8 = TimTieuMucStaticLanTruoc(val, KHBS_LanTruoc);
				hTKKCMCCellInfo4 = ((string.IsNullOrEmpty(value) || val == null || string.IsNullOrEmpty(text8)) ? KHBS_LanTruoc?.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath) : ((KHBS_LanTruoc == null) ? null : ((!string.IsNullOrEmpty(text8) && text8 != text7) ? null : KHBS_LanTruoc.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath))));
				if (sheetView2 != null && hTKKCMCCellInfo != null)
				{
					obj = ((!(hTKKCMCCellInfo.Controltype == "16")) ? sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim() : sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
				}
				text6 = FormatValueDieuChinhStatic(hTKKCMCCellInfo, obj);
				text6 = text6.Trim();
				string text9 = FormatValueLanTruocStatic(hTKKCMCCellInfo, (hTKKCMCCellInfo4 != null) ? hTKKCMCCellInfo4.Value : "");
				text9 = text9.Trim();
				if ((hTKKCMCCellInfo4 != null && text9 == text6) || (hTKKCMCCellInfo4 != null && GeneralFunctions.IsNumeric(text9) && GeneralFunctions.IsNumeric(text6) && Convert.ToDouble(text9) == Convert.ToDouble(text6)))
				{
					return nextRow;
				}
				if (nextRow > 0)
				{
					GeneralFunctions.InsertTableRow(sheetView, dataLocationTemplateRowStart + nextRow - 1, -1, dataLocationTemplateRowStart + nextRow - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS], toTableindex);
				}
				ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text3);
				TextCellType textCellType = new TextCellType();
				ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text3);
				if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(text7))
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text2;
				}
				else
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text2 + " - Tiểu mục " + text7;
				}
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num4 - 1].Value = text;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num5 - 1].Value = text;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = text;
				if (hTKKCMCCellInfo != null && hTKKCMCCellInfo.Controltype == "16")
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = ((hTKKCMCCellInfo4 != null && GeneralFunctions.IsNumeric(text9)) ? Convert.ToDouble(text9) : 0.0);
				}
				else
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = ((hTKKCMCCellInfo4 != null) ? text9 : "");
				}
				if (hTKKCMCCellInfo != null && hTKKCMCCellInfo.Controltype == "16")
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = (GeneralFunctions.IsNumeric(obj.ToString()) ? obj : ((object)0));
				}
				else
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = obj;
				}
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num9 - 1].Value = text7;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = num;
				if (!string.IsNullOrEmpty(text3) && num21 >= 0)
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num13 - 1].Value = value2;
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num14 - 1].Value = value3;
				}
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num12 - 1].Value = ((hTKKCMCCellInfo != null) ? hTKKCMCCellInfo.Controltype : "");
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num15 - 1].Value = text4;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num16 - 1].Value = (string.IsNullOrEmpty(text5) ? "0" : text5);
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num8 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num18 - 1].Value;
				if (hTKKCMCCellInfo != null)
				{
					GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex);
				}
				FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + nextRow, num7, sheetView, num6, num7, num8, col);
				UpdateCtForDoiTac(text4, dataLocationTemplateRowStart + nextRow);
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num19 - 1].Value = MCTPath;
				if (!string.IsNullOrEmpty(text3))
				{
					ChangeCellTypeHoachToan(dataLocationTemplateRowStart + nextRow, num7);
					SetLoaiTienStatic(sheetView2, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
				}
				GeneralFunctions.SetRowHeight(sheetView, dataLocationTemplateRowStart + nextRow);
				nextRow++;
			}
			if (KHBS_LanTruoc != null)
			{
				object obj = null;
				string text6 = "";
				XmlNode val = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] == null)));
				string text7 = TimTieuMucStaticLanTruoc(val, KHBS_LanTruoc);
				HTKKCMCCellInfo hTKKCMCCellInfo5 = KHBS_LanTruoc?.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = KHBS_LanTruoc?.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == DVTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo3 = KHBS_LanTruoc?.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maDVTPath);
				string value2 = ((hTKKCMCCellInfo2 == null) ? "" : hTKKCMCCellInfo2.Value);
				string value3 = ((hTKKCMCCellInfo3 == null) ? "" : hTKKCMCCellInfo3.Value);
				obj = ((hTKKCMCCellInfo5 != null && hTKKCMCCellInfo5.Controltype == "16") ? "0" : "");
				text6 = obj.ToString().Trim();
				string text10 = text2;
				if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(text7))
				{
					text10 = text2 + " - Tiểu mục " + text7;
				}
				text10 = text10.Trim();
				if (checkExitsTrongsView(sheetView, hTKKCMCTableInfo, num3, text10, num5, text))
				{
					return nextRow;
				}
				string text9 = ((hTKKCMCCellInfo5 == null) ? "" : FormatValueLanTruocStatic(hTKKCMCCellInfo5, (hTKKCMCCellInfo5 != null) ? hTKKCMCCellInfo5.Value : ""));
				text9 = text9.Trim();
				if (text9 == text6 || (!string.IsNullOrEmpty(text9) && GeneralFunctions.IsNumeric(text9) && GeneralFunctions.IsNumeric(text6) && Convert.ToDouble(text9) == Convert.ToDouble(text6)))
				{
					return nextRow;
				}
				if (nextRow > 0)
				{
					GeneralFunctions.InsertTableRow(sheetView, dataLocationTemplateRowStart + nextRow - 1, -1, dataLocationTemplateRowStart + nextRow - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS], toTableindex);
				}
				ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text3);
				TextCellType textCellType = new TextCellType();
				ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + nextRow, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text3);
				if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(text7))
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text2;
				}
				else
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num3 - 1].Value = text2 + " - Tiểu mục " + text7;
				}
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num4 - 1].Value = text;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num5 - 1].Value = text;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num11 - 1].Value = text;
				if (hTKKCMCCellInfo5 != null && hTKKCMCCellInfo5.Controltype == "16")
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = ((hTKKCMCCellInfo5 != null && GeneralFunctions.IsNumeric(text9)) ? Convert.ToDouble(text9) : 0.0);
				}
				else
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num6 - 1].Value = ((hTKKCMCCellInfo5 != null) ? text9 : "");
				}
				if (hTKKCMCCellInfo5 != null && hTKKCMCCellInfo5.Controltype == "16")
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = ((obj != null && GeneralFunctions.IsNumeric(obj.ToString())) ? obj : ((object)0));
				}
				else
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num7 - 1].Value = ((obj != null) ? obj : "");
				}
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num9 - 1].Value = text7;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num20 - 1].Value = num;
				if (!string.IsNullOrEmpty(text3) && hTKKCMCCellInfo2 != null)
				{
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num13 - 1].Value = value2;
					sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num14 - 1].Value = value3;
				}
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num12 - 1].Value = ((hTKKCMCCellInfo5 != null) ? hTKKCMCCellInfo5.Controltype : "");
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num15 - 1].Value = text4;
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num16 - 1].Value = (string.IsNullOrEmpty(text5) ? "0" : text5);
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num8 - 1].Value = sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num18 - 1].Value;
				if (hTKKCMCCellInfo5 != null)
				{
					GeneralFunctions.FormatCol4567KHBSBySinfoStatic(hTKKCMCCellInfo5, sheetView, dataLocationTemplateRowStart + nextRow, num6, num7, num8, col);
				}
				UpdateCtForDoiTac(text4, dataLocationTemplateRowStart + nextRow);
				sheetView.Cells[dataLocationTemplateRowStart + nextRow - 1, num19 - 1].Value = MCTPath;
				if (!string.IsNullOrEmpty(text3))
				{
					ChangeCellTypeHoachToan(dataLocationTemplateRowStart + nextRow, num7);
				}
				GeneralFunctions.SetRowHeight(sheetView, dataLocationTemplateRowStart + nextRow);
				nextRow++;
			}
			AnHienSoThueThuHoiHoanTienNopCham_01_KHBS();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return nextRow;
	}

	public void SetDefaultManualRow_01_1_KHBS(SheetView sView, int tableIndex, int row)
	{
		try
		{
			int item = HTKKFile.ColumnLettersToIndex("H");
			int item2 = HTKKFile.ColumnLettersToIndex("P");
			int item3 = HTKKFile.ColumnLettersToIndex("R");
			int item4 = HTKKFile.ColumnLettersToIndex("V");
			int num = HTKKFile.ColumnLettersToIndex("BN");
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			GeneralFunctions.LockUnLockCols(sView, row, list, locked: false, deleteData: false);
			sView.Cells[row - 1, num - 1].Value = 1;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void LockUnlock_01_1_KHBS()
	{
		try
		{
			int item = HTKKFile.ColumnLettersToIndex("H");
			int item2 = HTKKFile.ColumnLettersToIndex("P");
			int item3 = HTKKFile.ColumnLettersToIndex("R");
			int item4 = HTKKFile.ColumnLettersToIndex("V");
			int colHide = HTKKFile.ColumnLettersToIndex("Z");
			int num = HTKKFile.ColumnLettersToIndex("BN");
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num2];
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num2].Tables;
			List<int> list = new List<int>();
			list.Add(item);
			list.Add(item2);
			list.Add(item3);
			list.Add(item4);
			for (int i = 0; i < tables.Count; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < tables[i].RowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					if (text == "1")
					{
						GeneralFunctions.LockUnLockCols(sheetView, dataLocationTemplateRowStart + j, list, locked: false, deleteData: false);
					}
					if (text == "2")
					{
						GeneralFunctions.LockUnLockCol(sheetView, dataLocationTemplateRowStart + j, colHide, locked: false, deleteData: false);
					}
				}
				InsertComboboxCt3_01_1_KHBS(i);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void InsertComboboxCt3_01_1_KHBS(int tableIndex)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("BS");
			int col = HTKKFile.ColumnLettersToIndex("P");
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num2];
			if (tableIndex >= WinSettings.ToKhaiInfo.SheetInfos[num2].Tables.Count)
			{
				return;
			}
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num2].Tables[tableIndex];
			if (hTKKCMCTableInfo == null || hTKKCMCTableInfo.Header.StaticCells.Count == 0)
			{
				return;
			}
			int rowIndex = hTKKCMCTableInfo.Header.StaticCells[1].CellTemplate.RowIndex;
			string text = sheetView.Cells[rowIndex - 1, num - 1].Text.Trim();
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			Dictionary<string, string> dMMaChiTieuTheoPhuLucKHBS = GetDMMaChiTieuTheoPhuLucKHBS(int.Parse(text));
			if (dMMaChiTieuTheoPhuLucKHBS.Count != 0)
			{
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
				{
					Functions.InsertCombobox(sheetView, dataLocationTemplateRowStart + i, col, dMMaChiTieuTheoPhuLucKHBS, deleteText: false);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void TinhSoThuePhaiNopKhauTruHoan_01_KHBS()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("AT");
			int num4 = HTKKFile.ColumnLettersToIndex("O");
			int num5 = HTKKFile.ColumnLettersToIndex("AU");
			int num6 = HTKKFile.ColumnLettersToIndex("U");
			int num7 = HTKKFile.ColumnLettersToIndex("AV");
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			int toTableIndex = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/ChiTiet");
			int toTableIndex2 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_II/ChiTiet");
			int toTableIndex3 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_III/ChiTiet");
			ChuyenTu_01_1_KHBS_Sang_01_KHBS(toTableIndex, "0");
			ChuyenTu_01_1_KHBS_Sang_01_KHBS(toTableIndex2, "1");
			ChuyenTu_01_1_KHBS_Sang_01_KHBS(toTableIndex3, "2");
			if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				HTKKCMCTableInfo tableTo = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/TongCong10");
				HTKKCMCTableInfo tableTo2 = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_II/TongCong12");
				HTKKCMCTableInfo tableTo3 = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_III/TongCong13");
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				dictionary.Add(num4, num4);
				GeneralFunctions.CopyPaseTableToTable(tables[toTableIndex], tableTo, sheetView, sheetView, dictionary, num7, num7, num6, num6, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
				GeneralFunctions.CopyPaseTableToTable(tables[toTableIndex2], tableTo2, sheetView, sheetView, dictionary, num7, num7, num6, num6, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
				GeneralFunctions.CopyPaseTableToTable(tables[toTableIndex3], tableTo3, sheetView, sheetView, dictionary, num7, num7, num6, num6, "", hTKKCMCSheetInfo, -1, "", usingCellType: true);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMPhuLucPhanBoKHBS()
	{
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			if (_mapMct != null)
			{
				foreach (XmlNode childNode in _mapMct.ChildNodes)
				{
					XmlNode val = childNode;
					int num3 = ((val.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value));
					string value = ((val.Attributes["SheetName"] == null) ? "" : ((XmlNode)val.Attributes["SheetName"]).Value);
					int num4 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num : num2));
					if (num2 == num4 && !dictionary.ContainsKey(num3.ToString()) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(num3.ToString(), value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	private void CollectionDynamicPhanBo_01_KHBS()
	{
		//IL_033e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0345: Expected O, but got Unknown
		try
		{
			if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				return;
			}
			int colStt = HTKKFile.ColumnLettersToIndex("B");
			int colCt = HTKKFile.ColumnLettersToIndex("E");
			int colMaHoSo = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetIndex_01_1_KHBS = HTKKFile.ColumnLettersToIndex("AU");
			int num = HTKKFile.ColumnLettersToIndex("AV");
			int num2 = HTKKFile.ColumnLettersToIndex("AW");
			int num3 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			int num4 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num3];
			SheetView sheetView = winFpSpread.Sheets[num3];
			int tableIndex = 1;
			int num5 = hTKKCMCSheetInfo.Tables.Count - 2;
			for (int num6 = num5 - 1; num6 >= 1; num6--)
			{
				if (num6 < hTKKCMCSheetInfo.Tables.Count - 1 && hTKKCMCSheetInfo.Tables[num6] != null)
				{
					int tableTemplateRowStart = hTKKCMCSheetInfo.Tables[num6].Location.TableTemplateRowStart;
					GeneralFunctions.RemoveTableF8New(sheetView, hTKKCMCSheetInfo, num6, setActiveCell: true);
				}
			}
			if (hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count > 1)
			{
				int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart;
				GeneralFunctions.RemoveDataRows(sheetView, tableIndex, dataLocationTemplateRowStart - 1, hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count - 1, hTKKCMCSheetInfo);
			}
			if (hTKKCMCSheetInfo.Tables[tableIndex].RowsValues.Count > 0)
			{
				int dataLocationTemplateRowStart = hTKKCMCSheetInfo.Tables[tableIndex].Location.DataLocationTemplateRowStart;
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCSheetInfo.Tables[tableIndex].RowsInfo.Cells)
				{
					sheetView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
			int num7 = 0;
			int num8 = 0;
			foreach (KeyValuePair<string, string> dMPhuLucPhanBoKHB in GetDMPhuLucPhanBoKHBS())
			{
				num8 = Convert.ToInt32(dMPhuLucPhanBoKHB.Key);
				if (num7 > 0)
				{
					GeneralFunctions.InsertTableF7(sheetView, hTKKCMCSheetInfo, num7);
				}
				IEnumerable<HTKKCMCTableInfo> source = WinSettings.ToKhaiInfo.SheetInfos[num3].Tables.Where((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_2/DSachPLuc/BKePLuc");
				tableIndex = source.Count();
				HTKKCMCTableInfo table = WinSettings.ToKhaiInfo.SheetInfos[num3].Tables[tableIndex];
				SetSheetNameToTableHeader(sheetView, table, dMPhuLucPhanBoKHB, colCt, colSheetIndex_01_1_KHBS, colStt, tableIndex + 1, colMaHoSo);
				foreach (XmlNode childNode in _mapMct.ChildNodes)
				{
					XmlNode val = childNode;
					int num9 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num4 : num3));
					string s = ((val.Attributes["SheetIndex"] != null) ? ((XmlNode)val.Attributes["SheetIndex"]).Value.Trim() : "");
					string text = ((val.Attributes["TenTinhPath"] != null) ? ((XmlNode)val.Attributes["TenTinhPath"]).Value.ToString().Trim() : "");
					string text2 = ((val.Attributes["PhanBoTheoTSC"] != null) ? ((XmlNode)val.Attributes["PhanBoTheoTSC"]).Value.ToString().Trim() : "");
					if (num8 == int.Parse(s) && num9 == num3)
					{
						if (val.Attributes["TenTinhPath"] != null)
						{
							CollectionDetailSoThuePhaiNopTungTinhHeader_01_KHBS(val, int.Parse(s), tableIndex);
							CollectionDetailSoThuePhaiNopTungTinh_01_KHBS(val, int.Parse(s), tableIndex);
						}
						else if (val.Attributes["PhanBoTheoTSC"] != null)
						{
							CollectionDetailSoThuePhaiNopTungTinh_Static_01_KHBS(val, int.Parse(s), tableIndex);
						}
					}
				}
				num7++;
			}
			IEnumerable<HTKKCMCTableInfo> source2 = WinSettings.ToKhaiInfo.SheetInfos[num3].Tables.Where((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_2/DSachPLuc/BKePLuc");
			RemoveTableF7IfNull(sheetView, hTKKCMCSheetInfo, source2.ToList());
			foreach (HTKKCMCTableInfo item in source2.ToList())
			{
				GeneralFunctions.fpSpReindexingTableTemplate(sheetView, item);
			}
			GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num3]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CollectionDetailSoThuePhaiNopTungTinh_01_KHBS(XmlNode mapCt, int sheetIndexToDieuChinh, int toTableIndex)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("BA");
			int num5 = HTKKFile.ColumnLettersToIndex("BC");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("BB");
			int num8 = HTKKFile.ColumnLettersToIndex("P");
			int num9 = HTKKFile.ColumnLettersToIndex("AO");
			int num10 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "V" : "T");
			int num11 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "BE" : "AP");
			int num12 = HTKKFile.ColumnLettersToIndex("X");
			int num13 = HTKKFile.ColumnLettersToIndex("AQ");
			int num14 = HTKKFile.ColumnLettersToIndex("AB");
			int num15 = HTKKFile.ColumnLettersToIndex("AR");
			int num16 = HTKKFile.ColumnLettersToIndex("AF");
			int num17 = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetName = HTKKFile.ColumnLettersToIndex("AV");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("AW");
			int num18 = HTKKFile.ColumnLettersToIndex("BD");
			int num19 = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			string text = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string tablePath = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			string text2 = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string maTCTPath = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string MSTPath = ((mapCt.Attributes["MSTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MSTPath"]).Value.ToString().Trim());
			string tenTinhPath = ((mapCt.Attributes["TenTinhPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenTinhPath"]).Value.ToString().Trim());
			string maTinhPath = ((mapCt.Attributes["MaTinhPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTinhPath"]).Value.ToString().Trim());
			string tenHuyenPath = ((mapCt.Attributes["TenHuyenPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenHuyenPath"]).Value.ToString().Trim());
			string maHuyenPath = ((mapCt.Attributes["MaHuyenPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaHuyenPath"]).Value.ToString().Trim());
			string tenCQTCapCucPath = ((mapCt.Attributes["TenCQTCapCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenCQTCapCucPath"]).Value.ToString().Trim());
			string maCQTCapCucPath = ((mapCt.Attributes["MaCQTCapCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaCQTCapCucPath"]).Value.ToString().Trim());
			string tenCQTChiCucPath = ((mapCt.Attributes["TenCQTChiCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenCQTChiCucPath"]).Value.ToString().Trim());
			string maCQTChiCucPath = ((mapCt.Attributes["MaCQTChiCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaCQTChiCucPath"]).Value.ToString().Trim());
			string valueLoaiThue = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string text3 = ((_mapMct != null && _mapMct.Attributes["SuaPhanBo"] != null) ? ((XmlNode)_mapMct.Attributes["SuaPhanBo"]).Value : "");
			string prefixPath = ((mapCt.Attributes["PrefixPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["PrefixPath"]).Value.Trim());
			if (num19 != sheetIndexToDieuChinh)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinFnProcess.ReadSheet_LanTruoc(num19);
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num].Tables[toTableIndex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num20 = hTKKCMCTableInfo.RowsValues.Count - 1;
			SheetView sheetView2 = ((WinSettings.ToKhaiInfo.SheetInfos[num19] == null) ? null : winFpSpread.Sheets[num19]);
			List<HTKKCMCTableInfo> list = ((WinSettings.ToKhaiInfo.SheetInfos[num19] == null) ? new List<HTKKCMCTableInfo>() : WinSettings.ToKhaiInfo.SheetInfos[num19].Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList());
			foreach (HTKKCMCTableInfo item in list)
			{
				int dataLocationTemplateRowStart2 = item.Location.DataLocationTemplateRowStart;
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
				int num21 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = (string.IsNullOrEmpty(MSTPath) ? null : item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MSTPath));
				int num22 = (string.IsNullOrEmpty(MSTPath) ? (-1) : item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MSTPath));
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo3 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
				int num23 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo4 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenTinhPath);
				int num24 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenTinhPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo5 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTinhPath);
				int num25 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTinhPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo6 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenHuyenPath);
				int num26 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenHuyenPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo7 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maHuyenPath);
				int num27 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maHuyenPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo8 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenCQTCapCucPath);
				int num28 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenCQTCapCucPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo9 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maCQTCapCucPath);
				int num29 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maCQTCapCucPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo10 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenCQTChiCucPath);
				int num30 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenCQTChiCucPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo11 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maCQTChiCucPath);
				int num31 = item.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maCQTChiCucPath);
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo12 = item.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
				if (hTKKCMCTableCellInfo3 == null)
				{
					continue;
				}
				ICellType cellType = sheetView.Cells[0, 0].CellType;
				for (int num32 = 0; num32 < item.RowsValues.Count; num32++)
				{
					try
					{
						if (sheetView2 == null || CheckTruongHopDacBietPhanBo_01_KHBSS_InView(sheetView2, num19, dataLocationTemplateRowStart2 + num32))
						{
							continue;
						}
						string text4 = ((hTKKCMCTableCellInfo3 != null) ? Convert.ToString(sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo3.CellTemplate.ColIndex - 1].Value) : "");
						string tenTieuMucTuMa = GetTenTieuMucTuMa(text4);
						if (tenTieuMucTuMa == null || (tenTieuMucTuMa != null && string.IsNullOrEmpty(tenTieuMucTuMa.ToString())))
						{
							continue;
						}
						double num33 = Convert.ToDouble(sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value);
						if (hTKKCMCTableCellInfo4 == null || hTKKCMCTableCellInfo6 == null || hTKKCMCTableCellInfo8 == null || hTKKCMCTableCellInfo10 == null)
						{
							continue;
						}
						object obj = ((hTKKCMCTableCellInfo2 == null) ? null : sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo2.CellTemplate.ColIndex - 1].Value);
						object value = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo4.CellTemplate.ColIndex - 1].Value;
						object value2 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo6.CellTemplate.ColIndex - 1].Value;
						object value3 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo8.CellTemplate.ColIndex - 1].Value;
						object value4 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo10.CellTemplate.ColIndex - 1].Value;
						object value5 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo5.CellTemplate.ColIndex - 1].Value;
						object value6 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo7.CellTemplate.ColIndex - 1].Value;
						object value7 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo9.CellTemplate.ColIndex - 1].Value;
						object value8 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo11.CellTemplate.ColIndex - 1].Value;
						string text5 = "";
						if (hTKKCMCTableCellInfo12 != null)
						{
							text5 = sheetView2.Cells[dataLocationTemplateRowStart2 + num32 - 1, hTKKCMCTableCellInfo12.CellTemplate.ColIndex - 1].Text.Trim();
						}
						if (string.IsNullOrEmpty(num33.ToString()))
						{
							if (item.Header.StaticCells.Count != 0)
							{
							}
						}
						else
						{
							if ((text3 == "" && (value == null || value2 == null || value3 == null || value4 == null)) || (text3 == "1" && value3 == null))
							{
								continue;
							}
							List<int> list2 = new List<int>();
							list2.Clear();
							list2.Add(hTKKCMCTableCellInfo3.CellTemplate.ColIndex);
							if (obj != null)
							{
								list2.Add(hTKKCMCTableCellInfo2.CellTemplate.ColIndex);
							}
							list2.Add(hTKKCMCTableCellInfo5.CellTemplate.ColIndex);
							list2.Add(hTKKCMCTableCellInfo7.CellTemplate.ColIndex);
							list2.Add(hTKKCMCTableCellInfo9.CellTemplate.ColIndex);
							list2.Add(hTKKCMCTableCellInfo11.CellTemplate.ColIndex);
							if (hTKKCMCTableCellInfo12 != null)
							{
								list2.Add(hTKKCMCTableCellInfo12.CellTemplate.ColIndex);
							}
							List<int> list3 = new List<int>();
							list3.Add(num4);
							if (obj != null)
							{
								list3.Add(num6);
							}
							list3.Add(num9);
							list3.Add(num11);
							list3.Add(num13);
							list3.Add(num15);
							if (!string.IsNullOrEmpty(prefixPath))
							{
								list3.Add(num18);
							}
							List<string> list4 = new List<string>();
							list4.Add(maTCTPath);
							if (obj != null)
							{
								list4.Add(MSTPath);
							}
							list4.Add(maTinhPath);
							list4.Add(maHuyenPath);
							list4.Add(maCQTCapCucPath);
							list4.Add(maCQTChiCucPath);
							if (!string.IsNullOrEmpty(prefixPath))
							{
								list4.Add(prefixPath);
							}
							List<string> list5 = new List<string>();
							list5.Add((text4 == null) ? "" : text4.ToString());
							if (obj != null)
							{
								list5.Add(obj.ToString());
							}
							list5.Add((value5 == null) ? "" : value5.ToString());
							list5.Add((value6 == null) ? "" : value6.ToString());
							list5.Add((value7 == null) ? "" : value7.ToString());
							list5.Add((value8 == null) ? "" : value8.ToString());
							if (hTKKCMCTableCellInfo12 != null)
							{
								list5.Add(text5);
							}
							num33 = TinhSumColMultiTableInSview(sheetView2, list, list2, list5, item.TablePath, hTKKCMCTableCellInfo.CellTemplate.ColIndex);
							bool exists = true;
							double num34 = ((hTKKCMCSheetInfo == null) ? 0.0 : TinhSumColMultiTableInSinfo(hTKKCMCSheetInfo.Tables, list4, list5, item.TablePath, MCTPath, ref exists));
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list3, list5);
							if (rowIndexInSViewByCompares.Count <= 0 && num34 != num33)
							{
								string value9 = sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Text.Trim();
								if (num20 == 0 && !string.IsNullOrEmpty(value9))
								{
									num20++;
								}
								if (num20 > 0)
								{
									WinSpreadProcess.InsertTableRow(sheetView, dataLocationTemplateRowStart + num20 - 1, -1, dataLocationTemplateRowStart + num20 - 1, -1, -1, sInfo, toTableIndex);
								}
								ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + num20, colSheetIndex, num19, colSheetName, text, colLoaiThue, valueLoaiThue);
								if (string.IsNullOrEmpty(text5))
								{
									sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text4, tenTieuMucTuMa.ToString(), " - ");
								}
								else
								{
									List<string> list6 = new List<string>();
									list6.Add(text5);
									list6.Add(text4);
									list6.Add(tenTieuMucTuMa.ToString());
									string value10 = GeneralFunctions.ListToString(list6, " - ");
									sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = value10;
								}
								string value11 = ((!string.IsNullOrEmpty(text5)) ? GeneralFunctions.FormatKey2(text5, tenTieuMucTuMa, " - ") : tenTieuMucTuMa);
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num4 - 1].Value = text4;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num5 - 1].Value = value11;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num6 - 1].Value = obj;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num7 - 1].Value = obj;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num8 - 1].Value = ((value != null) ? value : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num9 - 1].Value = ((value5 != null) ? value5 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num10 - 1].Value = ((value2 != null) ? value2 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num11 - 1].Value = ((value6 != null) ? value6 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num12 - 1].Value = ((value3 != null) ? value3 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num13 - 1].Value = ((value7 != null) ? value7 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num14 - 1].Value = ((value4 != null) ? value4 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num15 - 1].Value = ((value8 != null) ? value8 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num16 - 1].Value = num33 - num34;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num18 - 1].Value = text5;
								num20++;
							}
							continue;
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
			}
			if (hTKKCMCSheetInfo != null)
			{
				foreach (HTKKCMCTableInfo tableKHBSLanTruoc in hTKKCMCSheetInfo.Tables)
				{
					if (tableKHBSLanTruoc == null || tableKHBSLanTruoc.RowsValues.Count <= 0)
					{
						continue;
					}
					foreach (HTKKCMCRowValue rowsValue in tableKHBSLanTruoc.RowsValues)
					{
						if (CheckTruongHopDacBietPhanBo_01_KHBSS_InTable(num19, tableKHBSLanTruoc, rowsValue))
						{
							continue;
						}
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
						int num21 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
						int num22 = (string.IsNullOrEmpty(MSTPath) ? (-1) : tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MSTPath));
						int num23 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
						int num24 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenTinhPath);
						int num25 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maTinhPath);
						int num26 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenHuyenPath);
						int num27 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maHuyenPath);
						int num28 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenCQTCapCucPath);
						int num29 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maCQTCapCucPath);
						int num30 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == tenCQTChiCucPath);
						int num31 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == maCQTChiCucPath);
						int num35 = -1;
						if (!string.IsNullOrEmpty(prefixPath))
						{
							num35 = tableKHBSLanTruoc.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
						}
						string text6 = ((num23 >= 0) ? rowsValue.RowValue[num23] : "");
						string text7 = ((num22 >= 0) ? rowsValue.RowValue[num22] : "");
						string text8 = ((num24 >= 0) ? rowsValue.RowValue[num24] : "");
						string text9 = ((num25 >= 0) ? rowsValue.RowValue[num25] : "");
						string text10 = ((num26 >= 0) ? rowsValue.RowValue[num26] : "");
						string text11 = ((num27 >= 0) ? rowsValue.RowValue[num27] : "");
						string text12 = ((num28 >= 0) ? rowsValue.RowValue[num28] : "");
						string text13 = ((num29 >= 0) ? rowsValue.RowValue[num29] : "");
						string text14 = ((num30 >= 0) ? rowsValue.RowValue[num30] : "");
						string text15 = ((num31 >= 0) ? rowsValue.RowValue[num31] : "");
						string text16 = "";
						if (!string.IsNullOrEmpty(prefixPath))
						{
							text16 = ((num35 >= 0) ? rowsValue.RowValue[num35] : "");
						}
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = (string.IsNullOrEmpty(MSTPath) ? null : tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MSTPath));
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo3 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTCTPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo4 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenTinhPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo5 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maTinhPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo6 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenHuyenPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo7 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maHuyenPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo8 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenCQTCapCucPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo9 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maCQTCapCucPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo10 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == tenCQTChiCucPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo11 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == maCQTChiCucPath);
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo12 = tableKHBSLanTruoc.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
						if (string.IsNullOrEmpty(text6))
						{
							if (tableKHBSLanTruoc.Header.StaticCells.Count == 0)
							{
								HTKKFile.Log(text + " body xem lai path MCT: " + MCTPath);
							}
						}
						else
						{
							if ((text3 == "" && (string.IsNullOrEmpty(text8.Trim()) || string.IsNullOrEmpty(text10.Trim()) || string.IsNullOrEmpty(text12.Trim()) || string.IsNullOrEmpty(text14.Trim()))) || (text3 == "1" && string.IsNullOrEmpty(text12.Trim())))
							{
								continue;
							}
							string tenTieuMucTuMa2 = GetTenTieuMucTuMa(text6);
							List<int> list3 = new List<int>();
							list3.Add(num4);
							if (!string.IsNullOrEmpty(text7))
							{
								list3.Add(num6);
							}
							list3.Add(num9);
							list3.Add(num11);
							list3.Add(num13);
							list3.Add(num15);
							if (!string.IsNullOrEmpty(prefixPath))
							{
								list3.Add(num18);
							}
							List<string> list4 = new List<string>();
							list4.Add(maTCTPath);
							if (!string.IsNullOrEmpty(text7))
							{
								list4.Add(MSTPath);
							}
							list4.Add(maTinhPath);
							list4.Add(maHuyenPath);
							list4.Add(maCQTCapCucPath);
							list4.Add(maCQTChiCucPath);
							if (!string.IsNullOrEmpty(prefixPath))
							{
								list4.Add(prefixPath);
							}
							List<string> list5 = new List<string>();
							list5.Add(text6);
							if (!string.IsNullOrEmpty(text7))
							{
								list5.Add(text7);
							}
							list5.Add(text9);
							list5.Add(text11);
							list5.Add(text13);
							list5.Add(text15);
							if (!string.IsNullOrEmpty(prefixPath))
							{
								list5.Add(text16);
							}
							bool exists = true;
							double num34 = TinhSumColMultiTableInSinfo(hTKKCMCSheetInfo.Tables, list4, list5, tableKHBSLanTruoc.TablePath, MCTPath, ref exists);
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list3, list5);
							if (rowIndexInSViewByCompares.Count > 0)
							{
								continue;
							}
							double num36 = 0.0;
							if (sheetView2 != null)
							{
								List<int> list2 = new List<int>();
								list2.Add(hTKKCMCTableCellInfo3.CellTemplate.ColIndex);
								if (!string.IsNullOrEmpty(text7))
								{
									list2.Add(hTKKCMCTableCellInfo2.CellTemplate.ColIndex);
								}
								list2.Add(hTKKCMCTableCellInfo5.CellTemplate.ColIndex);
								list2.Add(hTKKCMCTableCellInfo7.CellTemplate.ColIndex);
								list2.Add(hTKKCMCTableCellInfo9.CellTemplate.ColIndex);
								list2.Add(hTKKCMCTableCellInfo11.CellTemplate.ColIndex);
								if (hTKKCMCTableCellInfo12 != null)
								{
									list2.Add(hTKKCMCTableCellInfo12.CellTemplate.ColIndex);
								}
								list = ((WinSettings.ToKhaiInfo.SheetInfos[num19] == null) ? new List<HTKKCMCTableInfo>() : WinSettings.ToKhaiInfo.SheetInfos[num19].Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tableKHBSLanTruoc.TablePath).ToList());
								num36 = TinhSumColMultiTableInSview(sheetView2, list, list2, list5, tableKHBSLanTruoc.TablePath, hTKKCMCTableCellInfo.CellTemplate.ColIndex);
							}
							if (num34 != num36)
							{
								string value9 = sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Text.Trim();
								if (num20 == 0 && !string.IsNullOrEmpty(value9))
								{
									num20++;
								}
								if (num20 > 0)
								{
									WinSpreadProcess.InsertTableRow(sheetView, dataLocationTemplateRowStart + num20 - 1, -1, dataLocationTemplateRowStart + num20 - 1, -1, -1, sInfo, toTableIndex);
								}
								ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + num20, colSheetIndex, num19, colSheetName, text, colLoaiThue, valueLoaiThue);
								if (string.IsNullOrEmpty(text16))
								{
									sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text6, tenTieuMucTuMa2, " - ");
								}
								else
								{
									List<string> list6 = new List<string>();
									list6.Add(text16);
									list6.Add(text6);
									list6.Add(tenTieuMucTuMa2);
									string value10 = GeneralFunctions.ListToString(list6, " - ");
									sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = value10;
								}
								string value11 = ((!string.IsNullOrEmpty(text16)) ? GeneralFunctions.FormatKey2(text16, tenTieuMucTuMa2, " - ") : tenTieuMucTuMa2);
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num5 - 1].Value = value11;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num4 - 1].Value = text6;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num6 - 1].Value = text7;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num7 - 1].Value = text7;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num8 - 1].Value = text8;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num9 - 1].Value = text9;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num10 - 1].Value = text10;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num11 - 1].Value = text11;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num12 - 1].Value = text12;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num13 - 1].Value = text13;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num14 - 1].Value = text14;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num15 - 1].Value = text15;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num16 - 1].Value = num36 - num34;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num18 - 1].Value = text16;
								num20++;
							}
						}
					}
				}
			}
			GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public double TinhSumColMultiTableInSview(SheetView sView, List<HTKKCMCTableInfo> tables, List<int> colChecks, List<string> valueCompares, string tablePath, int colSum)
	{
		double num = 0.0;
		if (tables.Count == 0)
		{
			return num;
		}
		foreach (HTKKCMCTableInfo table in tables)
		{
			if (table.TablePath != tablePath)
			{
				continue;
			}
			List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sView, table, colChecks, valueCompares);
			if (rowIndexInSViewByCompares.Count <= 0)
			{
				continue;
			}
			foreach (int item in rowIndexInSViewByCompares)
			{
				num += Convert.ToDouble(sView.Cells[item - 1, colSum - 1].Value);
			}
		}
		return num;
	}

	public double TinhSumColMultiTableInSinfo(List<HTKKCMCTableInfo> tables, List<string> listPath, List<string> listValues, string tablePath, string MCTPath, ref bool exists)
	{
		try
		{
			double num = 0.0;
			if (tables.Count == 0)
			{
				exists = false;
				return num;
			}
			List<HTKKCMCTableInfo> tableByKeyBody = GeneralFunctions.GetTableByKeyBody(tables, listPath, listValues, tablePath);
			if (tableByKeyBody == null)
			{
				exists = false;
				return num;
			}
			foreach (HTKKCMCTableInfo item in tableByKeyBody)
			{
				int num2 = item?.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == MCTPath) ?? (-1);
				List<HTKKCMCRowValue> list = ((item == null) ? null : GeneralFunctions.GetRowValueInTable(item, listPath, listValues));
				if (list == null)
				{
					continue;
				}
				foreach (HTKKCMCRowValue item2 in list)
				{
					num += ((item2 != null && num2 >= 0) ? Convert.ToDouble(item2.RowValue[num2]) : 0.0);
				}
			}
			exists = true;
			return num;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			exists = false;
			return 0.0;
		}
	}

	public bool CheckTruongHopDacBietPhanBo_01_KHBSS_InView(SheetView sView, int sheetIndex, int iRow)
	{
		try
		{
			if (WinSettings.MenuID == "N8" && sheetIndex == 12)
			{
				string loaiPath = "loai";
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables[0].RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == loaiPath);
				if (hTKKCMCTableCellInfo == null)
				{
					return false;
				}
				if (sView.Cells[iRow - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text.Trim() == "03")
				{
					return true;
				}
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public bool CheckTruongHopDacBietPhanBo_01_KHBSS_InTable(int sheetIndex, HTKKCMCTableInfo table, HTKKCMCRowValue row)
	{
		try
		{
			if (WinSettings.MenuID == "N8" && sheetIndex == 12)
			{
				string loaiPath = "loai";
				int num = table.RowsInfo.Cells.FindIndex((HTKKCMCTableCellInfo x) => x.Path == loaiPath);
				string text = ((num >= 0) ? row.RowValue[num] : "");
				if (text.Trim() == "03")
				{
					return true;
				}
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	private void CollectionDetailSoThuePhaiNopTungTinhHeader_01_KHBS(XmlNode mapCt, int sheetIndexToDieuChinh, int toTableIndex)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			ICellType cellType = null;
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("BA");
			int num5 = HTKKFile.ColumnLettersToIndex("BC");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("BB");
			int num8 = HTKKFile.ColumnLettersToIndex("P");
			int num9 = HTKKFile.ColumnLettersToIndex("AO");
			int num10 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "V" : "T");
			int num11 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "BE" : "AP");
			int num12 = HTKKFile.ColumnLettersToIndex("X");
			int num13 = HTKKFile.ColumnLettersToIndex("AQ");
			int num14 = HTKKFile.ColumnLettersToIndex("AB");
			int num15 = HTKKFile.ColumnLettersToIndex("AR");
			int num16 = HTKKFile.ColumnLettersToIndex("AF");
			int num17 = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetName = HTKKFile.ColumnLettersToIndex("AV");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("AW");
			int num18 = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			string sheetname = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string tablePath = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			string text = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string maTCTPath = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string MSTPath = ((mapCt.Attributes["MSTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MSTPath"]).Value.ToString().Trim());
			string tenTinhPath = ((mapCt.Attributes["TenTinhPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenTinhPath"]).Value.ToString().Trim());
			string maTinhPath = ((mapCt.Attributes["MaTinhPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTinhPath"]).Value.ToString().Trim());
			string tenHuyenPath = ((mapCt.Attributes["TenHuyenPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenHuyenPath"]).Value.ToString().Trim());
			string maHuyenPath = ((mapCt.Attributes["MaHuyenPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaHuyenPath"]).Value.ToString().Trim());
			string tenCQTCapCucPath = ((mapCt.Attributes["TenCQTCapCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenCQTCapCucPath"]).Value.ToString().Trim());
			string maCQTCapCucPath = ((mapCt.Attributes["MaCQTCapCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaCQTCapCucPath"]).Value.ToString().Trim());
			string tenCQTChiCucPath = ((mapCt.Attributes["TenCQTChiCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TenCQTChiCucPath"]).Value.ToString().Trim());
			string maCQTChiCucPath = ((mapCt.Attributes["MaCQTChiCucPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaCQTChiCucPath"]).Value.ToString().Trim());
			string valueLoaiThue = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			if (num18 != sheetIndexToDieuChinh || WinSettings.ToKhaiInfo.SheetInfos[num18] == null)
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinFnProcess.ReadSheet_LanTruoc(num18);
			List<HTKKCMCTableInfo> list = ((hTKKCMCSheetInfo != null) ? hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TablePath == tablePath).ToList() : new List<HTKKCMCTableInfo>());
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num].Tables[toTableIndex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			SheetView sheetView2 = winFpSpread.Sheets[num18];
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num18].Tables;
			int num19 = hTKKCMCTableInfo.RowsValues.Count - 1;
			foreach (HTKKCMCTableInfo item in tables)
			{
				try
				{
					int dataLocationTemplateRowStart2 = item.Location.DataLocationTemplateRowStart;
					HTKKCMCCellInfo cellMCTToKhaiDieuChinh = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
					int num20 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == MCTPath);
					HTKKCMCCellInfo hTKKCMCCellInfo = (string.IsNullOrEmpty(MSTPath) ? null : item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MSTPath));
					int num21 = (string.IsNullOrEmpty(MSTPath) ? (-1) : item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == MSTPath));
					HTKKCMCCellInfo hTKKCMCCellInfo2 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTCTPath);
					int num22 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maTCTPath);
					HTKKCMCCellInfo hTKKCMCCellInfo3 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenTinhPath);
					int num23 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == tenTinhPath);
					HTKKCMCCellInfo hTKKCMCCellInfo4 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTinhPath);
					int num24 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maTinhPath);
					HTKKCMCCellInfo hTKKCMCCellInfo5 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenHuyenPath);
					int num25 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == tenHuyenPath);
					HTKKCMCCellInfo hTKKCMCCellInfo6 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maHuyenPath);
					int num26 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maHuyenPath);
					HTKKCMCCellInfo hTKKCMCCellInfo7 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenCQTCapCucPath);
					int num27 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == tenCQTCapCucPath);
					HTKKCMCCellInfo hTKKCMCCellInfo8 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maCQTCapCucPath);
					int num28 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maCQTCapCucPath);
					HTKKCMCCellInfo hTKKCMCCellInfo9 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenCQTChiCucPath);
					int num29 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == tenCQTChiCucPath);
					HTKKCMCCellInfo hTKKCMCCellInfo10 = item.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maCQTChiCucPath);
					int num30 = item.Header.StaticCells.FindIndex((HTKKCMCCellInfo x) => x.Path == maCQTChiCucPath);
					if (cellMCTToKhaiDieuChinh == null)
					{
						if (item.Header.StaticCells.Count <= 0)
						{
						}
						continue;
					}
					ICellType cellType2 = sheetView.Cells[0, 0].CellType;
					List<int> list2 = new List<int>();
					list2.Add(hTKKCMCCellInfo2.CellTemplate.ColIndex);
					if (hTKKCMCCellInfo != null)
					{
						list2.Add(hTKKCMCCellInfo.CellTemplate.ColIndex);
					}
					list2.Add(hTKKCMCCellInfo4.CellTemplate.ColIndex);
					list2.Add(hTKKCMCCellInfo6.CellTemplate.ColIndex);
					list2.Add(hTKKCMCCellInfo8.CellTemplate.ColIndex);
					list2.Add(hTKKCMCCellInfo10.CellTemplate.ColIndex);
					if (sheetView2 == null)
					{
						continue;
					}
					string text2 = ((hTKKCMCCellInfo2 != null) ? Convert.ToString(sheetView2.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value) : "");
					string tenTieuMucTuMa = GetTenTieuMucTuMa(text2);
					if (tenTieuMucTuMa == null || string.IsNullOrEmpty(tenTieuMucTuMa.ToString()))
					{
						continue;
					}
					double num31 = Convert.ToDouble(sheetView2.Cells[cellMCTToKhaiDieuChinh.CellTemplate.RowIndex - 1, cellMCTToKhaiDieuChinh.CellTemplate.ColIndex - 1].Value);
					object obj = ((hTKKCMCCellInfo == null) ? null : sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
					object value = sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value;
					object value2 = sheetView2.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value;
					object value3 = sheetView2.Cells[hTKKCMCCellInfo7.CellTemplate.RowIndex - 1, hTKKCMCCellInfo7.CellTemplate.ColIndex - 1].Value;
					object value4 = sheetView2.Cells[hTKKCMCCellInfo9.CellTemplate.RowIndex - 1, hTKKCMCCellInfo9.CellTemplate.ColIndex - 1].Value;
					object value5 = sheetView2.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value;
					object value6 = sheetView2.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].Value;
					object value7 = sheetView2.Cells[hTKKCMCCellInfo8.CellTemplate.RowIndex - 1, hTKKCMCCellInfo8.CellTemplate.ColIndex - 1].Value;
					object value8 = sheetView2.Cells[hTKKCMCCellInfo10.CellTemplate.RowIndex - 1, hTKKCMCCellInfo10.CellTemplate.ColIndex - 1].Value;
					if (value == null || value2 == null || value3 == null || value4 == null)
					{
						continue;
					}
					List<int> list3 = new List<int>();
					list3.Add(num4);
					if (obj != null)
					{
						list3.Add(num6);
					}
					list3.Add(num9);
					list3.Add(num11);
					list3.Add(num13);
					list3.Add(num15);
					List<string> list4 = new List<string>();
					list4.Add(maTCTPath);
					if (obj != null)
					{
						list4.Add(MSTPath);
					}
					list4.Add(maTinhPath);
					list4.Add(maHuyenPath);
					list4.Add(maCQTCapCucPath);
					list4.Add(maCQTChiCucPath);
					List<string> list5 = new List<string>();
					list5.Add((text2 == null) ? "" : text2.ToString());
					if (obj != null)
					{
						list5.Add(obj.ToString());
					}
					list5.Add((value5 == null) ? "" : value5.ToString());
					list5.Add((value6 == null) ? "" : value6.ToString());
					list5.Add((value7 == null) ? "" : value7.ToString());
					list5.Add((value8 == null) ? "" : value8.ToString());
					HTKKCMCTableInfo tableByHeader = GeneralFunctions.GetTableByHeader(list, list4, list5);
					double num32 = 0.0;
					if (tableByHeader != null)
					{
						HTKKCMCCellInfo hTKKCMCCellInfo11 = tableByHeader.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == cellMCTToKhaiDieuChinh.Path);
						if (hTKKCMCCellInfo11 != null)
						{
							string value9 = FormatValueLanTruocStatic(cellMCTToKhaiDieuChinh, hTKKCMCCellInfo11.Value);
							if (!string.IsNullOrEmpty(value9))
							{
								num32 = Convert.ToDouble(value9);
							}
						}
					}
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list3, list5);
					if (rowIndexInSViewByCompares.Count > 0)
					{
						object value10 = sheetView.Cells[rowIndexInSViewByCompares[0] - 1, num16 - 1].Value;
						sheetView.Cells[rowIndexInSViewByCompares[0] - 1, num16 - 1].Value = ((value10 != null && !string.IsNullOrEmpty(value10.ToString().Trim())) ? (Convert.ToDouble(value10) + num31) : num31);
					}
					else if (num32 != num31)
					{
						if (num19 > 0)
						{
							WinSpreadProcess.InsertTableRow(sheetView, dataLocationTemplateRowStart + num19 - 1, -1, dataLocationTemplateRowStart + num19 - 1, -1, -1, sInfo, toTableIndex);
						}
						ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + num19, colSheetIndex, num18, colSheetName, sheetname, colLoaiThue, valueLoaiThue);
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text2, tenTieuMucTuMa.ToString(), " - ");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num4 - 1].Value = text2;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num5 - 1].Value = tenTieuMucTuMa;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num6 - 1].Value = obj;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num7 - 1].Value = obj;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num8 - 1].Value = ((value != null) ? value : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num9 - 1].Value = ((value5 != null) ? value5 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num10 - 1].Value = ((value2 != null) ? value2 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num11 - 1].Value = ((value6 != null) ? value6 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num12 - 1].Value = ((value3 != null) ? value3 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num13 - 1].Value = ((value7 != null) ? value7 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num14 - 1].Value = ((value4 != null) ? value4 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num15 - 1].Value = ((value8 != null) ? value8 : "");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num16 - 1].Value = num31 - num32;
						num19++;
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			if (list != null && list.Count > 0)
			{
				foreach (HTKKCMCTableInfo item2 in list)
				{
					try
					{
						if (item2 == null || item2.Header == null || item2.Header.StaticCells == null || item2.Header.StaticCells.Count <= 0)
						{
							continue;
						}
						HTKKCMCCellInfo cellMCTLanTruoc = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo12 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTCTPath);
						HTKKCMCCellInfo hTKKCMCCellInfo13 = (string.IsNullOrEmpty(MSTPath) ? null : item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MSTPath));
						HTKKCMCCellInfo hTKKCMCCellInfo14 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenTinhPath);
						HTKKCMCCellInfo hTKKCMCCellInfo15 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTinhPath);
						HTKKCMCCellInfo hTKKCMCCellInfo16 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenHuyenPath);
						HTKKCMCCellInfo hTKKCMCCellInfo17 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maHuyenPath);
						HTKKCMCCellInfo hTKKCMCCellInfo18 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenCQTCapCucPath);
						HTKKCMCCellInfo hTKKCMCCellInfo19 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maCQTCapCucPath);
						HTKKCMCCellInfo hTKKCMCCellInfo20 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == tenCQTChiCucPath);
						HTKKCMCCellInfo hTKKCMCCellInfo21 = item2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maCQTChiCucPath);
						string value11 = ((cellMCTLanTruoc == null) ? "" : cellMCTLanTruoc.Value);
						string text3 = ((hTKKCMCCellInfo12 == null) ? "" : hTKKCMCCellInfo12.Value);
						string text4 = ((hTKKCMCCellInfo13 == null) ? "" : hTKKCMCCellInfo13.Value);
						string text5 = ((hTKKCMCCellInfo14 == null) ? "" : hTKKCMCCellInfo14.Value);
						string text6 = ((hTKKCMCCellInfo15 == null) ? "" : hTKKCMCCellInfo15.Value);
						string text7 = ((hTKKCMCCellInfo16 == null) ? "" : hTKKCMCCellInfo16.Value);
						string text8 = ((hTKKCMCCellInfo17 == null) ? "" : hTKKCMCCellInfo17.Value);
						string text9 = ((hTKKCMCCellInfo18 == null) ? "" : hTKKCMCCellInfo18.Value);
						string text10 = ((hTKKCMCCellInfo19 == null) ? "" : hTKKCMCCellInfo19.Value);
						string text11 = ((hTKKCMCCellInfo20 == null) ? "" : hTKKCMCCellInfo20.Value);
						string text12 = ((hTKKCMCCellInfo21 == null) ? "" : hTKKCMCCellInfo21.Value);
						string tenTieuMucTuMa2;
						double num32;
						if (string.IsNullOrEmpty(text3))
						{
							if (item2.Header.StaticCells.Count <= 0)
							{
							}
						}
						else
						{
							if (string.IsNullOrEmpty(text5.Trim()) || string.IsNullOrEmpty(text7.Trim()) || string.IsNullOrEmpty(text9.Trim()) || string.IsNullOrEmpty(text11.Trim()))
							{
								continue;
							}
							tenTieuMucTuMa2 = GetTenTieuMucTuMa(text3);
							num32 = 0.0;
							if (!string.IsNullOrEmpty(value11))
							{
								num32 = Convert.ToDouble(value11);
							}
							List<int> list3 = new List<int>();
							list3.Add(num4);
							if (!string.IsNullOrEmpty(text4))
							{
								list3.Add(num6);
							}
							list3.Add(num9);
							list3.Add(num11);
							list3.Add(num13);
							list3.Add(num15);
							List<string> list4 = new List<string>();
							list4.Add(maTCTPath);
							if (!string.IsNullOrEmpty(text4))
							{
								list4.Add(MSTPath);
							}
							list4.Add(maTinhPath);
							list4.Add(maHuyenPath);
							list4.Add(maCQTCapCucPath);
							list4.Add(maCQTChiCucPath);
							List<string> list5 = new List<string>();
							list5.Add(text3);
							if (!string.IsNullOrEmpty(text4))
							{
								list5.Add(text4);
							}
							list5.Add(text6);
							list5.Add(text8);
							list5.Add(text10);
							list5.Add(text12);
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list3, list5);
							if (rowIndexInSViewByCompares.Count <= 0)
							{
								double num33 = 0.0;
								if (sheetView2 == null)
								{
									goto IL_18f6;
								}
								HTKKCMCTableInfo hTKKCMCTableInfo2 = DetectIndexTableHeaderBySview(sheetView2, tables, list4, list5, item2.TablePath);
								if (hTKKCMCTableInfo2 == null)
								{
									goto IL_18f6;
								}
								HTKKCMCCellInfo hTKKCMCCellInfo22 = hTKKCMCTableInfo2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == cellMCTLanTruoc.Path);
								num33 = Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo22.CellTemplate.RowIndex - 1, hTKKCMCCellInfo22.CellTemplate.ColIndex - 1].Value);
								if (num32 != num33)
								{
									goto IL_18f6;
								}
							}
							continue;
						}
						goto end_IL_133d;
						IL_18f6:
						if (num19 > 0)
						{
							WinSpreadProcess.InsertTableRow(sheetView, dataLocationTemplateRowStart + num19 - 1, -1, dataLocationTemplateRowStart + num19 - 1, -1, -1, sInfo, toTableIndex);
						}
						ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + num19, colSheetIndex, num18, colSheetName, sheetname, colLoaiThue, valueLoaiThue);
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text3, tenTieuMucTuMa2, " - ");
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num5 - 1].Value = tenTieuMucTuMa2;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num4 - 1].Value = text3;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num6 - 1].Value = text4;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num7 - 1].Value = text4;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num8 - 1].Value = text5;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num9 - 1].Value = text6;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num10 - 1].Value = text7;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num11 - 1].Value = text8;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num12 - 1].Value = text9;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num13 - 1].Value = text10;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num14 - 1].Value = text11;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num15 - 1].Value = text12;
						sheetView.Cells[dataLocationTemplateRowStart + num19 - 1, num16 - 1].Value = -1.0 * num32;
						num19++;
						end_IL_133d:;
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
			}
			GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void CollectionDetailSoThuePhaiNopTungTinh_Static_01_KHBS(XmlNode mapCt, int sheetIndexToDieuChinh, int toTableIndex)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo sInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("E");
			int num4 = HTKKFile.ColumnLettersToIndex("BA");
			int num5 = HTKKFile.ColumnLettersToIndex("BC");
			int num6 = HTKKFile.ColumnLettersToIndex("M");
			int num7 = HTKKFile.ColumnLettersToIndex("BB");
			int num8 = HTKKFile.ColumnLettersToIndex("P");
			int num9 = HTKKFile.ColumnLettersToIndex("AO");
			int num10 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "V" : "T");
			int num11 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "BE" : "AP");
			int num12 = HTKKFile.ColumnLettersToIndex("X");
			int num13 = HTKKFile.ColumnLettersToIndex("AQ");
			int num14 = HTKKFile.ColumnLettersToIndex("AB");
			int num15 = HTKKFile.ColumnLettersToIndex("AR");
			int num16 = HTKKFile.ColumnLettersToIndex("AF");
			int num17 = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("AU");
			int colSheetName = HTKKFile.ColumnLettersToIndex("AV");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("AW");
			int num18 = HTKKFile.ColumnLettersToIndex("BD");
			int num19 = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			string sheetname = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string text = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string TCTPath = ((mapCt.Attributes["TCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim());
			string maTCTPath = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string valueLoaiThue = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string text2 = ((_mapMct != null && _mapMct.Attributes["SuaPhanBo"] != null) ? ((XmlNode)_mapMct.Attributes["SuaPhanBo"]).Value : "");
			string prefixPath = ((mapCt.Attributes["PrefixPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["PrefixPath"]).Value.Trim());
			string text3 = ((mapCt.Attributes["PhanBoTheoTSC"] == null) ? "" : ((XmlNode)mapCt.Attributes["PhanBoTheoTSC"]).Value.ToString().Trim());
			if (num19 != sheetIndexToDieuChinh || string.IsNullOrEmpty(text3) || text3 != "1")
			{
				return;
			}
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinFnProcess.ReadSheet_LanTruoc(num19);
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num].Tables[toTableIndex];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			int num20 = hTKKCMCTableInfo.RowsValues.Count - 1;
			List<int> list = new List<int>();
			list.Add(num4);
			list.Add(num6);
			list.Add(num9);
			list.Add(num11);
			list.Add(num13);
			list.Add(num15);
			if (!string.IsNullOrEmpty(prefixPath))
			{
				list.Add(num18);
			}
			List<int> list2 = new List<int>();
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = WinSettings.ToKhaiInfo.SheetInfos[num19];
			SheetView sheetView2 = ((hTKKCMCSheetInfo2 == null) ? null : winFpSpread.Sheets[num19]);
			if (hTKKCMCSheetInfo2 != null && sheetView2 != null)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo2.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo2.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == TCTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo2.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTCTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo2.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == prefixPath);
				if (hTKKCMCCellInfo3 != null)
				{
					ICellType cellType = sheetView.Cells[0, 0].CellType;
					if (list2.Count == 0)
					{
						list2.Add(hTKKCMCCellInfo3.CellTemplate.ColIndex);
						if (hTKKCMCCellInfo4 != null)
						{
							list2.Add(hTKKCMCCellInfo4.CellTemplate.ColIndex);
						}
					}
					string text4 = ((hTKKCMCCellInfo3 != null) ? Convert.ToString(sheetView2.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value) : "");
					string tenTieuMucTuMa = GetTenTieuMucTuMa(text4);
					if (tenTieuMucTuMa != null && !string.IsNullOrEmpty(tenTieuMucTuMa.ToString()))
					{
						double num21 = Convert.ToDouble(sheetView2.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
						string mst = WinSettings.HTKKHeaderInfo.mst;
						string tenTinhNNT = WinSettings.HTKKHeaderInfo.tenTinhNNT;
						string text5 = ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? WinSettings.HTKKHeaderInfo.tenXaPhuongNNT : WinSettings.HTKKHeaderInfo.tenHuyenNNT);
						string tenCQTCapCuc = WinSettings.HTKKHeaderInfo.tenCQTCapCuc;
						string tenCQTNoiNop = WinSettings.HTKKHeaderInfo.tenCQTNoiNop;
						string maTinhNNT = WinSettings.HTKKHeaderInfo.maTinhNNT;
						string text6 = ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? WinSettings.HTKKHeaderInfo.maXaPhuongNNT : WinSettings.HTKKHeaderInfo.maHuyenNNT);
						string maCQTCapCuc = WinSettings.HTKKHeaderInfo.maCQTCapCuc;
						string maCQTNoiNop = WinSettings.HTKKHeaderInfo.maCQTNoiNop;
						string text7 = "";
						if (hTKKCMCCellInfo4 != null)
						{
							text7 = hTKKCMCCellInfo4.Value;
						}
						if (!string.IsNullOrEmpty(num21.ToString()))
						{
							List<string> list3 = new List<string>();
							list3.Add((text4 == null) ? "" : text4.ToString());
							list3.Add((mst == null) ? "" : mst.ToString());
							list3.Add((maTinhNNT == null) ? "" : maTinhNNT.ToString());
							list3.Add((text6 == null) ? "" : text6.ToString());
							list3.Add((maCQTCapCuc == null) ? "" : maCQTCapCuc.ToString());
							list3.Add((maCQTNoiNop == null) ? "" : maCQTNoiNop.ToString());
							if (hTKKCMCCellInfo4 != null)
							{
								list3.Add(text7);
							}
							double num22 = 0.0;
							if (hTKKCMCSheetInfo != null)
							{
								HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
								if (hTKKCMCCellInfo5 != null)
								{
									num22 = Convert.ToDouble(hTKKCMCCellInfo5.Value);
								}
							}
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list3);
							if (rowIndexInSViewByCompares.Count > 0)
							{
								object value = sheetView.Cells[rowIndexInSViewByCompares[0] - 1, num16 - 1].Value;
								sheetView.Cells[rowIndexInSViewByCompares[0] - 1, num16 - 1].Value = ((value != null && !string.IsNullOrEmpty(value.ToString().Trim())) ? (Convert.ToDouble(value) + num21) : num21);
							}
							else if (num22 != num21)
							{
								if (num20 > 0)
								{
									WinSpreadProcess.InsertTableRow(sheetView, dataLocationTemplateRowStart + num20 - 1, -1, dataLocationTemplateRowStart + num20 - 1, -1, -1, sInfo, toTableIndex);
								}
								ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + num20, colSheetIndex, num19, colSheetName, sheetname, colLoaiThue, valueLoaiThue);
								if (string.IsNullOrEmpty(text7))
								{
									sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text4, tenTieuMucTuMa.ToString(), " - ");
								}
								else
								{
									List<string> list4 = new List<string>();
									list4.Add(text7);
									list4.Add(text4);
									list4.Add(tenTieuMucTuMa.ToString());
									string value2 = GeneralFunctions.ListToString(list4, " - ");
									sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = value2;
								}
								string value3 = ((!string.IsNullOrEmpty(text7)) ? GeneralFunctions.FormatKey2(text7, tenTieuMucTuMa, " - ") : tenTieuMucTuMa);
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num4 - 1].Value = text4;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num5 - 1].Value = value3;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num6 - 1].Value = mst;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num7 - 1].Value = mst;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num8 - 1].Value = ((tenTinhNNT != null) ? tenTinhNNT : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num9 - 1].Value = ((maTinhNNT != null) ? maTinhNNT : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num10 - 1].Value = ((text5 != null) ? text5 : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num11 - 1].Value = ((text6 != null) ? text6 : "");
								tenCQTCapCuc = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(maCQTCapCuc, isGetFullName: false);
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num12 - 1].Value = ((tenCQTCapCuc != null) ? tenCQTCapCuc : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num13 - 1].Value = ((maCQTCapCuc != null) ? maCQTCapCuc : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num14 - 1].Value = ((tenCQTNoiNop != null) ? tenCQTNoiNop : "");
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num15 - 1].Value = ((maCQTNoiNop != null) ? maCQTNoiNop : "");
								object value4 = sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num16 - 1].Value;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num16 - 1].Value = num21 - num22;
								sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num18 - 1].Value = text7;
								num20++;
							}
						}
					}
				}
			}
			else if (hTKKCMCSheetInfo != null)
			{
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo7 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == maTCTPath);
				HTKKCMCCellInfo hTKKCMCCellInfo8 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == prefixPath);
				double num22 = 0.0;
				string value5 = hTKKCMCCellInfo6.Value;
				string value6 = hTKKCMCCellInfo7.Value;
				string mst2 = WinSettings.HTKKHeaderInfo.mst;
				string tenTinhNNT2 = WinSettings.HTKKHeaderInfo.tenTinhNNT;
				string maTinhNNT2 = WinSettings.HTKKHeaderInfo.maTinhNNT;
				string value7 = ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? WinSettings.HTKKHeaderInfo.tenXaPhuongNNT : WinSettings.HTKKHeaderInfo.tenHuyenNNT);
				string text8 = ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? WinSettings.HTKKHeaderInfo.maXaPhuongNNT : WinSettings.HTKKHeaderInfo.maHuyenNNT);
				string tenCQTCapCuc2 = WinSettings.HTKKHeaderInfo.tenCQTCapCuc;
				string maCQTCapCuc2 = WinSettings.HTKKHeaderInfo.maCQTCapCuc;
				string tenCQTNoiNop2 = WinSettings.HTKKHeaderInfo.tenCQTNoiNop;
				string maCQTNoiNop2 = WinSettings.HTKKHeaderInfo.maCQTNoiNop;
				string text9 = "";
				if (!string.IsNullOrEmpty(prefixPath))
				{
					text9 = hTKKCMCCellInfo8.Value;
				}
				string tenTieuMucTuMa2 = GetTenTieuMucTuMa(value6);
				if (!string.IsNullOrEmpty(tenTieuMucTuMa2))
				{
					List<string> list3 = new List<string>();
					list3.Add(value6);
					list3.Add(mst2);
					list3.Add(maTinhNNT2);
					list3.Add(text8);
					list3.Add(maCQTCapCuc2);
					list3.Add(maCQTNoiNop2);
					if (!string.IsNullOrEmpty(prefixPath))
					{
						list3.Add(text9);
					}
					value5 = FormatValueLanTruocStatic(hTKKCMCCellInfo6, value5);
					if (!string.IsNullOrEmpty(value5))
					{
						num22 = Convert.ToDouble(value5);
					}
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo, list, list3);
					if (rowIndexInSViewByCompares.Count <= 0 && num22 != 0.0)
					{
						if (num20 > 0)
						{
							WinSpreadProcess.InsertTableRow(sheetView, dataLocationTemplateRowStart + num20 - 1, -1, dataLocationTemplateRowStart + num20 - 1, -1, -1, sInfo, toTableIndex);
						}
						ThongTinThemChoKHBSMoi(sheetView, dataLocationTemplateRowStart + num20, colSheetIndex, num19, colSheetName, sheetname, colLoaiThue, valueLoaiThue);
						if (string.IsNullOrEmpty(text9))
						{
							sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(value6, tenTieuMucTuMa2, " - ");
						}
						else
						{
							List<string> list4 = new List<string>();
							list4.Add(text9);
							list4.Add(value6);
							list4.Add(tenTieuMucTuMa2);
							string value2 = GeneralFunctions.ListToString(list4, " - ");
							sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num3 - 1].Value = value2;
						}
						string value3 = ((!string.IsNullOrEmpty(text9)) ? GeneralFunctions.FormatKey2(text9, tenTieuMucTuMa2, " - ") : tenTieuMucTuMa2);
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num5 - 1].Value = value3;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num4 - 1].Value = value6;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num6 - 1].Value = mst2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num7 - 1].Value = mst2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num8 - 1].Value = tenTinhNNT2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num9 - 1].Value = maTinhNNT2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num10 - 1].Value = value7;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num11 - 1].Value = text8;
						tenCQTCapCuc2 = HTKKWinDictionaryDiaDanh.GetTenCoQuanThueCapCuc(maCQTCapCuc2, isGetFullName: false);
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num12 - 1].Value = tenCQTCapCuc2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num13 - 1].Value = maCQTCapCuc2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num14 - 1].Value = tenCQTNoiNop2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num15 - 1].Value = maCQTNoiNop2;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num16 - 1].Value = -1.0 * num22;
						sheetView.Cells[dataLocationTemplateRowStart + num20 - 1, num18 - 1].Value = text9;
						num20++;
					}
				}
			}
			GeneralFunctions.fpSpReindexingSheetTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public string GetTenTieuMucTuMa(string maTieuMuc)
	{
		try
		{
			if (string.IsNullOrEmpty(maTieuMuc))
			{
				return "";
			}
			DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
			return GeneralFunctions.GetTenTmByMaTm(maTieuMuc, ngayDauKyKeKhai);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return "";
		}
	}

	private void ChuyenTu_01_1_KHBS_Sang_01_KHBS(int toTableIndex, string loaiThue)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			int num2 = HTKKFile.ColumnLettersToIndex("BR");
			int num3 = HTKKFile.ColumnLettersToIndex("BV");
			int num4 = HTKKFile.ColumnLettersToIndex("AH");
			int num5 = HTKKFile.ColumnLettersToIndex("CA");
			int num6 = HTKKFile.ColumnLettersToIndex("AD");
			int num7 = HTKKFile.ColumnLettersToIndex("BS");
			int num8 = HTKKFile.ColumnLettersToIndex("BX");
			int num9 = HTKKFile.ColumnLettersToIndex("BQ");
			int num10 = HTKKFile.ColumnLettersToIndex("C");
			int num11 = HTKKFile.ColumnLettersToIndex("AT");
			int num12 = HTKKFile.ColumnLettersToIndex("O");
			int num13 = HTKKFile.ColumnLettersToIndex("AU");
			int num14 = HTKKFile.ColumnLettersToIndex("U");
			int num15 = HTKKFile.ColumnLettersToIndex("AV");
			int num16 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num16].Tables[toTableIndex];
			SheetView sheetView2 = winFpSpread.Sheets[num16];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			if (hTKKCMCTableInfo.RowsValues.Count > 1)
			{
				WinSpreadProcess.RemoveDataRows(winFpSpread.Sheets[num16], toTableIndex, hTKKCMCTableInfo.Location.DataLocationTemplateRowStart - 1, hTKKCMCTableInfo.RowsValues.Count - 1, WinSettings.ToKhaiInfo.SheetInfos[num16]);
			}
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < hTKKCMCTableInfo.RowsValues.Count; i++)
			{
				foreach (HTKKCMCTableCellInfo cell in hTKKCMCTableInfo.RowsInfo.Cells)
				{
					sheetView2.Cells[dataLocationTemplateRowStart2 + i - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
			int num17 = ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID) ? (WinSettings.ToKhaiInfo.SheetInfos[num].Tables.Count - 3) : WinSettings.ToKhaiInfo.SheetInfos[num].Tables.Count);
			List<string> list = new List<string>();
			int num18 = 0;
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			for (int j = 0; j < num17; j++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo2 = WinSettings.ToKhaiInfo.SheetInfos[num].Tables[j];
				int dataLocationTemplateRowStart3 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				for (int k = 0; k < hTKKCMCTableInfo2.RowsValues.Count; k++)
				{
					try
					{
						string text = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num2 - 1].Text.Trim();
						string text2 = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num3 - 1].Text.Trim();
						string value = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num4 - 1].Text.Trim();
						string text3 = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num5 - 1].Text.Trim();
						ICellType cellType = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num6 - 1].CellType;
						int digits = 0;
						if ((object)cellType.GetType() == numberCellType.GetType())
						{
							digits = ((NumberCellType)cellType).DecimalPlaces;
						}
						if ((object)cellType.GetType() == currencyCellType.GetType())
						{
							digits = ((CurrencyCellType)cellType).DecimalPlaces;
						}
						if (!string.IsNullOrEmpty(text2) && string.IsNullOrEmpty(text))
						{
							string item = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num7 - 1].Text.Trim();
							string item2 = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num9 - 1].Text.Trim();
							string item3 = sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num8 - 1].Text.Trim();
							List<int> list2 = new List<int>();
							list2.Add(num7);
							list2.Add(num9);
							list2.Add(num8);
							List<string> list3 = new List<string>();
							list3.Add(item);
							list3.Add(item2);
							list3.Add(item3);
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView, hTKKCMCTableInfo2, list2, list3);
							foreach (int item4 in rowIndexInSViewByCompares)
							{
								string text4 = sheetView.Cells[item4 - 1, num2 - 1].Text.Trim();
								if (!string.IsNullOrEmpty(text4))
								{
									text = text4;
									sheetView.Cells[dataLocationTemplateRowStart3 + k - 1, num2 - 1].Text = text;
									break;
								}
							}
						}
						string text5 = text + text3;
						if (string.IsNullOrEmpty(text5) || string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text2) || !(text2 == loaiThue))
						{
							continue;
						}
						double num19 = TinhSoThuePhaiNopCho01_1_Khbs(text, loaiThue, text3);
						if (Math.Round(num19, digits) == 0.0)
						{
							continue;
						}
						if (!list.Contains(text5))
						{
							if (num18 > 0)
							{
								GeneralFunctions.InsertTableRow(sheetView2, dataLocationTemplateRowStart + num18 - 1, -1, dataLocationTemplateRowStart + num18 - 1, -1, -1, WinSettings.ToKhaiInfo.SheetInfos[num16], toTableIndex);
							}
							DateTime ngayDauKyKeKhai = GeneralFunctions.GetNgayDauKyKeKhai(WinSettings.TTinChung, WinSettings.TTPeriodInfo);
							string tenTmByMaTm = GeneralFunctions.GetTenTmByMaTm(text, ngayDauKyKeKhai);
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num10 - 1].Value = GeneralFunctions.FormatKey2(text, tenTmByMaTm, " - ");
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num11 - 1].Value = text;
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num12 - 1].Value = num19;
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num13 - 1].Value = tenTmByMaTm;
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num14 - 1].Value = value;
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num15 - 1].Value = text3;
							list.Add(text5);
							sheetView2.Cells[dataLocationTemplateRowStart + num18 - 1, num12 - 1].CellType = cellType;
							GeneralFunctions.SetRowHeight(sheetView2, dataLocationTemplateRowStart + num18, 30f);
							num18++;
						}
						else
						{
							List<int> list2 = new List<int>();
							list2.Add(num11);
							list2.Add(num15);
							List<string> list3 = new List<string>();
							list3.Add(text);
							list3.Add(text3);
							double num20 = TinhSoThuePhaiNopCho01_1_Khbs(text, loaiThue, text3);
							List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares(sheetView2, hTKKCMCTableInfo, list2, list3);
							if (rowIndexInSViewByCompares.Count > 0)
							{
								int num21 = rowIndexInSViewByCompares[0];
								sheetView2.Cells[rowIndexInSViewByCompares[0] - 1, num12 - 1].Value = num20;
							}
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
			}
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView2, WinSettings.ToKhaiInfo.SheetInfos[num16].Tables[toTableIndex]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private double TinhSoThuePhaiNopCho01_1_Khbs(string maTieuMuc, string loaiThue, string maDonViTien)
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("AD");
			int num2 = HTKKFile.ColumnLettersToIndex("BR");
			int num3 = HTKKFile.ColumnLettersToIndex("BV");
			int num4 = HTKKFile.ColumnLettersToIndex("CA");
			double num5 = 0.0;
			int num6 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			for (int i = 0; i < WinSettings.ToKhaiInfo.SheetInfos[num6].Tables.Count; i++)
			{
				HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num6].Tables[i];
				SheetView sheetView = winFpSpread.Sheets[num6];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int j = 0; j < hTKKCMCTableInfo.RowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Text.Trim();
					double num7 = Convert.ToDouble(sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Value);
					string text2 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num2 - 1].Text.Trim();
					string text3 = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num4 - 1].Text.Trim();
					if (string.IsNullOrEmpty(maDonViTien))
					{
						if (text2 == maTieuMuc && text == loaiThue)
						{
							num5 += num7;
						}
					}
					else if (text2 == maTieuMuc && text == loaiThue && text3 == maDonViTien)
					{
						num5 += num7;
					}
				}
			}
			return num5;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return -1.0;
		}
	}

	private void ThongTinThemChoKHBSMoi(SheetView sView, int row, int colSheetIndex, int sheetIndex, int colSheetName, string sheetname, int colLoaiThue, string valueLoaiThue)
	{
		try
		{
			sView.Cells[row - 1, colSheetIndex - 1].Value = sheetIndex;
			sView.Cells[row - 1, colSheetName - 1].Value = sheetname;
			sView.Cells[row - 1, colLoaiThue - 1].Value = valueLoaiThue;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private bool checkExitsTrongsView(SheetView sView, HTKKCMCTableInfo table, int colCheck, string valueCompare, int colCheck2, string valueCompare2)
	{
		try
		{
			int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
			for (int i = 0; i < table.RowsValues.Count; i++)
			{
				if (sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck - 1].Value != null && sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck - 1].Value.ToString().Trim() == valueCompare.Trim() && sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck2 - 1].Value != null && sView.Cells[dataLocationTemplateRowStart + i - 1, colCheck2 - 1].Value.ToString().Trim() == valueCompare2.Trim())
				{
					return true;
				}
			}
			return false;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void CellChange_Event_01_1_KHBS(int iCol, int iRow)
	{
		try
		{
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void NeuClickNhapText_01_1_KHBS(SheetView sView, int iRow, int colNhapText_KHBS, int colCol2_KHBS_Backup, int colMa2_KHBS_Backup, int col2_KHBS, int colMa2_KHBS)
	{
		try
		{
			if (!Convert.ToBoolean(winFpSpread.ActiveSheet.Cells[iRow - 1, colNhapText_KHBS - 1].Value))
			{
				winFpSpread.ActiveSheet.Cells[iRow - 1, colCol2_KHBS_Backup - 1].Value = sView.Cells[iRow, col2_KHBS - 1].Value;
				winFpSpread.ActiveSheet.Cells[iRow - 1, colCol2_KHBS_Backup - 1].CellType = sView.Cells[iRow, col2_KHBS - 1].CellType;
				winFpSpread.ActiveSheet.Cells[iRow - 1, colMa2_KHBS_Backup - 1].Value = sView.Cells[iRow, colMa2_KHBS - 1].Value;
				winFpSpread.ActiveSheet.Cells[iRow - 1, col2_KHBS - 1].Value = "";
				winFpSpread.ActiveSheet.Cells[iRow - 1, col2_KHBS - 1].CellType = sView.Cells[0, 0].CellType;
				winFpSpread.ActiveSheet.Cells[iRow - 1, colMa2_KHBS - 1].Value = "";
			}
			else
			{
				winFpSpread.ActiveSheet.Cells[iRow - 1, col2_KHBS - 1].Value = sView.Cells[iRow - 1, colCol2_KHBS_Backup - 1].Value;
				winFpSpread.ActiveSheet.Cells[iRow - 1, col2_KHBS - 1].CellType = sView.Cells[iRow - 1, colCol2_KHBS_Backup - 1].CellType;
				winFpSpread.ActiveSheet.Cells[iRow - 1, colMa2_KHBS - 1].Value = sView.Cells[iRow - 1, colMa2_KHBS_Backup - 1].Value;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public XmlNode ChonChiTieu03Truoc_01_1_KHBS(SheetView sView, int iRow, int col3ID_KHBS, int col2_KHBS, int cellDynamic, int colSheetIndex)
	{
		int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
		SheetView sheetView = winFpSpread.Sheets[num];
		XmlNode val = null;
		try
		{
			string MCT = "";
			string tablePath = "";
			string text = sView.Cells[iRow - 1, col3ID_KHBS - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
				if (mCT != null && text.Contains("___"))
				{
					string[] array = Functions.DeserializeKey(text);
					tablePath = array[0];
					MCT = array[1].Replace("_", "");
					if (mCT != null)
					{
						val = ((IEnumerable)mCT.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["TablePath"]).Value == tablePath && ((XmlNode)x.Attributes["MCT"]).Value == MCT));
						int num2 = Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value);
						string TCTPath = ((XmlNode)val.Attributes["TCTPath"]).Value;
						int num3 = Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value);
						SheetView sheetView2 = winFpSpread.Sheets[num2];
						HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num2].Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath);
						if (hTKKCMCTableInfo != null)
						{
							int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
							HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == TCTPath);
							ICellType cellType = sheetView2.Cells[dataLocationTemplateRowStart - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].CellType;
							TextCellType textCellType = new TextCellType();
							if (cellType.GetType().FullName == textCellType.GetType().FullName)
							{
								Dictionary<string, string> dictionary = new Dictionary<string, string>();
								for (int num4 = 0; num4 < hTKKCMCTableInfo.RowsValues.Count; num4++)
								{
									string text2 = sheetView2.Cells[dataLocationTemplateRowStart + num4 - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
									string key = Functions.RemoveUnikey(text2).Replace(" ", "");
									if (!dictionary.ContainsKey(key))
									{
										dictionary.Add(key, text2);
									}
								}
								Functions.InsertCombobox(sheetView, iRow, col2_KHBS, dictionary, deleteText: false);
							}
							else
							{
								sView.Cells[iRow - 1, col2_KHBS - 1].CellType = cellType;
							}
						}
						sheetView.Cells[iRow - 1, cellDynamic - 1].Value = num3;
						sheetView.Cells[iRow - 1, colSheetIndex - 1].Value = num2;
					}
				}
				else
				{
					MCT = text;
					val = ((IEnumerable)mCT.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["Dynamic"]).Value == "0" && ((XmlNode)x.Attributes["MCT"]).Value == MCT));
					sView.Cells[iRow - 1, col2_KHBS - 1].Text = ((XmlNode)val.Attributes["Caption"]).Value;
					sView.Cells[iRow - 1, col2_KHBS - 1].CellType = sView.Cells[0, 0].CellType;
					int num3 = Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value);
					int num2 = Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value);
					sheetView.Cells[iRow - 1, cellDynamic - 1].Value = num3;
					sheetView.Cells[iRow - 1, colSheetIndex - 1].Value = num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return val;
	}

	public XmlNode ChonChiTieu02Truoc_01_1_KHBS(SheetView sView, int iRow, int colMa2_KHBS, int col2_KHBS, int col3_KHBS, int col3Ma_KHBS, int col2ID_KHBS, int cellDynamic, int colSheetIndex)
	{
		int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
		SheetView sheetView = winFpSpread.Sheets[num];
		XmlNode val = null;
		try
		{
			string MCT = "";
			string tablePath = "";
			string text = sView.Cells[iRow - 1, col2ID_KHBS - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text))
			{
				XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
				string value = ((XmlNode)mCT.Attributes["ChonMCTTruoc"]).Value;
				if (text.Contains("___"))
				{
					string[] array = Functions.DeserializeKey(text);
					if (array.Count() != 3)
					{
						return val;
					}
					tablePath = array[2];
					MCT = array[1].Replace("_", "");
					sView.Cells[iRow - 1, col3Ma_KHBS - 1].Text = MCT;
					sView.Cells[iRow - 1, colMa2_KHBS - 1].Text = array[0];
					if (mCT != null)
					{
						val = ((IEnumerable)mCT.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["TablePath"]).Value == tablePath && ((XmlNode)x.Attributes["MCT"]).Value == MCT));
						int num2 = Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value);
						string value2 = ((XmlNode)val.Attributes["TCTPath"]).Value;
						int num3 = Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value);
						sView.Cells[iRow - 1, col3_KHBS - 1].Text = ((XmlNode)val.Attributes["Caption"]).Value.Trim();
						sheetView.Cells[iRow - 1, cellDynamic - 1].Value = num3;
						sheetView.Cells[iRow - 1, colSheetIndex - 1].Value = num2;
					}
				}
				else
				{
					MCT = text;
					val = ((IEnumerable)mCT.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["Dynamic"]).Value == "0" && ((XmlNode)x.Attributes["MCT"]).Value == MCT));
					int num3 = Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value);
					int num2 = Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value);
					sView.Cells[iRow - 1, col3_KHBS - 1].Text = ((XmlNode)val.Attributes["Caption"]).Value.Trim();
					sView.Cells[iRow - 1, col3Ma_KHBS - 1].Text = MCT;
					sheetView.Cells[iRow - 1, cellDynamic - 1].Value = num3;
					sheetView.Cells[iRow - 1, colSheetIndex - 1].Value = num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return val;
	}

	public string GetMaTieuMuc(int SheetIndex, string MCTPath)
	{
		try
		{
			if (_mapMctTieuMuc != null)
			{
				XmlNode val = ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == SheetIndex.ToString()));
				return (val != null) ? ((XmlNode)val.Attributes["MaTieuMuc"]).Value.ToString() : "";
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	public XmlNode GetItemMapMct(XmlNode mapMct, string valueCol3ID)
	{
		int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
		SheetView sheetView = winFpSpread.Sheets[num];
		XmlNode val = null;
		try
		{
			if (!string.IsNullOrEmpty(valueCol3ID))
			{
				if (mapMct != null && valueCol3ID.Contains("___"))
				{
					string[] array = Functions.DeserializeKey(valueCol3ID);
					string tablePath = array[0];
					string MCTPath = array[1];
					if (mapMct != null)
					{
						val = ((IEnumerable)mapMct.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath));
					}
				}
				else
				{
					val = ((IEnumerable)mapMct.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => ((XmlNode)x.Attributes["Dynamic"]).Value == "0" && ((XmlNode)x.Attributes["MCTPath"]).Value == valueCol3ID));
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
		return val;
	}

	public void UpdateTieuMucFor_01_KHBS()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num].Tables;
			int num2 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/ChiTiet");
			int num3 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_II/ChiTiet");
			int num4 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_III/ChiTiet");
			SheetView sView = winFpSpread.Sheets[num];
			int col = HTKKFile.ColumnLettersToIndex("C");
			HTKKCMCTableInfo hTKKCMCTableInfo = tables[num2];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num5 = 0; num5 < hTKKCMCTableInfo.RowsValues.Count; num5++)
			{
				GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart + num5, col, dicTieuMucKHBS);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo2 = tables[num3];
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int num5 = 0; num5 < hTKKCMCTableInfo2.RowsValues.Count; num5++)
			{
				GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart2 + num5, col, dicTieuMucKHBS);
			}
			HTKKCMCTableInfo hTKKCMCTableInfo3 = tables[num4];
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			for (int num5 = 0; num5 < hTKKCMCTableInfo3.RowsValues.Count; num5++)
			{
				GeneralFunctions.InsertCombobox(sView, dataLocationTemplateRowStart3 + num5, col, dicTieuMucKHBS);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public Dictionary<string, string> GetDMTenChiTieuKHBS()
	{
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Expected O, but got Unknown
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
			if (mCT == null)
			{
				return dictionary;
			}
			foreach (XmlNode childNode in mCT.ChildNodes)
			{
				XmlNode val = childNode;
				try
				{
					int num3 = ((val.Attributes["Dynamic"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value));
					int num4 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num : num2));
					if (num3 < 0 || num4 < 0)
					{
						break;
					}
					if (num != num4)
					{
						continue;
					}
					string value;
					string str2;
					if (num3 == 0)
					{
						value = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value.ToString().Trim());
						string str = ((val.Attributes["MCTPath"] == null) ? "" : ((XmlNode)val.Attributes["MCTPath"]).Value.ToString().Trim());
						str2 = ((val.Attributes["MCT"] == null) ? "" : ((XmlNode)val.Attributes["MCT"]).Value.ToString().Trim());
						string key = Functions.FomartKey(str2, str);
						if (!dictionary.ContainsKey(key))
						{
							dictionary.Add(key, value);
						}
					}
					if (num3 != 1)
					{
						continue;
					}
					value = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value.ToString().Trim());
					string tablePath = ((val.Attributes["TablePath"] == null) ? "" : ((XmlNode)val.Attributes["TablePath"]).Value.ToString().Trim());
					str2 = ((val.Attributes["MCT"] == null) ? "" : ((XmlNode)val.Attributes["MCT"]).Value.ToString().Trim());
					int num5 = ((val.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value));
					string TCTPath = ((val.Attributes["TCTPath"] == null) ? "" : ((XmlNode)val.Attributes["TCTPath"]).Value);
					if (num5 < 0 || string.IsNullOrEmpty(TCTPath))
					{
						continue;
					}
					SheetView sheetView = winFpSpread.Sheets[num5];
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[num5].Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath);
					if (hTKKCMCTableInfo == null)
					{
						continue;
					}
					int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == TCTPath);
					ICellType cellType = sheetView.GetCellType(dataLocationTemplateRowStart - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1);
					ComboBoxCellType comboBoxCellType = new ComboBoxCellType();
					if (!(cellType.GetType().FullName == comboBoxCellType.GetType().FullName))
					{
						continue;
					}
					string[] items = ((ComboBoxCellType)cellType).Items;
					string[] itemData = ((ComboBoxCellType)cellType).ItemData;
					if (itemData.Count() == 0)
					{
						listTenChiTieuKHBS = items;
						continue;
					}
					for (int num6 = 0; num6 < items.Count(); num6++)
					{
						try
						{
							string key = Functions.FomartKey(itemData[num6], str2);
							key = Functions.FomartKey(key, tablePath);
							if (!dictionary.ContainsKey(key) && !string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(items[num6]))
							{
								dictionary.Add(key, items[num6]);
							}
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
					}
					continue;
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					continue;
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public Dictionary<string, string> GetDMMaChiTieuTheoPhuLucKHBS(int sheetIndex)
	{
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Expected O, but got Unknown
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			XmlNode mCT = WinFnProcess.GetMCT01("MapMCT.xml", WinSettings.ToKhaiInfo.MenuID);
			if (mCT == null)
			{
				return dictionary;
			}
			foreach (XmlNode childNode in mCT.ChildNodes)
			{
				XmlNode val = childNode;
				int num3 = ((val.Attributes["Dynamic"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["Dynamic"]).Value));
				int num4 = ((val.Attributes["ToSheetIndex"] == null) ? (-1) : ((((XmlNode)val.Attributes["ToSheetIndex"]).Value.Trim() == "01-1") ? num : num2));
				int num5 = ((val.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val.Attributes["SheetIndex"]).Value));
				string text = ((val.Attributes["Manual"] == null) ? "" : ((XmlNode)val.Attributes["Manual"]).Value.Trim());
				if (num != num4 || !(text == "1") || num5 != sheetIndex)
				{
					continue;
				}
				if (num3 == 0)
				{
					string value = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value.ToString().Trim());
					string text2 = ((val.Attributes["MCTPath"] == null) ? "" : ((XmlNode)val.Attributes["MCTPath"]).Value.ToString().Trim());
					if (!dictionary.ContainsKey(text2) && !string.IsNullOrEmpty(text2) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(text2, value);
					}
				}
				if (num3 == 1)
				{
					string value = ((val.Attributes["Caption"] == null) ? "" : ((XmlNode)val.Attributes["Caption"]).Value.ToString().Trim());
					string str = ((val.Attributes["TablePath"] == null) ? "" : ((XmlNode)val.Attributes["TablePath"]).Value.ToString().Trim());
					string text2 = ((val.Attributes["MCTPath"] == null) ? "" : ((XmlNode)val.Attributes["MCTPath"]).Value.ToString().Trim());
					string text3 = Functions.FomartKey(str, text2);
					if (!dictionary.ContainsKey(text3) && !string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(value))
					{
						dictionary.Add(text3, value);
					}
				}
			}
			return dictionary;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return null;
		}
	}

	public void CellChange_Event_01_KHBS(int iCol, int iRow)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			int[] array = GeneralFunctions.ArrayCellId("AA_38");
			HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_1"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = WinSettings.ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_2"));
			int num2 = HTKKFile.ColumnLettersToIndex("C");
			int num3 = HTKKFile.ColumnLettersToIndex("AT");
			int num4 = HTKKFile.ColumnLettersToIndex("C");
			int num5 = HTKKFile.ColumnLettersToIndex("B");
			int num6 = HTKKFile.ColumnLettersToIndex("E");
			int num7 = HTKKFile.ColumnLettersToIndex("AU");
			int num8 = HTKKFile.ColumnLettersToIndex("AT");
			int num9 = HTKKFile.ColumnLettersToIndex("AY");
			int num10 = HTKKFile.ColumnLettersToIndex("AZ");
			SheetView sView_01_KHBS = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			if (iCol == hTKKCMCCellInfo2.CellTemplate.ColIndex - 1 && iRow == hTKKCMCCellInfo2.CellTemplate.RowIndex - 1)
			{
				TinhSoTienPhatNopCham();
			}
			AnHienSoThueThuHoiHoanTienNopCham_01_KHBS();
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_3_ngay");
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_II/ct_1");
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_1");
			if (iRow == hTKKCMCCellInfo3.CellTemplate.RowIndex - 1 && iCol == hTKKCMCCellInfo3.CellTemplate.ColIndex - 1)
			{
				TinhSoNgayNhanDuocTienHoanThue_01_KHBS(sView_01_KHBS, is_edit: true);
				TinhSoTienChamNop_01_KHBS(sView_01_KHBS, is_edit: true);
			}
			else if (iRow == hTKKCMCCellInfo5.CellTemplate.RowIndex - 1 && iCol == hTKKCMCCellInfo5.CellTemplate.ColIndex - 1)
			{
				TinhSoTienChamNop_01_KHBS(sView_01_KHBS, is_edit: true);
			}
			else if (iRow == hTKKCMCCellInfo4.CellTemplate.RowIndex - 1 && iCol == hTKKCMCCellInfo4.CellTemplate.ColIndex - 1)
			{
				TinhSoTienChamNop_01_KHBS(sView_01_KHBS, is_edit: true);
			}
			AnHienSoNgayVaSoTienNopCham_01_KHBS();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhSoNgayNopCham()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			int[] array = GeneralFunctions.ArrayCellId("AA_36");
			HTKKCMCCellInfo hTKKCMCCellInfo = WinSettings.ToKhaiInfo.SheetInfos[num].StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_2"));
			winFpSpread.Sheets[num].Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = winFpSpread.Sheets[num].Cells[array[1], array[0]].Value;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhSoTienPhatNopCham()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			SheetView sheetView = winFpSpread.Sheets[num];
			if (!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				int[] array = GeneralFunctions.ArrayCellId("AA_38");
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3b"));
				if (hTKKCMCCellInfo != null)
				{
					sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value = sheetView.Cells[array[1], array[0]].Value;
				}
				return;
			}
			int key = HTKKFile.ColumnLettersToIndex("AW");
			int key2 = HTKKFile.ColumnLettersToIndex("U");
			int key3 = HTKKFile.ColumnLettersToIndex("AV");
			int value = HTKKFile.ColumnLettersToIndex("E");
			int value2 = HTKKFile.ColumnLettersToIndex("M");
			int value3 = HTKKFile.ColumnLettersToIndex("AN");
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			HTKKCMCTableInfo fromTable = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/TongCong10");
			int toTableIndex = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_3/ChiTietSoTienChamNop");
			Dictionary<int, int> dictionary = new Dictionary<int, int>();
			dictionary.Add(key2, value2);
			dictionary.Add(key, value);
			dictionary.Add(key3, value3);
			Dictionary<string, int> checkExists = new Dictionary<string, int>();
			GeneralFunctions.CopyTableToTable2(sheetView, hTKKCMCSheetInfo, fromTable, toTableIndex, dictionary, deleteDataOld: true, ref checkExists, usingCellType: true);
			GeneralFunctions.fpSpReindexingTableTemplate(sheetView, WinSettings.ToKhaiInfo.SheetInfos[num].Tables[toTableIndex]);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AddFullPLImportXML()
	{
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Expected O, but got Unknown
		try
		{
			XmlNodeList val = null;
			XmlDocument val2 = new XmlDocument();
			val2.LoadXml(WinSettings.xmlMenuNode.OuterXml);
			val = val2.GetElementsByTagName("Sheet");
			int num = ((!(WinSettings.ToKhaiInfo.KieuKHBS == "3")) ? 1 : 2);
			for (int i = 1; i < val.Count - num; i++)
			{
				if (!WinFnProcess.IsWorking(((XmlNode)val[i].Attributes["SheetName"]).Value))
				{
					HTKKCMCSheetInfo hTKKSheetInfo = WinFnProcess.GetHTKKSheetInfo(WinSettings.ToKhaiInfo, val, i, getTemplate: true, WinSettings.TTPeriodInfo);
					WinSettings.ToKhaiInfo.SheetInfos[i] = hTKKSheetInfo;
					WinSpreadProcess.ProcessSheet(i, winFpSpread, WinSettings.ToKhaiInfo.SheetInfos);
					winFpSpread.Sheets[i].Visible = true;
					WinSettings.DicSheet.Add(((XmlNode)val[i].Attributes["SheetName"]).Value, 2);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RevertFromXml()
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("BT");
			int num2 = HTKKFile.ColumnLettersToIndex("BS");
			int num3 = HTKKFile.ColumnLettersToIndex("BG");
			int num4 = HTKKFile.ColumnLettersToIndex("BH");
			int num5 = HTKKFile.ColumnLettersToIndex("BL");
			int num6 = HTKKFile.ColumnLettersToIndex("BF");
			int num7 = HTKKFile.ColumnLettersToIndex("BJ");
			int num8 = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num8];
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num8].Tables;
			if (_mapMct == null)
			{
				return;
			}
			DanhDauNeuCoPhanBo(sheetView, _mapMct);
			foreach (HTKKCMCTableInfo item in tables)
			{
				try
				{
					if (item.Header.StaticCells.Count == 0)
					{
						continue;
					}
					int rowIndex = item.Header.StaticCells[1].CellTemplate.RowIndex;
					int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
					string sheetIndex_XML = "0";
					string maHoSo = sheetView.Cells[rowIndex - 1, num - 1].Text.Trim();
					string sheetNameOrSheetIndexFromMaHoSo_01_1_KHBS = GetSheetNameOrSheetIndexFromMaHoSo_01_1_KHBS(maHoSo);
					if (string.IsNullOrEmpty(sheetNameOrSheetIndexFromMaHoSo_01_1_KHBS))
					{
						continue;
					}
					if (sheetNameOrSheetIndexFromMaHoSo_01_1_KHBS == "0")
					{
						sheetIndex_XML = sheetNameOrSheetIndexFromMaHoSo_01_1_KHBS;
					}
					else
					{
						string sheetName = GetSheetNameOrSheetIndexFromMaHoSo_01_1_KHBS(maHoSo);
						XmlNode val = ((IEnumerable)_mapMct.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["SheetName"] != null && ((XmlNode)x.Attributes["SheetName"]).Value.Trim() == sheetName));
						if (val != null)
						{
							sheetIndex_XML = ((XmlNode)val.Attributes["SheetIndex"]).Value.Trim();
						}
					}
					if (string.IsNullOrEmpty(sheetIndex_XML))
					{
						continue;
					}
					sheetView.Cells[rowIndex - 1, num2 - 1].Value = sheetIndex_XML;
					for (int num9 = 0; num9 < item.RowsValues.Count; num9++)
					{
						try
						{
							string MCT_XML = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num3 - 1].Text.Trim();
							string caption_XML = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num4 - 1].Text.Trim();
							string text = "";
							List<string> list = new List<string>();
							string value = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num6 - 1].Text.Trim();
							string value2 = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num7 - 1].Text.Trim();
							string text2 = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num5 - 1].Text.Trim();
							string text3 = "";
							if (WinSettings.MenuID != "N8")
							{
								if (!string.IsNullOrEmpty(text2) && (!string.IsNullOrEmpty(value) || !string.IsNullOrEmpty(value2)))
								{
									text = text2;
								}
								if (string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2) && text2.Contains(" - "))
								{
									List<string> list2 = GeneralFunctions.SplitString(text2, " - ");
									if (list2.Count > 0)
									{
										list = list2;
									}
								}
							}
							if (string.IsNullOrEmpty(MCT_XML) || string.IsNullOrEmpty(caption_XML) || _mapMct == null)
							{
								continue;
							}
							if (caption_XML.Contains(" - Tiểu mục") && string.IsNullOrEmpty(text2))
							{
								string[] array = caption_XML.Split(new char[1] { ' ' });
								text3 = array[array.Count() - 1];
								caption_XML = caption_XML.Substring(0, caption_XML.LastIndexOf(" - Tiểu mục " + text3)).Trim();
							}
							XmlNode val2 = null;
							IEnumerable<XmlNode> enumerable = null;
							string caption = "";
							enumerable = from XmlNode x in (IEnumerable)_mapMct.ChildNodes
								where ((XmlNode)x.Attributes["SheetIndex"]).Value.Trim() == sheetIndex_XML && ((XmlNode)x.Attributes["MCT"]).Value.Trim() == MCT_XML && x.Attributes["Caption"] != null && ((XmlNode)x.Attributes["Caption"]).Value.Trim() == caption_XML
								select x;
							if (!string.IsNullOrEmpty(text) && enumerable.Count() > 1)
							{
								string tablePath = GetTablePathFromPrefix_01_1_KHBS(enumerable, int.Parse(sheetIndex_XML), text);
								if (!string.IsNullOrEmpty(tablePath))
								{
									enumerable = from XmlNode x in (IEnumerable)_mapMct.ChildNodes
										where ((XmlNode)x.Attributes["SheetIndex"]).Value.Trim() == sheetIndex_XML && ((XmlNode)x.Attributes["MCT"]).Value.Trim() == MCT_XML && x.Attributes["Caption"] != null && ((XmlNode)x.Attributes["Caption"]).Value.Trim() == caption_XML && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath
										select x;
									if (enumerable != null && enumerable.Count() == 1)
									{
										val2 = enumerable.FirstOrDefault();
									}
								}
							}
							if (val2 == null && list.Count > 0 && enumerable.Count() > 1)
							{
								string text4 = "";
								foreach (string item2 in list)
								{
									text4 = ((!string.IsNullOrEmpty(text4)) ? (text4 + " - " + item2) : item2);
									string tablePath2 = GetTablePathFromPrefix_01_1_KHBS(enumerable, int.Parse(sheetIndex_XML), text4);
									if (!string.IsNullOrEmpty(tablePath2))
									{
										enumerable = from XmlNode x in (IEnumerable)_mapMct.ChildNodes
											where ((XmlNode)x.Attributes["SheetIndex"]).Value.Trim() == sheetIndex_XML && ((XmlNode)x.Attributes["MCT"]).Value.Trim() == MCT_XML && x.Attributes["Caption"] != null && ((XmlNode)x.Attributes["Caption"]).Value.Trim() == caption_XML && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath2
											select x;
										if (enumerable != null && enumerable.Count() == 1)
										{
											val2 = enumerable.FirstOrDefault();
											break;
										}
									}
								}
							}
							if (val2 == null && (enumerable == null || (enumerable != null && enumerable.Count() != 1)))
							{
								caption = GeneralFunctions.FormatKey2(MCT_XML, caption_XML, " - ");
								enumerable = from XmlNode x in (IEnumerable)_mapMct.ChildNodes
									where ((XmlNode)x.Attributes["SheetIndex"]).Value.Trim() == sheetIndex_XML && ((XmlNode)x.Attributes["MCT"]).Value.Trim() == MCT_XML && x.Attributes["Caption"] != null && ((XmlNode)x.Attributes["Caption"]).Value.Trim() == caption
									select x;
								if (!string.IsNullOrEmpty(text) && enumerable.Count() > 1)
								{
									string tablePath3 = GetTablePathFromPrefix_01_1_KHBS(enumerable, int.Parse(sheetIndex_XML), text);
									if (!string.IsNullOrEmpty(tablePath3))
									{
										enumerable = from XmlNode x in (IEnumerable)_mapMct.ChildNodes
											where ((XmlNode)x.Attributes["SheetIndex"]).Value.Trim() == sheetIndex_XML && ((XmlNode)x.Attributes["MCT"]).Value.Trim() == MCT_XML && x.Attributes["Caption"] != null && ((XmlNode)x.Attributes["Caption"]).Value.Trim() == caption && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath3
											select x;
										if (enumerable != null && enumerable.Count() == 1)
										{
											val2 = enumerable.FirstOrDefault();
										}
									}
								}
								if (val2 == null && list.Count > 0 && enumerable.Count() > 1)
								{
									string text4 = "";
									foreach (string item3 in list)
									{
										text4 = ((!string.IsNullOrEmpty(text4)) ? (text4 + " - " + item3) : item3);
										string tablePath4 = GetTablePathFromPrefix_01_1_KHBS(enumerable, int.Parse(sheetIndex_XML), text4);
										if (!string.IsNullOrEmpty(tablePath4))
										{
											enumerable = from XmlNode x in (IEnumerable)_mapMct.ChildNodes
												where ((XmlNode)x.Attributes["SheetIndex"]).Value.Trim() == sheetIndex_XML && ((XmlNode)x.Attributes["MCT"]).Value.Trim() == MCT_XML && x.Attributes["Caption"] != null && ((XmlNode)x.Attributes["Caption"]).Value.Trim() == caption && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath4
												select x;
											if (enumerable != null && enumerable.Count() == 1)
											{
												val2 = enumerable.FirstOrDefault();
												break;
											}
										}
									}
								}
							}
							if (val2 != null || (enumerable != null && (enumerable == null || enumerable.Count() == 1)))
							{
								if (val2 == null && enumerable != null && enumerable.Count() == 1)
								{
									val2 = enumerable.FirstOrDefault();
								}
								int num10 = ((val2.Attributes["Dynamic"] == null) ? (-1) : Convert.ToInt32(((XmlNode)val2.Attributes["Dynamic"]).Value));
								if (num10 == 1)
								{
									RevertDynamicHeaderFromXml_01_1_KHBS(val2, dataLocationTemplateRowStart + num9);
									RevertDynamicFromXml_01_1_KHBS(val2, dataLocationTemplateRowStart + num9);
								}
								else
								{
									RevertStaticFromXml_01_1_KHBS(val2, dataLocationTemplateRowStart + num9, text3);
								}
							}
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
					}
					WinSpreadProcess.fpSpReindexingTableTemplate(sheetView, item);
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
			}
			RevertDynamicFromXml_01_KHBS();
			RevertPhanBo_01_KHBS();
			LockUnlock_01_1_KHBS();
			WinDoWorkEvent.SetUpDataKhbs(winFpSpread, WinSettings.TTPeriodInfo, isClickTongHopSNNC: true, WinSettings.HTKKHeaderInfo);
			isClickTongHopKHBS = true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public string GetTablePathFromPrefix_01_1_KHBS(IEnumerable<XmlNode> nodes, int sheetIndex, string prefix)
	{
		try
		{
			if (!string.IsNullOrEmpty(prefix) && nodes != null && nodes.Count() > 1)
			{
				foreach (XmlNode node in nodes)
				{
					foreach (HTKKCMCTableInfo table in WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].Tables)
					{
						string text = ((node.Attributes["Prefix"] == null) ? "" : ((XmlNode)node.Attributes["Prefix"]).Value.Trim());
						if (!string.IsNullOrEmpty(text) && text == prefix)
						{
							return (node.Attributes["TablePath"] == null) ? "" : ((XmlNode)node.Attributes["TablePath"]).Value.Trim();
						}
						string prefixPath = ((node.Attributes["PrefixPath"] == null) ? "" : ((XmlNode)node.Attributes["PrefixPath"]).Value.Trim());
						if (string.IsNullOrEmpty(prefixPath))
						{
							continue;
						}
						HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == prefixPath);
						if (hTKKCMCTableCellInfo != null)
						{
							List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(winFpSpread.Sheets[sheetIndex], table, hTKKCMCTableCellInfo.CellTemplate.ColIndex, prefix);
							if (rowIndexInSViewByCompare.Count > 0)
							{
								return table.TablePath;
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return "";
	}

	public void RevertDynamicFromXml_01_1_KHBS(XmlNode mapCt, int row_01_1_KHBS)
	{
		try
		{
			int sheetIndex = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			string sheetname = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string tablePath = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			string value = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string text = ((mapCt.Attributes["TCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim());
			string text2 = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string text3 = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string value2 = ((mapCt.Attributes["Caption"] == null) ? "" : ((XmlNode)mapCt.Attributes["Caption"]).Value.ToString().Trim());
			string text4 = ((mapCt.Attributes["DVTPath"] != null) ? ((XmlNode)mapCt.Attributes["DVTPath"]).Value.ToString().Trim() : "");
			string text5 = ((mapCt.Attributes["MaDVTPath"] != null) ? ((XmlNode)mapCt.Attributes["MaDVTPath"]).Value.ToString().Trim() : "");
			string text6 = ((mapCt.Attributes["LoaiChiTieu"] != null) ? ((XmlNode)mapCt.Attributes["LoaiChiTieu"]).Value.ToString().Trim() : "");
			string text7 = ((mapCt.Attributes["Manual"] != null) ? ((XmlNode)mapCt.Attributes["Manual"]).Value.ToString().Trim() : "");
			int num = ((mapCt.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
			string maTieuMuc = "";
			XmlNode mapTieuMuc = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath)));
			if (mapTieuMuc != null && mapTieuMuc.Attributes["MaTieuMuc"] != null)
			{
				maTieuMuc = ((XmlNode)mapTieuMuc.Attributes["MaTieuMuc"]).Value.Trim();
			}
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("BY");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int col = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("V");
			int num7 = HTKKFile.ColumnLettersToIndex("Z");
			int col2 = HTKKFile.ColumnLettersToIndex("AD");
			int num8 = HTKKFile.ColumnLettersToIndex("AH");
			int num9 = HTKKFile.ColumnLettersToIndex("CA");
			int num10 = HTKKFile.ColumnLettersToIndex("BR");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("BS");
			int colSheetName = HTKKFile.ColumnLettersToIndex("BT");
			int num11 = HTKKFile.ColumnLettersToIndex("G");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("BV");
			int num12 = HTKKFile.ColumnLettersToIndex("BX");
			int num13 = HTKKFile.ColumnLettersToIndex("BW");
			int num14 = HTKKFile.ColumnLettersToIndex("BD");
			int num15 = HTKKFile.ColumnLettersToIndex("BN");
			int num16 = HTKKFile.ColumnLettersToIndex("BO");
			int num17 = HTKKFile.ColumnLettersToIndex("BM");
			int num18 = HTKKFile.ColumnLettersToIndex("BD");
			int num19 = HTKKFile.ColumnLettersToIndex("BE");
			int num20 = HTKKFile.ColumnLettersToIndex("BF");
			int num21 = HTKKFile.ColumnLettersToIndex("BG");
			int num22 = HTKKFile.ColumnLettersToIndex("BH");
			int num23 = HTKKFile.ColumnLettersToIndex("BI");
			int num24 = HTKKFile.ColumnLettersToIndex("BJ");
			int num25 = HTKKFile.ColumnLettersToIndex("BK");
			int num26 = HTKKFile.ColumnLettersToIndex("BL");
			int num27 = HTKKFile.ColumnLettersToIndex("BQ");
			int num28 = HTKKFile.ColumnLettersToIndex("CE");
			int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex_KHBS];
			string text8 = sheetView.Cells[row_01_1_KHBS - 1, num20 - 1].Text.Trim();
			string text9 = sheetView.Cells[row_01_1_KHBS - 1, num19 - 1].Text.Trim();
			string text10 = sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Text.Trim();
			string text11 = sheetView.Cells[row_01_1_KHBS - 1, num21 - 1].Text.Trim();
			string text12 = sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Text.Trim();
			string text13 = sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Text.Trim();
			string text14 = sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Text.Trim();
			string text15 = sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Text.Trim();
			string valueCompare = sheetView.Cells[row_01_1_KHBS - 1, num6 - 1].Text.Trim();
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = null;
			if (hTKKCMCSheetInfo == null)
			{
				hTKKCMCSheetInfo2 = WinFnProcess.ReadSheet_LanTruoc(sheetIndex);
			}
			SheetView sheetView2 = ((hTKKCMCSheetInfo == null) ? null : winFpSpread.Sheets[sheetIndex]);
			HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo != null) ? hTKKCMCSheetInfo.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath) : hTKKCMCSheetInfo2.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath));
			HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo?.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
			if (hTKKCMCTableCellInfo == null)
			{
				return;
			}
			int num29 = hTKKCMCTableInfo?.Location.DataLocationTemplateRowStart ?? (-1);
			int num30 = -1;
			if (!string.IsNullOrEmpty(text) && hTKKCMCTableInfo != null && sheetView2 != null)
			{
				List<HTKKCMCTableCellInfo> list = new List<HTKKCMCTableCellInfo>();
				List<string> list2 = GeneralFunctions.SplitString(text, "###");
				foreach (string path in list2)
				{
					List<HTKKCMCTableCellInfo> cells = hTKKCMCTableInfo.RowsInfo.Cells;
					Func<HTKKCMCTableCellInfo, bool> predicate = (HTKKCMCTableCellInfo x) => x.Path == path;
					list.Add(cells.FirstOrDefault(predicate));
				}
				List<int> list3 = new List<int>();
				foreach (HTKKCMCTableCellInfo item in list)
				{
					list3.Add(item.CellTemplate.ColIndex);
				}
				if (list3.Count > 1)
				{
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares2(sheetView2, hTKKCMCTableInfo, list3, text14);
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView2, hTKKCMCTableInfo, hTKKCMCTableCellInfo.CellTemplate.ColIndex, valueCompare);
					List<int> valuesExistsCompare = GeneralFunctions.GetValuesExistsCompare(rowIndexInSViewByCompares, rowIndexInSViewByCompare);
					if (valuesExistsCompare.Count == 1)
					{
						num30 = valuesExistsCompare[0];
					}
				}
				else if (list3.Count > 0)
				{
					List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView2, hTKKCMCTableInfo, list3[0], (!string.IsNullOrEmpty(text8)) ? text8 : ((!string.IsNullOrEmpty(text12)) ? text12 : text14));
					if (rowIndexInSViewByCompare2.Count == 1)
					{
						num30 = rowIndexInSViewByCompare2[0];
					}
				}
				if (mapTieuMuc != null && string.IsNullOrEmpty(maTieuMuc) && num30 >= 0 && mapTieuMuc.Attributes["MaTieuMucPath"] != null && sheetView2 != null && hTKKCMCTableInfo != null)
				{
					maTieuMuc = TimTieuMucDynamicInTable(mapTieuMuc, sheetView2, hTKKCMCTableInfo, num30);
				}
			}
			else
			{
				int num31 = FindIndexTieuMuc(hTKKCMCTableInfo, sheetIndex, tablePath, MCTPath, ref mapTieuMuc, ref maTieuMuc);
				if (string.IsNullOrEmpty(maTieuMuc) && num31 >= 0)
				{
					foreach (HTKKCMCRowValue rowsValue in hTKKCMCTableInfo.RowsValues)
					{
						maTieuMuc = rowsValue.RowValue[num31];
						if (!string.IsNullOrEmpty(maTieuMuc))
						{
							break;
						}
					}
				}
			}
			ThongTinThemChoKHBSMoi(sheetView, row_01_1_KHBS, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text3);
			sheetView.Cells[row_01_1_KHBS - 1, num14 - 1].Value = text6;
			sheetView.Cells[row_01_1_KHBS - 1, num15 - 1].Value = (string.IsNullOrEmpty(text7) ? "0" : text7);
			if (text6 == "1")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num20 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num19 - 1].Value;
			}
			if (text6 == "2")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num21 - 1].Value;
			}
			if (text6 == "3")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Value;
			}
			if (text6 == "4")
			{
				string value3 = sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Value.ToString().Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = value3;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value;
			}
			sheetView.Cells[row_01_1_KHBS - 1, num5 - 1].Value = value2;
			sheetView.Cells[row_01_1_KHBS - 1, num12 - 1].Value = value;
			sheetView.Cells[row_01_1_KHBS - 1, num27 - 1].Value = Functions.FomartKey(tablePath, MCTPath);
			if (sheetView2 != null && num29 > 0)
			{
				GeneralFunctions.UpsertCellType(sheetIndex_KHBS, row_01_1_KHBS, num6, sheetView2, (num30 < 0) ? num29 : num30, hTKKCMCTableCellInfo.CellTemplate.ColIndex);
				FormatCol4567(sheetIndex_KHBS, row_01_1_KHBS, num6, sheetView, col, num6, num7, col2);
			}
			else
			{
				GeneralFunctions.FormatCol4567KHBSBySinfo(hTKKCMCTableCellInfo, sheetView, row_01_1_KHBS, col, num6, num7, col2);
			}
			sheetView.Cells[row_01_1_KHBS - 1, num10 - 1].Value = maTieuMuc;
			sheetView.Cells[row_01_1_KHBS - 1, num28 - 1].Value = num;
			sheetView.Cells[row_01_1_KHBS - 1, num13 - 1].Value = hTKKCMCTableCellInfo.Controltype;
			sheetView.Cells[row_01_1_KHBS - 1, num7 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num17 - 1].Value;
			if (text6 != "4" && text6 != "2" && !string.IsNullOrEmpty(text14))
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text14, text8, " - ");
			}
			if (!string.IsNullOrEmpty(text3))
			{
				ChangeCellTypeHoachToan(row_01_1_KHBS, num6);
				if (sheetView2 != null)
				{
					SetLoaiTienStatic(sheetView2, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
				}
			}
			GeneralFunctions.SetRowHeight(sheetView, row_01_1_KHBS);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RevertDynamicHeaderFromXml_01_1_KHBS(XmlNode mapCt, int row_01_1_KHBS)
	{
		try
		{
			int sheetIndex = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			string sheetname = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string tablePath = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			string value = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string text = ((mapCt.Attributes["TCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim());
			string text2 = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string text3 = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string value2 = ((mapCt.Attributes["Caption"] == null) ? "" : ((XmlNode)mapCt.Attributes["Caption"]).Value.ToString().Trim());
			string text4 = ((mapCt.Attributes["DVTPath"] != null) ? ((XmlNode)mapCt.Attributes["DVTPath"]).Value.ToString().Trim() : "");
			string text5 = ((mapCt.Attributes["MaDVTPath"] != null) ? ((XmlNode)mapCt.Attributes["MaDVTPath"]).Value.ToString().Trim() : "");
			string text6 = ((mapCt.Attributes["LoaiChiTieu"] != null) ? ((XmlNode)mapCt.Attributes["LoaiChiTieu"]).Value.ToString().Trim() : "");
			string text7 = ((mapCt.Attributes["Manual"] != null) ? ((XmlNode)mapCt.Attributes["Manual"]).Value.ToString().Trim() : "");
			int num = ((mapCt.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
			string maTieuMuc = "";
			XmlNode mapTieuMuc = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] != null && ((XmlNode)x.Attributes["TablePath"]).Value == tablePath)));
			if (mapTieuMuc != null && mapTieuMuc.Attributes["MaTieuMuc"] != null)
			{
				maTieuMuc = ((XmlNode)mapTieuMuc.Attributes["MaTieuMuc"]).Value.Trim();
			}
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("BY");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int col = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("V");
			int num7 = HTKKFile.ColumnLettersToIndex("Z");
			int col2 = HTKKFile.ColumnLettersToIndex("AD");
			int num8 = HTKKFile.ColumnLettersToIndex("AH");
			int num9 = HTKKFile.ColumnLettersToIndex("CA");
			int num10 = HTKKFile.ColumnLettersToIndex("BR");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("BS");
			int colSheetName = HTKKFile.ColumnLettersToIndex("BT");
			int num11 = HTKKFile.ColumnLettersToIndex("G");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("BV");
			int num12 = HTKKFile.ColumnLettersToIndex("BX");
			int num13 = HTKKFile.ColumnLettersToIndex("BW");
			int num14 = HTKKFile.ColumnLettersToIndex("BD");
			int num15 = HTKKFile.ColumnLettersToIndex("BN");
			int num16 = HTKKFile.ColumnLettersToIndex("BO");
			int num17 = HTKKFile.ColumnLettersToIndex("BM");
			int num18 = HTKKFile.ColumnLettersToIndex("BD");
			int num19 = HTKKFile.ColumnLettersToIndex("BE");
			int num20 = HTKKFile.ColumnLettersToIndex("BF");
			int num21 = HTKKFile.ColumnLettersToIndex("BG");
			int num22 = HTKKFile.ColumnLettersToIndex("BH");
			int num23 = HTKKFile.ColumnLettersToIndex("BI");
			int num24 = HTKKFile.ColumnLettersToIndex("BJ");
			int num25 = HTKKFile.ColumnLettersToIndex("BK");
			int num26 = HTKKFile.ColumnLettersToIndex("BL");
			int num27 = HTKKFile.ColumnLettersToIndex("BQ");
			int num28 = HTKKFile.ColumnLettersToIndex("CE");
			int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex_KHBS];
			string text8 = sheetView.Cells[row_01_1_KHBS - 1, num20 - 1].Text.Trim();
			string text9 = sheetView.Cells[row_01_1_KHBS - 1, num19 - 1].Text.Trim();
			string text10 = sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Text.Trim();
			string text11 = sheetView.Cells[row_01_1_KHBS - 1, num21 - 1].Text.Trim();
			string text12 = sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Text.Trim();
			string text13 = sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Text.Trim();
			string text14 = sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Text.Trim();
			string text15 = sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Text.Trim();
			string valueCompare = sheetView.Cells[row_01_1_KHBS - 1, num6 - 1].Text.Trim();
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = null;
			if (hTKKCMCSheetInfo == null)
			{
				hTKKCMCSheetInfo2 = WinFnProcess.ReadSheet_LanTruoc(sheetIndex);
			}
			SheetView sheetView2 = ((hTKKCMCSheetInfo == null) ? null : winFpSpread.Sheets[sheetIndex]);
			HTKKCMCTableInfo hTKKCMCTableInfo = ((hTKKCMCSheetInfo != null) ? hTKKCMCSheetInfo.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath) : hTKKCMCSheetInfo2.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath));
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo?.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
			if (hTKKCMCCellInfo == null)
			{
				return;
			}
			int num29 = hTKKCMCTableInfo?.Location.DataLocationTemplateRowStart ?? (-1);
			int num30 = -1;
			if (!string.IsNullOrEmpty(text) && hTKKCMCTableInfo != null && sheetView2 != null)
			{
				List<HTKKCMCCellInfo> list = new List<HTKKCMCCellInfo>();
				List<string> list2 = GeneralFunctions.SplitString(text, "###");
				foreach (string path in list2)
				{
					List<HTKKCMCCellInfo> staticCells = hTKKCMCTableInfo.Header.StaticCells;
					Func<HTKKCMCCellInfo, bool> predicate = (HTKKCMCCellInfo x) => x.Path == path;
					list.Add(staticCells.FirstOrDefault(predicate));
				}
				List<int> list3 = new List<int>();
				foreach (HTKKCMCCellInfo item in list)
				{
					list3.Add(item.CellTemplate.ColIndex);
				}
				if (list3.Count > 1)
				{
					List<int> rowIndexInSViewByCompares = GeneralFunctions.GetRowIndexInSViewByCompares2(sheetView2, hTKKCMCTableInfo, list3, text14);
					List<int> rowIndexInSViewByCompare = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView2, hTKKCMCTableInfo, hTKKCMCCellInfo.CellTemplate.ColIndex, valueCompare);
					List<int> valuesExistsCompare = GeneralFunctions.GetValuesExistsCompare(rowIndexInSViewByCompares, rowIndexInSViewByCompare);
					if (valuesExistsCompare.Count == 1)
					{
						num30 = valuesExistsCompare[0];
					}
				}
				else if (list3.Count > 0)
				{
					List<int> rowIndexInSViewByCompare2 = GeneralFunctions.GetRowIndexInSViewByCompare(sheetView2, hTKKCMCTableInfo, list3[0], (!string.IsNullOrEmpty(text8)) ? text8 : ((!string.IsNullOrEmpty(text12)) ? text12 : text14));
					if (rowIndexInSViewByCompare2.Count == 1)
					{
						num30 = rowIndexInSViewByCompare2[0];
					}
				}
				if (mapTieuMuc != null && string.IsNullOrEmpty(maTieuMuc) && num30 >= 0 && mapTieuMuc.Attributes["MaTieuMucPath"] != null && sheetView2 != null && hTKKCMCTableInfo != null)
				{
					maTieuMuc = TimTieuMucDynamicInTable(mapTieuMuc, sheetView2, hTKKCMCTableInfo, num30);
				}
			}
			else
			{
				int num31 = FindIndexTieuMuc(hTKKCMCTableInfo, sheetIndex, tablePath, MCTPath, ref mapTieuMuc, ref maTieuMuc);
				if (string.IsNullOrEmpty(maTieuMuc) && num31 >= 0)
				{
					foreach (HTKKCMCRowValue rowsValue in hTKKCMCTableInfo.RowsValues)
					{
						maTieuMuc = rowsValue.RowValue[num31];
						if (!string.IsNullOrEmpty(maTieuMuc))
						{
							break;
						}
					}
				}
			}
			ThongTinThemChoKHBSMoi(sheetView, row_01_1_KHBS, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text3);
			sheetView.Cells[row_01_1_KHBS - 1, num14 - 1].Value = text6;
			sheetView.Cells[row_01_1_KHBS - 1, num15 - 1].Value = (string.IsNullOrEmpty(text7) ? "0" : text7);
			if (text6 == "1")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num20 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num19 - 1].Value;
			}
			if (text6 == "2")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num21 - 1].Value;
			}
			if (text6 == "3")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Value;
			}
			if (text6 == "4")
			{
				string value3 = sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Value.ToString().Trim();
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = value3;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value;
			}
			sheetView.Cells[row_01_1_KHBS - 1, num5 - 1].Value = value2;
			sheetView.Cells[row_01_1_KHBS - 1, num12 - 1].Value = value;
			sheetView.Cells[row_01_1_KHBS - 1, num27 - 1].Value = Functions.FomartKey(tablePath, MCTPath);
			if (sheetView2 != null && num30 > 0)
			{
				GeneralFunctions.UpsertCellType(sheetIndex_KHBS, row_01_1_KHBS, num6, sheetView2, (num30 < 0) ? num29 : num30, hTKKCMCCellInfo.CellTemplate.ColIndex);
				FormatCol4567(sheetIndex_KHBS, row_01_1_KHBS, num6, sheetView, col, num6, num7, col2);
			}
			else
			{
				GeneralFunctions.FormatCol4567KHBSBySinfoStatic(hTKKCMCCellInfo, sheetView, row_01_1_KHBS, col, num6, num7, col2);
			}
			sheetView.Cells[row_01_1_KHBS - 1, num10 - 1].Value = maTieuMuc;
			sheetView.Cells[row_01_1_KHBS - 1, num28 - 1].Value = num;
			sheetView.Cells[row_01_1_KHBS - 1, num13 - 1].Value = hTKKCMCCellInfo.Controltype;
			sheetView.Cells[row_01_1_KHBS - 1, num7 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num17 - 1].Value;
			if (text6 != "4" && text6 != "2" && !string.IsNullOrEmpty(text14))
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = GeneralFunctions.FormatKey2(text14, text8, " - ");
			}
			if (!string.IsNullOrEmpty(text3))
			{
				ChangeCellTypeHoachToan(row_01_1_KHBS, num6);
				if (sheetView2 != null)
				{
					SetLoaiTienStatic(sheetView2, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
				}
			}
			GeneralFunctions.SetRowHeight(sheetView, row_01_1_KHBS);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RevertStaticFromXml_01_1_KHBS(XmlNode mapCt, int row_01_1_KHBS, string maTieuMuc)
	{
		try
		{
			int sheetIndex = ((mapCt.Attributes["SheetIndex"] == null) ? (-1) : Convert.ToInt32(((XmlNode)mapCt.Attributes["SheetIndex"]).Value));
			string sheetname = ((mapCt.Attributes["SheetName"] == null) ? "" : ((XmlNode)mapCt.Attributes["SheetName"]).Value.ToString().Trim());
			string text = ((mapCt.Attributes["TablePath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TablePath"]).Value.ToString().Trim());
			string value = ((mapCt.Attributes["MCT"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCT"]).Value.ToString().Trim());
			string MCTPath = ((mapCt.Attributes["MCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MCTPath"]).Value.ToString().Trim());
			string text2 = ((mapCt.Attributes["TCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["TCTPath"]).Value.ToString().Trim());
			string text3 = ((mapCt.Attributes["MaTCTPath"] == null) ? "" : ((XmlNode)mapCt.Attributes["MaTCTPath"]).Value.ToString().Trim());
			string text4 = ((mapCt.Attributes["LoaiThue"] == null) ? "" : ((XmlNode)mapCt.Attributes["LoaiThue"]).Value.Trim());
			string text5 = ((mapCt.Attributes["Caption"] == null) ? "" : ((XmlNode)mapCt.Attributes["Caption"]).Value.ToString().Trim());
			string text6 = ((mapCt.Attributes["DVTPath"] != null) ? ((XmlNode)mapCt.Attributes["DVTPath"]).Value.ToString().Trim() : "");
			string text7 = ((mapCt.Attributes["MaDVTPath"] != null) ? ((XmlNode)mapCt.Attributes["MaDVTPath"]).Value.ToString().Trim() : "");
			string text8 = ((mapCt.Attributes["LoaiChiTieu"] != null) ? ((XmlNode)mapCt.Attributes["LoaiChiTieu"]).Value.ToString().Trim() : "");
			string text9 = ((mapCt.Attributes["Manual"] != null) ? ((XmlNode)mapCt.Attributes["Manual"]).Value.ToString().Trim() : "");
			int num = ((mapCt.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)mapCt.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
			XmlNode val = null;
			if (string.IsNullOrEmpty(maTieuMuc))
			{
				val = ((_mapMctTieuMuc == null) ? null : ((IEnumerable)_mapMctTieuMuc.ChildNodes).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == MCTPath && ((XmlNode)x.Attributes["SheetIndex"]).Value == sheetIndex.ToString() && x.Attributes["TablePath"] == null)));
			}
			if (string.IsNullOrEmpty(maTieuMuc) && val != null && val.Attributes["MaTieuMuc"] != null)
			{
				maTieuMuc = ((XmlNode)val.Attributes["MaTieuMuc"]).Value.Trim();
			}
			int num2 = HTKKFile.ColumnLettersToIndex("B");
			int num3 = HTKKFile.ColumnLettersToIndex("H");
			int num4 = HTKKFile.ColumnLettersToIndex("BY");
			int num5 = HTKKFile.ColumnLettersToIndex("P");
			int col = HTKKFile.ColumnLettersToIndex("R");
			int num6 = HTKKFile.ColumnLettersToIndex("V");
			int num7 = HTKKFile.ColumnLettersToIndex("Z");
			int col2 = HTKKFile.ColumnLettersToIndex("AD");
			int num8 = HTKKFile.ColumnLettersToIndex("AH");
			int num9 = HTKKFile.ColumnLettersToIndex("CA");
			int num10 = HTKKFile.ColumnLettersToIndex("BR");
			int colSheetIndex = HTKKFile.ColumnLettersToIndex("BS");
			int colSheetName = HTKKFile.ColumnLettersToIndex("BT");
			int num11 = HTKKFile.ColumnLettersToIndex("G");
			int colLoaiThue = HTKKFile.ColumnLettersToIndex("BV");
			int num12 = HTKKFile.ColumnLettersToIndex("BX");
			int num13 = HTKKFile.ColumnLettersToIndex("BW");
			int num14 = HTKKFile.ColumnLettersToIndex("BD");
			int num15 = HTKKFile.ColumnLettersToIndex("BN");
			int num16 = HTKKFile.ColumnLettersToIndex("BO");
			int num17 = HTKKFile.ColumnLettersToIndex("BM");
			int num18 = HTKKFile.ColumnLettersToIndex("BD");
			int num19 = HTKKFile.ColumnLettersToIndex("BE");
			int num20 = HTKKFile.ColumnLettersToIndex("BF");
			int num21 = HTKKFile.ColumnLettersToIndex("BG");
			int num22 = HTKKFile.ColumnLettersToIndex("BH");
			int num23 = HTKKFile.ColumnLettersToIndex("BI");
			int num24 = HTKKFile.ColumnLettersToIndex("BJ");
			int num25 = HTKKFile.ColumnLettersToIndex("BK");
			int num26 = HTKKFile.ColumnLettersToIndex("BL");
			int num27 = HTKKFile.ColumnLettersToIndex("BQ");
			int num28 = HTKKFile.ColumnLettersToIndex("CE");
			int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[sheetIndex_KHBS];
			string text10 = sheetView.Cells[row_01_1_KHBS - 1, num20 - 1].Text.Trim();
			string text11 = sheetView.Cells[row_01_1_KHBS - 1, num19 - 1].Text.Trim();
			string text12 = sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Text.Trim();
			string text13 = sheetView.Cells[row_01_1_KHBS - 1, num21 - 1].Text.Trim();
			string text14 = sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Text.Trim();
			string text15 = sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Text.Trim();
			string text16 = sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Text.Trim();
			string text17 = sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Text.Trim();
			string text18 = sheetView.Cells[row_01_1_KHBS - 1, num6 - 1].Text.Trim();
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = null;
			if (hTKKCMCSheetInfo == null)
			{
				hTKKCMCSheetInfo2 = WinFnProcess.ReadSheet_LanTruoc(sheetIndex);
			}
			SheetView sheetView2 = ((hTKKCMCSheetInfo == null) ? null : winFpSpread.Sheets[sheetIndex]);
			HTKKCMCCellInfo hTKKCMCCellInfo = ((hTKKCMCSheetInfo != null) ? hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath) : hTKKCMCSheetInfo2.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath));
			if (sheetView2 != null && val != null && string.IsNullOrEmpty(maTieuMuc) && val.Attributes["MaTieuMucPath"] != null)
			{
				maTieuMuc = TimTieuMucStatic(val, hTKKCMCSheetInfo, sheetView2);
			}
			else if (string.IsNullOrEmpty(maTieuMuc))
			{
				maTieuMuc = TimTieuMucStaticLanTruoc(val, hTKKCMCSheetInfo2);
			}
			ThongTinThemChoKHBSMoi(sheetView, row_01_1_KHBS, colSheetIndex, sheetIndex, colSheetName, sheetname, colLoaiThue, text4);
			sheetView.Cells[row_01_1_KHBS - 1, num14 - 1].Value = text8;
			sheetView.Cells[row_01_1_KHBS - 1, num15 - 1].Value = (string.IsNullOrEmpty(text9) ? "0" : text9);
			if (text8 == "1")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num20 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num19 - 1].Value;
			}
			if (text8 == "2")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num22 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num21 - 1].Value;
			}
			if (text8 == "3")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num24 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num23 - 1].Value;
			}
			if (text8 == "4")
			{
				sheetView.Cells[row_01_1_KHBS - 1, num3 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num26 - 1].Value;
				sheetView.Cells[row_01_1_KHBS - 1, num4 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num25 - 1].Value;
			}
			sheetView.Cells[row_01_1_KHBS - 1, num5 - 1].Value = value;
			sheetView.Cells[row_01_1_KHBS - 1, num12 - 1].Value = value;
			sheetView.Cells[row_01_1_KHBS - 1, num27 - 1].Value = MCTPath;
			if (sheetView2 != null)
			{
				GeneralFunctions.UpsertCellType(sheetIndex_KHBS, row_01_1_KHBS, num6, sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex);
				FormatCol4567(sheetIndex_KHBS, row_01_1_KHBS, num6, sheetView, col, num6, num7, col2);
			}
			else
			{
				GeneralFunctions.FormatCol4567KHBSBySinfoStatic(hTKKCMCCellInfo, sheetView, row_01_1_KHBS, col, num6, num7, col2);
			}
			sheetView.Cells[row_01_1_KHBS - 1, num10 - 1].Value = maTieuMuc;
			sheetView.Cells[row_01_1_KHBS - 1, num28 - 1].Value = num;
			sheetView.Cells[row_01_1_KHBS - 1, num13 - 1].Value = hTKKCMCCellInfo.Controltype;
			sheetView.Cells[row_01_1_KHBS - 1, num7 - 1].Value = sheetView.Cells[row_01_1_KHBS - 1, num17 - 1].Value;
			if (!string.IsNullOrEmpty(text4))
			{
				ChangeCellTypeHoachToan(row_01_1_KHBS, num6);
				if (sheetView2 != null)
				{
					SetLoaiTienStatic(sheetView2, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], mapCt, "");
				}
			}
			GeneralFunctions.SetRowHeight(sheetView, row_01_1_KHBS);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RevertDynamicFromXml_01_KHBS()
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("C");
			int num2 = HTKKFile.ColumnLettersToIndex("AT");
			int num3 = HTKKFile.ColumnLettersToIndex("O");
			int num4 = HTKKFile.ColumnLettersToIndex("AU");
			int num5 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView = winFpSpread.Sheets[num5];
			List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num5].Tables;
			int num6 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/ChiTiet");
			int num7 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_II/ChiTiet");
			int num8 = tables.FindIndex((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_III/ChiTiet");
			HTKKCMCTableInfo hTKKCMCTableInfo = tables[num6];
			HTKKCMCTableInfo hTKKCMCTableInfo2 = tables[num7];
			HTKKCMCTableInfo hTKKCMCTableInfo3 = tables[num8];
			int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
			for (int num9 = 0; num9 < hTKKCMCTableInfo.RowsValues.Count; num9++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num4 - 1].Text.Trim();
				string text2 = sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num2 - 1].Text.Trim();
				sheetView.Cells[dataLocationTemplateRowStart + num9 - 1, num - 1].Text = ((!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2)) ? GeneralFunctions.FormatKey2(text2, text, " - ") : text2);
				GeneralFunctions.SetRowHeight(sheetView, dataLocationTemplateRowStart + num9, 30f);
			}
			int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
			for (int num9 = 0; num9 < hTKKCMCTableInfo2.RowsValues.Count; num9++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart2 + num9 - 1, num4 - 1].Text.Trim();
				string text2 = sheetView.Cells[dataLocationTemplateRowStart2 + num9 - 1, num2 - 1].Text.Trim();
				sheetView.Cells[dataLocationTemplateRowStart2 + num9 - 1, num - 1].Text = ((!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2)) ? GeneralFunctions.FormatKey2(text2, text, " - ") : text2);
				GeneralFunctions.SetRowHeight(sheetView, dataLocationTemplateRowStart2 + num9, 30f);
			}
			int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
			for (int num9 = 0; num9 < hTKKCMCTableInfo3.RowsValues.Count; num9++)
			{
				string text = sheetView.Cells[dataLocationTemplateRowStart3 + num9 - 1, num4 - 1].Text.Trim();
				string text2 = sheetView.Cells[dataLocationTemplateRowStart3 + num9 - 1, num2 - 1].Text.Trim();
				sheetView.Cells[dataLocationTemplateRowStart3 + num9 - 1, num - 1].Text = ((!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2)) ? GeneralFunctions.FormatKey2(text2, text, " - ") : text2);
				GeneralFunctions.SetRowHeight(sheetView, dataLocationTemplateRowStart3 + num9, 30f);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ReSetControltypeAndConfig()
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("BQ");
			int col = HTKKFile.ColumnLettersToIndex("R");
			int num2 = HTKKFile.ColumnLettersToIndex("V");
			int col2 = HTKKFile.ColumnLettersToIndex("Z");
			int col3 = HTKKFile.ColumnLettersToIndex("AD");
			int num3 = HTKKFile.ColumnLettersToIndex("CE");
			int sheetIndex_KHBS = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex_KHBS];
			SheetView sheetView = winFpSpread.Sheets[sheetIndex_KHBS];
			List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo.Tables;
			for (int i = 0; i < tables.Count - 1; i++)
			{
				int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
				for (int j = 0; j < tables[i].RowsValues.Count; j++)
				{
					string text = sheetView.Cells[dataLocationTemplateRowStart + j - 1, num - 1].Text.Trim();
					if (string.IsNullOrEmpty(text))
					{
						continue;
					}
					XmlNode itemMapMct = GetItemMapMct(_mapMct, text);
					if (itemMapMct == null)
					{
						continue;
					}
					string MCTPath = ((itemMapMct.Attributes["MCTPath"] != null) ? ((XmlNode)itemMapMct.Attributes["MCTPath"]).Value : "");
					int sheetIndex = ((itemMapMct.Attributes["SheetIndex"] != null) ? Convert.ToInt32(((XmlNode)itemMapMct.Attributes["SheetIndex"]).Value) : 0);
					int num4 = ((itemMapMct.Attributes["Dynamic"] != null) ? Convert.ToInt32(((XmlNode)itemMapMct.Attributes["Dynamic"]).Value) : 0);
					string tablePath = ((itemMapMct.Attributes["TablePath"] != null) ? ((XmlNode)itemMapMct.Attributes["TablePath"]).Value : "");
					string value = ((itemMapMct.Attributes["LoaiThue"] != null) ? ((XmlNode)itemMapMct.Attributes["LoaiThue"]).Value.Trim() : "");
					int num5 = ((itemMapMct.Attributes["NopThua"] != null) ? Convert.ToInt32(((XmlNode)itemMapMct.Attributes["NopThua"]).Value.ToString().Trim()) : 0);
					ICellType cellType = sheetView.Cells[0, 0].CellType;
					HTKKCMCSheetInfo hTKKCMCSheetInfo2 = WinSettings.ToKhaiInfo.SheetInfos[sheetIndex];
					SheetView sheetView2 = ((hTKKCMCSheetInfo2 == null) ? null : winFpSpread.Sheets[sheetIndex]);
					if (num4 == 1)
					{
						if (hTKKCMCSheetInfo2 != null && sheetView2 != null)
						{
							HTKKCMCTableInfo hTKKCMCTableInfo = hTKKCMCSheetInfo2.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath);
							if (hTKKCMCTableInfo != null)
							{
								int dataLocationTemplateRowStart2 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
								HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
								if (hTKKCMCTableCellInfo != null)
								{
									GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + j, num2, sheetView2, dataLocationTemplateRowStart2, hTKKCMCTableCellInfo.CellTemplate.ColIndex);
									FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + j, num2, sheetView, col, num2, col2, col3);
								}
								else
								{
									HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCTableInfo.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
									if (hTKKCMCCellInfo != null)
									{
										GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + j, num2, sheetView2, hTKKCMCCellInfo.CellTemplate.RowIndex, hTKKCMCCellInfo.CellTemplate.ColIndex);
										FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + j, num2, sheetView, col, num2, col2, col3);
									}
								}
								if (!string.IsNullOrEmpty(value))
								{
									ChangeCellTypeHoachToan(dataLocationTemplateRowStart + j, num2);
									SetLoaiTienStatic(sheetView2, hTKKCMCSheetInfo2, itemMapMct, "");
								}
							}
						}
						else
						{
							HTKKCMCTableInfo hTKKCMCTableInfo2 = WinFnProcess.ReadSheet_LanTruoc(sheetIndex)?.Tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == tablePath);
							if (hTKKCMCTableInfo2 != null && hTKKCMCTableInfo2.RowsValues.Count > 0)
							{
								HTKKCMCTableCellInfo hTKKCMCTableCellInfo2 = hTKKCMCTableInfo2.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == MCTPath);
								if (hTKKCMCTableCellInfo2 != null)
								{
									GeneralFunctions.FormatCol4567KHBSBySinfo(hTKKCMCTableCellInfo2, sheetView, dataLocationTemplateRowStart + j, col, num2, col2, col3);
								}
								else
								{
									HTKKCMCCellInfo fromCell = hTKKCMCTableInfo2.Header.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
									GeneralFunctions.FormatCol4567KHBSBySinfoStatic(fromCell, sheetView, dataLocationTemplateRowStart + j, col, num2, col2, col3);
								}
							}
						}
					}
					else if (hTKKCMCSheetInfo2 != null && sheetView2 != null)
					{
						HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo2.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
						GeneralFunctions.UpsertCellType(sheetIndex_KHBS, dataLocationTemplateRowStart + j, num2, sheetView2, hTKKCMCCellInfo2.CellTemplate.RowIndex, hTKKCMCCellInfo2.CellTemplate.ColIndex);
						FormatCol4567(sheetIndex_KHBS, dataLocationTemplateRowStart + j, num2, sheetView, col, num2, col2, col3);
						if (!string.IsNullOrEmpty(value))
						{
							ChangeCellTypeHoachToan(dataLocationTemplateRowStart + j, num2);
							SetLoaiTienStatic(sheetView2, WinSettings.ToKhaiInfo.SheetInfos[sheetIndex], itemMapMct, "");
						}
					}
					else
					{
						HTKKCMCSheetInfo hTKKCMCSheetInfo3 = WinFnProcess.ReadSheet_LanTruoc(sheetIndex);
						if (hTKKCMCSheetInfo3 != null)
						{
							HTKKCMCCellInfo fromCell2 = hTKKCMCSheetInfo3.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == MCTPath);
							GeneralFunctions.FormatCol4567KHBSBySinfoStatic(fromCell2, sheetView, dataLocationTemplateRowStart + j, col, num2, col2, col3);
						}
					}
					sheetView.Cells[dataLocationTemplateRowStart + j - 1, num3 - 1].Value = num5;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ChangeCellTypeNhieuLoaiTien(SheetView sView, HTKKCMCSheetInfo sInfo)
	{
		try
		{
			if (!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				return;
			}
			foreach (HTKKCMCTableInfo table in sInfo.Tables)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				HTKKCMCTableCellInfo hTKKCMCTableCellInfo = table.RowsInfo.Cells.FirstOrDefault((HTKKCMCTableCellInfo x) => x.Path == "ma_DonViTien");
				if (hTKKCMCTableCellInfo == null)
				{
					continue;
				}
				for (int num = 0; num < table.RowsValues.Count; num++)
				{
					object value = sView.Cells[dataLocationTemplateRowStart + num - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value;
					if (value == null)
					{
						continue;
					}
					if (value.ToString() == "VND")
					{
						foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
						{
							if (cell.Controltype == "16" || (table.TableName == "KHBS_01_1_TABLE_1" && (cell.Path == "ct04" || cell.Path == "ct05")))
							{
								object value2 = sView.Cells[dataLocationTemplateRowStart + num - 1, cell.CellTemplate.ColIndex - 1].Value;
								if (value2 != null)
								{
									GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + num, cell.CellTemplate.ColIndex, value2.ToString(), cell.Validate, "0");
								}
							}
						}
						continue;
					}
					foreach (HTKKCMCTableCellInfo cell2 in table.RowsInfo.Cells)
					{
						if (cell2.Controltype == "16" || (table.TableName == "KHBS_01_1_TABLE_1" && (cell2.Path == "ct04" || cell2.Path == "ct05")))
						{
							object value2 = sView.Cells[dataLocationTemplateRowStart + num - 1, cell2.CellTemplate.ColIndex - 1].Value;
							if (value2 != null)
							{
								GeneralFunctions.InsertNumberCell(sView, dataLocationTemplateRowStart + num, cell2.CellTemplate.ColIndex, value2.ToString(), cell2.Validate, "0,00");
							}
						}
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void ChangeCellTypeHoachToan(int rowMCT_01_1_KHBS, int colMCT_01_1_KHBS)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			int num2 = HTKKFile.ColumnLettersToIndex("O");
			int num3 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView2 = winFpSpread.Sheets[num3];
			HTKKCMCSheetInfo hTKKCMCSheetInfo2 = WinSettings.ToKhaiInfo.SheetInfos[num3];
			if (!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				string pathCt07_01_1_KHBS = "PL01_1_KHBS/Muc_A/DSachHSo/TongCong/tongCong_7";
				string pathCt08_01_1_KHBS = "PL01_1_KHBS/Muc_A/DSachHSo/TongCong/tongCong_8";
				string pathCt09_01_1_KHBS = "PL01_1_KHBS/Muc_A/DSachHSo/TongCong/tongCong_9";
				string pathCt10_01_KHBS = "PL01_KHBS/Muc_A/Muc_I/Muc_1/tongCong_ct10";
				string pathCt11_01_KHBS = "PL01_KHBS/Muc_A/Muc_I/Muc_2/DSachPLuc/tongCong_ct11";
				string pathCt12_01_KHBS = "PL01_KHBS/Muc_A/Muc_II/tongCong_ct12";
				HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt07_01_1_KHBS));
				HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt08_01_1_KHBS));
				HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt09_01_1_KHBS));
				HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt10_01_KHBS));
				HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt11_01_KHBS));
				HTKKCMCCellInfo hTKKCMCCellInfo6 = hTKKCMCSheetInfo2.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathCt12_01_KHBS));
				ICellType cellType = sheetView.Cells[rowMCT_01_1_KHBS - 1, colMCT_01_1_KHBS - 1].CellType;
				sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].CellType = cellType;
				sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].CellType = cellType;
				sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].CellType = cellType;
				sheetView2.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].CellType = cellType;
				sheetView2.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].CellType = cellType;
				sheetView2.Cells[hTKKCMCCellInfo6.CellTemplate.RowIndex - 1, hTKKCMCCellInfo6.CellTemplate.ColIndex - 1].CellType = cellType;
				List<HTKKCMCTableInfo> tables = WinSettings.ToKhaiInfo.SheetInfos[num3].Tables;
				HTKKCMCTableInfo hTKKCMCTableInfo = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_I/Muc_1/ChiTiet");
				HTKKCMCTableInfo hTKKCMCTableInfo2 = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_II/ChiTiet");
				HTKKCMCTableInfo hTKKCMCTableInfo3 = tables.FirstOrDefault((HTKKCMCTableInfo x) => x.TablePath == "PL01_KHBS/Muc_A/Muc_III/ChiTiet");
				int dataLocationTemplateRowStart = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
				for (int num4 = 0; num4 < hTKKCMCTableInfo.RowsValues.Count; num4++)
				{
					sheetView2.Cells[dataLocationTemplateRowStart + num4 - 1, num2 - 1].CellType = cellType;
				}
				int dataLocationTemplateRowStart2 = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				for (int num4 = 0; num4 < hTKKCMCTableInfo2.RowsValues.Count; num4++)
				{
					sheetView2.Cells[dataLocationTemplateRowStart2 + num4 - 1, num2 - 1].CellType = cellType;
				}
				int dataLocationTemplateRowStart3 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
				for (int num4 = 0; num4 < hTKKCMCTableInfo3.RowsValues.Count; num4++)
				{
					sheetView2.Cells[dataLocationTemplateRowStart3 + num4 - 1, num2 - 1].CellType = cellType;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SetDefaultLoaiTienStatic()
	{
		try
		{
			if (_mapMct == null)
			{
				return;
			}
			string text = ((_mapMct.Attributes["LoaiTienPath"] != null) ? ((XmlNode)_mapMct.Attributes["LoaiTienPath"]).Value.Trim() : "");
			if (string.IsNullOrEmpty(text))
			{
				return;
			}
			List<HTKKCMCSheetInfo> sheetInfos = WinSettings.ToKhaiInfo.SheetInfos;
			for (int i = 0; i < sheetInfos.Count; i++)
			{
				if (sheetInfos[i] != null && !sheetInfos[i].SheetName.Contains("KHBS"))
				{
					SetLoaiTienStatic(winFpSpread.Sheets[i], sheetInfos[i], null, text);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool SetLoaiTienStatic(SheetView sViewToDieuChinh, HTKKCMCSheetInfo sInfoToDieuChinh, XmlNode mapCt, string pathLoaiTienMap)
	{
		try
		{
			if (!"I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				if (mapCt == null && string.IsNullOrEmpty(pathLoaiTienMap))
				{
					return false;
				}
				if (string.IsNullOrEmpty(pathLoaiTienMap))
				{
					pathLoaiTienMap = ((mapCt.Attributes["LoaiTienPath"] != null) ? ((XmlNode)mapCt.Attributes["LoaiTienPath"]).Value : "");
				}
				if (string.IsNullOrEmpty(pathLoaiTienMap))
				{
					return false;
				}
				HTKKCMCCellInfo hTKKCMCCellInfo = sInfoToDieuChinh.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals(pathLoaiTienMap));
				if (hTKKCMCCellInfo == null)
				{
					return false;
				}
				string text = sViewToDieuChinh.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				int[] array = GeneralFunctions.ArrayCellId("BQ_39");
				int[] array2 = GeneralFunctions.ArrayCellId("AA_37");
				int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 1;
				SheetView sheetView = winFpSpread.Sheets[num];
				int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
				SheetView sheetView2 = winFpSpread.Sheets[num2];
				if (!string.IsNullOrEmpty(text))
				{
					sheetView.Cells[array[1], array[0]].Text = text;
					sheetView2.Cells[array2[1], array2[0]].Text = text;
				}
			}
			return true;
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
			return false;
		}
	}

	public void AnHienSoThueThuHoiHoanTienNopCham_01_KHBS()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			if (hTKKCMCSheetInfo.SheetName != "01-KHBS")
			{
				return;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_III/tongCong_ct13"));
			if (hTKKCMCCellInfo == null)
			{
				return;
			}
			SheetView sheetView = winFpSpread.Sheets[num];
			double num2 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			foreach (HTKKCMCCellInfo staticCell in hTKKCMCSheetInfo.StaticCells)
			{
				if (staticCell.CellTemplate.RowIndex > hTKKCMCCellInfo.CellTemplate.RowIndex)
				{
					GeneralFunctions.LockUnLockCol(sheetView, staticCell.CellTemplate.RowIndex, staticCell.CellTemplate.ColIndex, num2 == 0.0, num2 == 0.0);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RevertPhanBo_01_KHBS()
	{
		try
		{
			int num = HTKKFile.ColumnLettersToIndex("E");
			int num2 = HTKKFile.ColumnLettersToIndex("BA");
			int num3 = HTKKFile.ColumnLettersToIndex("BC");
			int num4 = HTKKFile.ColumnLettersToIndex("BD");
			int num5 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView = winFpSpread.Sheets[num5];
			IEnumerable<HTKKCMCTableInfo> enumerable = WinSettings.ToKhaiInfo.SheetInfos[num5].Tables.Where((HTKKCMCTableInfo x) => x.TableName == "KHBS_01_TABLE_2");
			foreach (HTKKCMCTableInfo item in enumerable)
			{
				try
				{
					int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
					for (int num6 = 0; num6 < item.RowsValues.Count; num6++)
					{
						object value = sheetView.Cells[dataLocationTemplateRowStart + num6 - 1, num - 1].Value;
						string text = sheetView.Cells[dataLocationTemplateRowStart + num6 - 1, num3 - 1].Text.ToString().Trim();
						string text2 = sheetView.Cells[dataLocationTemplateRowStart + num6 - 1, num2 - 1].Text.ToString().Trim();
						if (value == null || string.IsNullOrEmpty(value.ToString().Trim()))
						{
							string text3 = "";
							string text4 = "";
							List<string> list = GeneralFunctions.SplitString(text, " - ");
							if (list.Count > 1)
							{
								text3 = list[0];
								List<string> list2 = new List<string>();
								list2.Add(text3);
								list2.Add(text2);
								list2.Add(list[list.Count - 1]);
								text4 = GeneralFunctions.ListToString(list2, " - ");
							}
							else
							{
								text4 = GeneralFunctions.FormatKey2(text2, text, " - ");
							}
							sheetView.Cells[dataLocationTemplateRowStart + num6 - 1, num - 1].Value = text4;
							sheetView.Cells[dataLocationTemplateRowStart + num6 - 1, num4 - 1].Value = text3;
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
				WinSpreadProcess.fpSpReindexingTableTemplate(sheetView, item);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void Validate_01_KHBS()
	{
		try
		{
			if (!(WinSettings.ToKhaiInfo.KieuKHBS == "3") || !(WinSettings.TTPeriodInfo.LoaiToKhai == "B"))
			{
				return;
			}
			int[] array = GeneralFunctions.ArrayCellId("L_8");
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			SheetView sheetView = winFpSpread.Sheets[num];
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			int num2 = 0;
			XMLMessage xMLMessage = new XMLMessage("../Project/Message.xml");
			string message = xMLMessage.GetMessage("0596");
			object value = sheetView.Cells[array[1], array[0]].Value;
			if (value == null || string.IsNullOrEmpty(value.ToString().Trim()))
			{
				HTKKFile.SetNodeBackGroundYellow(sheetView, array[1], array[0], message);
				num2 = 1;
			}
			else
			{
				HTKKFile.RemoveValidate(sheetView, array[1], array[0]);
			}
			if (CheckB_I_PL_01_KHBS(sheetView, hTKKCMCSheetInfo, message))
			{
				num2 = 2;
			}
			string message2 = xMLMessage.GetMessage("0632");
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_2_ngay");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_3_ngay");
			string text = sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
			string text2 = sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Text.Trim();
			if (!string.IsNullOrEmpty(text) && !string.IsNullOrEmpty(text2))
			{
				DateTime t = DateTime.ParseExact(text, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				DateTime t2 = DateTime.ParseExact(text2, "dd/MM/yyyy", CultureInfo.CurrentCulture);
				if (DateTime.Compare(t, t2) > 0)
				{
					HTKKFile.SetNodeBackGroundPink(sheetView, hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1, message2);
					num2 = 2;
				}
			}
			CheckCotHuyenPhanBo_PL_01_KHBS(sheetView, _mapMct);
			switch (num2)
			{
			case 2:
				winFpSpread.TabStrip[num].BackColor = Color.Pink;
				break;
			case 1:
				winFpSpread.TabStrip[num].BackColor = Color.Yellow;
				break;
			default:
				winFpSpread.TabStrip[num].BackColor = Color.White;
				break;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RemoveTableF7IfNull(SheetView sView, HTKKCMCSheetInfo sInfo, List<HTKKCMCTableInfo> tables)
	{
		try
		{
			if ("I7,I4,C7,I8,C8".Contains(WinSettings.MenuID))
			{
				return;
			}
			List<int> list = new List<int>();
			NumberCellType numberCellType = new NumberCellType();
			CurrencyCellType currencyCellType = new CurrencyCellType();
			for (int i = 0; i < tables.Count; i++)
			{
				int num = 0;
				try
				{
					int dataLocationTemplateRowStart = tables[i].Location.DataLocationTemplateRowStart;
					for (int j = 0; j < tables[i].RowsInfo.Cells.Count; j++)
					{
						try
						{
							HTKKCMCTableCellInfo hTKKCMCTableCellInfo = tables[i].RowsInfo.Cells[j];
							if (hTKKCMCTableCellInfo == null || ((!(tables[i].TableName == "KHBS_01_1_TABLE_1") || j < 1 || j > 7) && (!(tables[i].TableName == "KHBS_01_TABLE_2") || j < 1 || j > 11)))
							{
								continue;
							}
							ICellType cellType = sView.Cells[dataLocationTemplateRowStart - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].CellType;
							object value = sView.Cells[dataLocationTemplateRowStart - 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex - 1].Value;
							if ((cellType != null && (object)cellType.GetType() == numberCellType.GetType()) || (cellType != null && (object)cellType.GetType() == currencyCellType.GetType()))
							{
								if ((!string.IsNullOrEmpty(hTKKCMCTableCellInfo.DefaultValue) || (value != null && (string.IsNullOrEmpty(value.ToString()) || Convert.ToDouble(value) != 0.0))) && ((!string.IsNullOrEmpty(hTKKCMCTableCellInfo.DefaultValue) && Convert.ToDouble(value) != Convert.ToDouble(hTKKCMCTableCellInfo.DefaultValue)) || (!string.IsNullOrEmpty(hTKKCMCTableCellInfo.DefaultValue) && Convert.ToDouble(value) != Convert.ToDouble(hTKKCMCTableCellInfo.DefaultValue))))
								{
									num++;
								}
							}
							else if (value != null && value.ToString() != hTKKCMCTableCellInfo.DefaultValue)
							{
								num++;
							}
						}
						catch (Exception ex)
						{
							HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
						}
					}
				}
				catch (Exception ex)
				{
					HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
				}
				if (num == 0)
				{
					list.Add(i);
				}
			}
			if (list.Count == 0)
			{
				return;
			}
			for (int i = list.Count - 1; i >= 0; i--)
			{
				int num2 = list[i];
				HTKKCMCTableInfo hTKKCMCTableInfo = tables[num2];
				if (num2 == 0 && hTKKCMCTableInfo.TableID == "1" && ((tables[num2].TableName == "KHBS_01_1_TABLE_1" && sInfo.Tables.Count == 2) || (tables[num2].TableName == "KHBS_01_TABLE_2" && sInfo.Tables.Count == 4)))
				{
					RemoveAllValueToDefault(sView, hTKKCMCTableInfo);
				}
				else
				{
					num2 = ((tables[list[i]].TableName == "KHBS_01_1_TABLE_1") ? num2 : (num2 + 1));
					int dataLocationTemplateRowStart = tables[list[i]].Location.TableTemplateRowStart;
					GeneralFunctions.DeleteTableF8(sView, sInfo, num2, dataLocationTemplateRowStart - 1, 0, batBuocRemove: true);
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void RemoveAllValueToDefault(SheetView sView, HTKKCMCTableInfo table)
	{
		try
		{
			if (table.RowsValues.Count > 0)
			{
				int dataLocationTemplateRowStart = table.Location.DataLocationTemplateRowStart;
				foreach (HTKKCMCTableCellInfo cell in table.RowsInfo.Cells)
				{
					sView.Cells[dataLocationTemplateRowStart - 1, cell.CellTemplate.ColIndex - 1].Value = cell.DefaultValue;
				}
			}
			foreach (HTKKCMCCellInfo staticCell in table.Header.StaticCells)
			{
				sView.Cells[staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1].Value = staticCell.DefaultValue;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public bool CheckB_I_PL_01_KHBS(SheetView sView, HTKKCMCSheetInfo sInfo, string errorBatBuocNhap)
	{
		bool flag = false;
		try
		{
			HTKKCMCCellInfo hTKKCMCCellInfo = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_1");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = sInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_3_ngay");
			foreach (HTKKCMCCellInfo staticCell in sInfo.StaticCells)
			{
				if (staticCell.CellTemplate.RowIndex >= hTKKCMCCellInfo.CellTemplate.RowIndex && staticCell.CellTemplate.RowIndex <= hTKKCMCCellInfo2.CellTemplate.RowIndex)
				{
					HTKKFile.RemoveValidate(sView, staticCell.CellTemplate.RowIndex - 1, staticCell.CellTemplate.ColIndex - 1);
				}
			}
			bool flag2 = false;
			foreach (HTKKCMCCellInfo staticCell2 in sInfo.StaticCells)
			{
				if (staticCell2.CellTemplate.RowIndex < hTKKCMCCellInfo.CellTemplate.RowIndex || staticCell2.CellTemplate.RowIndex > hTKKCMCCellInfo2.CellTemplate.RowIndex)
				{
					continue;
				}
				FarPoint.Win.Spread.Cell cell = sView.Cells[staticCell2.CellTemplate.RowIndex - 1, staticCell2.CellTemplate.ColIndex - 1];
				if (staticCell2.Controltype == "16")
				{
					if (staticCell2 != null && ((!string.IsNullOrEmpty(staticCell2.DefaultValue) && Convert.ToDouble(cell.Value) != Convert.ToDouble(staticCell2.DefaultValue)) || (string.IsNullOrEmpty(staticCell2.DefaultValue) && Convert.ToDouble(cell.Value) != 0.0)))
					{
						flag2 = true;
						break;
					}
				}
				else if (staticCell2.Controltype != "16" && staticCell2 != null && cell.Text.Trim() != staticCell2.DefaultValue)
				{
					flag2 = true;
					break;
				}
			}
			foreach (HTKKCMCCellInfo staticCell3 in sInfo.StaticCells)
			{
				if (!flag2 || staticCell3.CellTemplate.RowIndex < hTKKCMCCellInfo.CellTemplate.RowIndex || staticCell3.CellTemplate.RowIndex > hTKKCMCCellInfo2.CellTemplate.RowIndex)
				{
					continue;
				}
				FarPoint.Win.Spread.Cell cell = sView.Cells[staticCell3.CellTemplate.RowIndex - 1, staticCell3.CellTemplate.ColIndex - 1];
				if (staticCell3.Controltype == "16")
				{
					if (staticCell3 != null && Convert.ToDouble(cell.Value) == 0.0)
					{
						HTKKFile.SetNodeBackGroundPink(sView, staticCell3.CellTemplate.RowIndex - 1, staticCell3.CellTemplate.ColIndex - 1, errorBatBuocNhap);
						flag = true;
					}
				}
				else if (staticCell3.Controltype != "16" && staticCell3 != null && string.IsNullOrEmpty(cell.Text.Trim()))
				{
					HTKKFile.SetNodeBackGroundPink(sView, staticCell3.CellTemplate.RowIndex - 1, staticCell3.CellTemplate.ColIndex - 1, errorBatBuocNhap);
					flag = true;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
		return flag;
	}

	public void SuaCotHuyenPhanBo_PL_01_KHBS(SheetView sView_01_KHBS, XmlNode mapCt)
	{
		try
		{
			string text = ((_mapMct != null && _mapMct.Attributes["SuaPhanBo"] != null) ? ((XmlNode)_mapMct.Attributes["SuaPhanBo"]).Value : "");
			if (!(text == "1") || !(WinSettings.ToKhaiInfo.KieuKHBS == "3"))
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("P");
			int num2 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "V" : "T");
			int num3 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "BE" : "AP");
			int num4 = HTKKFile.ColumnLettersToIndex("AO");
			int num5 = HTKKFile.ColumnLettersToIndex("AQ");
			int num6 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num6];
			IEnumerable<HTKKCMCTableInfo> enumerable = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == "KHBS_01_TABLE_2");
			foreach (HTKKCMCTableInfo item in enumerable)
			{
				int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
				for (int num7 = 0; num7 < item.RowsValues.Count; num7++)
				{
					HTKKFile.RemoveValidate(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num - 1);
					HTKKFile.RemoveValidate(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num2 - 1);
					string text2 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num5 - 1].Text.Trim();
					string text3 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Text.Trim();
					if (text2 == "82500")
					{
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num, locked: false, deleteData: false);
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, locked: false, deleteData: false);
					}
					else if (string.IsNullOrEmpty(text3) || text3 == "0")
					{
						string maTinh = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num4 - 1].Text.Trim();
						if (!string.IsNullOrEmpty(maTinh) && maTinh != "0")
						{
							Dictionary<string, string> dictionary = HTKKWinDictionaryDiaDanh.dicQuanHuyen.Where((KeyValuePair<string, string> x) => x.Key.Substring(0, 3) == maTinh).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
							if (dictionary != null && dictionary.Count() > 0)
							{
								GeneralFunctions.InsertCombobox(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, dictionary, deleteText: false);
								GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num, locked: true, deleteData: false);
								GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, locked: false, deleteData: false);
							}
						}
						else if (!string.IsNullOrEmpty(text2))
						{
							GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num, locked: false, deleteData: false);
							GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, locked: false, deleteData: false);
						}
						else
						{
							GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num, locked: true, deleteData: true);
							GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, locked: true, deleteData: true);
						}
					}
					else
					{
						string maTinh2 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num4 - 1].Text.Trim();
						Dictionary<string, string> dictionary = HTKKWinDictionaryDiaDanh.dicQuanHuyen.Where((KeyValuePair<string, string> x) => x.Key.Substring(0, 3) == maTinh2).ToDictionary((KeyValuePair<string, string> x) => x.Key, (KeyValuePair<string, string> x) => x.Value);
						GeneralFunctions.InsertCombobox(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, dictionary, deleteText: false);
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num, locked: true, deleteData: false);
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, locked: true, deleteData: false);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void SuaCot7PhanBo_PL_01_KHBS(SheetView sView_01_KHBS, XmlNode mapCt)
	{
		try
		{
			string text = ((_mapMct != null && _mapMct.Attributes["SuaPhanBo"] != null) ? ((XmlNode)_mapMct.Attributes["SuaPhanBo"]).Value : "");
			if (!(text == "1") || !(WinSettings.ToKhaiInfo.KieuKHBS == "3"))
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("P");
			int colHide = HTKKFile.ColumnLettersToIndex("AF");
			int num2 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num2];
			IEnumerable<HTKKCMCTableInfo> enumerable = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == "KHBS_01_TABLE_2");
			foreach (HTKKCMCTableInfo item in enumerable)
			{
				int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
				for (int num3 = 0; num3 < item.RowsValues.Count; num3++)
				{
					string value = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num3 - 1, num - 1].Text.Trim();
					if (string.IsNullOrEmpty(value))
					{
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num3, colHide, locked: true, deleteData: false);
					}
					else
					{
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num3, colHide, locked: false, deleteData: false);
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void CheckCotHuyenPhanBo_PL_01_KHBS(SheetView sView_01_KHBS, XmlNode mapCt)
	{
		try
		{
			string text = ((_mapMct != null && _mapMct.Attributes["SuaPhanBo"] != null) ? ((XmlNode)_mapMct.Attributes["SuaPhanBo"]).Value : "");
			if (!(text == "1") || !(WinSettings.ToKhaiInfo.KieuKHBS == "3"))
			{
				return;
			}
			int num = HTKKFile.ColumnLettersToIndex("X");
			int num2 = HTKKFile.ColumnLettersToIndex("P");
			int num3 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "V" : "T");
			int num4 = HTKKFile.ColumnLettersToIndex("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,R5,Y6,Z8,Z5".Contains(WinSettings.MenuID) ? "BE" : "AP");
			int num5 = HTKKFile.ColumnLettersToIndex("AO");
			int num6 = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num6];
			IEnumerable<HTKKCMCTableInfo> enumerable = hTKKCMCSheetInfo.Tables.Where((HTKKCMCTableInfo x) => x.TableName == "KHBS_01_TABLE_2");
			foreach (HTKKCMCTableInfo item in enumerable)
			{
				int dataLocationTemplateRowStart = item.Location.DataLocationTemplateRowStart;
				for (int num7 = 0; num7 < item.RowsValues.Count; num7++)
				{
					HTKKFile.RemoveValidate(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num2 - 1);
					HTKKFile.RemoveValidate(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num3 - 1);
					string value = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num - 1].Text.Trim();
					if (string.IsNullOrEmpty(value))
					{
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num2, locked: true, deleteData: false);
						GeneralFunctions.LockUnLockCol(sView_01_KHBS, dataLocationTemplateRowStart + num7, num3, locked: true, deleteData: false);
						continue;
					}
					string value2 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num2 - 1].Text.Trim();
					string value3 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Text.Trim();
					string value4 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num5 - 1].Text.Trim();
					string value5 = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num4 - 1].Text.Trim();
					bool locked = sView_01_KHBS.Cells[dataLocationTemplateRowStart + num7 - 1, num3 - 1].Locked;
					if (string.IsNullOrEmpty(value2) && !locked)
					{
						HTKKFile.SetNodeBackGroundPink(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num2 - 1, HTKKFile.GetMessage("0596"));
						sView_01_KHBS.FpSpread.TabStrip[num6].BackColor = Color.Pink;
						WinSettings.ToKhaiInfo.SheetInfos[num6].isValidData = false;
					}
					if (string.IsNullOrEmpty(value3) && !locked)
					{
						HTKKFile.SetNodeBackGroundPink(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num3 - 1, HTKKFile.GetMessage("0596"));
						sView_01_KHBS.FpSpread.TabStrip[num6].BackColor = Color.Pink;
						WinSettings.ToKhaiInfo.SheetInfos[num6].isValidData = false;
					}
					if (!string.IsNullOrEmpty(value2) && string.IsNullOrEmpty(value4) && !locked)
					{
						HTKKFile.SetNodeBackGroundPink(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num2 - 1, HTKKFile.GetMessage("0629"));
						sView_01_KHBS.FpSpread.TabStrip[num6].BackColor = Color.Pink;
						WinSettings.ToKhaiInfo.SheetInfos[num6].isValidData = false;
					}
					if (!string.IsNullOrEmpty(value3) && string.IsNullOrEmpty(value5) && !locked)
					{
						HTKKFile.SetNodeBackGroundPink(sView_01_KHBS, dataLocationTemplateRowStart + num7 - 1, num3 - 1, HTKKFile.GetMessage("0629"));
						sView_01_KHBS.FpSpread.TabStrip[num6].BackColor = Color.Pink;
						WinSettings.ToKhaiInfo.SheetInfos[num6].isValidData = false;
					}
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void UpdateXmlFileMapMCTSpecial(ref XmlNode mapCt, ref XmlNode mapCTTieuMuc)
	{
		try
		{
			if (!(WinSettings.MenuID == "97"))
			{
				return;
			}
			string pathCt27 = "ct27";
			string pathCt30 = "ct30";
			string pathCt31 = "ct31";
			if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.Nam))
			{
				int num = Convert.ToInt32(WinSettings.TTPeriodInfo.Nam);
				XmlNode val = ((IEnumerable)mapCt).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == pathCt27));
				XmlNode val2 = ((IEnumerable)mapCt).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == pathCt30));
				XmlNode val3 = ((IEnumerable)mapCt).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == pathCt31));
				if (num < 2015)
				{
					((XmlNode)val.Attributes["LoaiThue"]).Value = "";
					((XmlNode)val2.Attributes["LoaiThue"]).Value = "";
					((XmlNode)val3.Attributes["LoaiThue"]).Value = "0";
				}
				else
				{
					((XmlNode)val.Attributes["LoaiThue"]).Value = "0";
					((XmlNode)val2.Attributes["LoaiThue"]).Value = "0";
					((XmlNode)val3.Attributes["LoaiThue"]).Value = "";
				}
				XmlNode val4 = ((IEnumerable)mapCTTieuMuc).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == pathCt27));
				XmlNode val5 = ((IEnumerable)mapCTTieuMuc).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == pathCt30));
				XmlNode val6 = ((IEnumerable)mapCTTieuMuc).Cast<XmlNode>().FirstOrDefault((Func<XmlNode, bool>)((XmlNode x) => x.Attributes["MCTPath"] != null && ((XmlNode)x.Attributes["MCTPath"]).Value == pathCt31));
				if (num < 2015)
				{
					((XmlNode)val4.Attributes["MaTieuMuc"]).Value = "";
					((XmlNode)val5.Attributes["MaTieuMuc"]).Value = "";
					((XmlNode)val6.Attributes["MaTieuMuc"]).Value = "1552";
					((XmlNode)val4.Attributes["TenTieuMuc"]).Value = "";
					((XmlNode)val5.Attributes["TenTieuMuc"]).Value = "";
					((XmlNode)val6.Attributes["TenTieuMuc"]).Value = "1552";
				}
				else
				{
					((XmlNode)val4.Attributes["MaTieuMuc"]).Value = "1552";
					((XmlNode)val5.Attributes["MaTieuMuc"]).Value = "1552";
					((XmlNode)val6.Attributes["MaTieuMuc"]).Value = "";
					((XmlNode)val4.Attributes["TenTieuMuc"]).Value = "1552";
					((XmlNode)val5.Attributes["TenTieuMuc"]).Value = "1552";
					((XmlNode)val6.Attributes["TenTieuMuc"]).Value = "";
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhSoNgayNhanDuocTienHoanThue_01_KHBS(SheetView sView_01_KHBS, bool is_edit)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			if (hTKKCMCSheetInfo.SheetName.ToLower().IndexOf("khbs") < 0)
			{
				return;
			}
			string ngayLapKHBS = WinSettings.TTPeriodInfo.NgayLapKHBS;
			if (string.IsNullOrEmpty(ngayLapKHBS))
			{
				return;
			}
			DateTime dateTime = DateTime.ParseExact(ngayLapKHBS, "dd/MM/yyyy", CultureInfo.CurrentCulture);
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_3_ngay");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_II/ct_1");
			if (hTKKCMCCellInfo2 == null)
			{
				return;
			}
			FarPoint.Win.Spread.Cell cell = sView_01_KHBS.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1];
			if (hTKKCMCCellInfo != null && !cell.Locked && cell.Value != null && (string.IsNullOrEmpty(cell.Value.ToString()) || Convert.ToDouble(cell.Value) == 0.0 || is_edit))
			{
				string text = sView_01_KHBS.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Text.Trim();
				if (text != null && !string.IsNullOrEmpty(text))
				{
					DateTime dateTime2 = DateTime.ParseExact(text.ToString(), "dd/MM/yyyy", CultureInfo.CurrentCulture);
					int days = (dateTime - dateTime2).Days;
					int num2 = days;
					sView_01_KHBS.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value = num2;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void TinhSoTienChamNop_01_KHBS(SheetView sView_01_KHBS, bool is_edit)
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			if (hTKKCMCSheetInfo.SheetName.ToLower().IndexOf("khbs") < 0)
			{
				return;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_I/ct_1");
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_II/ct_1");
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.FirstOrDefault((HTKKCMCCellInfo x) => x.Path == "PL01_KHBS/Muc_B/Muc_II/ct_2");
			FarPoint.Win.Spread.Cell cell = sView_01_KHBS.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1];
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo6) => hTKKCMCCellInfo6.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_1/tongCong_ct10"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo hTKKCMCCellInfo6) => hTKKCMCCellInfo6.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_2/DSachPLuc/tongCong_ct11"));
			double num2 = 0.0;
			double num3 = 0.0;
			if (hTKKCMCCellInfo4 != null)
			{
				num2 = Convert.ToDouble(sView_01_KHBS.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
			}
			if (hTKKCMCCellInfo5 != null)
			{
				num3 = Convert.ToDouble(sView_01_KHBS.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value);
			}
			if (hTKKCMCCellInfo == null || (num2 == 0.0 && num3 == 0.0) || cell.Value == null || (!string.IsNullOrEmpty(cell.Value.ToString()) && Convert.ToDouble(cell.Value) != 0.0 && !is_edit))
			{
				return;
			}
			object value = sView_01_KHBS.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value;
			if (value == null)
			{
				return;
			}
			double num4 = ((value != null && !string.IsNullOrEmpty(value.ToString())) ? Convert.ToDouble(value) : 0.0);
			object value2 = sView_01_KHBS.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value;
			if (value2 != null)
			{
				double num5 = ((value2 != null && !string.IsNullOrEmpty(value2.ToString())) ? Convert.ToDouble(value2) : 0.0);
				if (value2 != null)
				{
					double num6 = Math.Round(num4 * num5 * 0.0003, 0);
					sView_01_KHBS.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = num6;
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	public void AnHienSoNgayVaSoTienNopCham_01_KHBS()
	{
		try
		{
			int num = WinSettings.ToKhaiInfo.SheetInfos.Count - 2;
			HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[num];
			if (hTKKCMCSheetInfo.SheetName != "01-KHBS")
			{
				return;
			}
			HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_1/tongCong_ct10"));
			HTKKCMCCellInfo hTKKCMCCellInfo2 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_2/DSachPLuc/tongCong_ct11"));
			HTKKCMCCellInfo hTKKCMCCellInfo3 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_1"));
			HTKKCMCCellInfo hTKKCMCCellInfo4 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3a_2"));
			HTKKCMCCellInfo hTKKCMCCellInfo5 = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo cell) => cell.Path.Equals("PL01_KHBS/Muc_A/Muc_I/Muc_3/ct3b"));
			SheetView sheetView = winFpSpread.Sheets[num];
			double num2 = 0.0;
			double num3 = 0.0;
			if (hTKKCMCCellInfo != null)
			{
				num2 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo.CellTemplate.RowIndex - 1, hTKKCMCCellInfo.CellTemplate.ColIndex - 1].Value);
			}
			if (hTKKCMCCellInfo2 != null)
			{
				num3 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo2.CellTemplate.RowIndex - 1, hTKKCMCCellInfo2.CellTemplate.ColIndex - 1].Value);
			}
			double num4 = 0.0;
			double num5 = 0.0;
			if (hTKKCMCCellInfo4 != null)
			{
				num4 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value);
			}
			if (hTKKCMCCellInfo5 != null)
			{
				num5 = Convert.ToDouble(sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value);
			}
			bool flag = ((num2 == 0.0 && num3 == 0.0) ? true : false);
			if (num2 == 0.0 && num3 == 0.0)
			{
				if (hTKKCMCCellInfo3 != null)
				{
					sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = "";
				}
				if (hTKKCMCCellInfo4 != null)
				{
					sheetView.Cells[hTKKCMCCellInfo4.CellTemplate.RowIndex - 1, hTKKCMCCellInfo4.CellTemplate.ColIndex - 1].Value = 0;
				}
				if (hTKKCMCCellInfo5 != null)
				{
					sheetView.Cells[hTKKCMCCellInfo5.CellTemplate.RowIndex - 1, hTKKCMCCellInfo5.CellTemplate.ColIndex - 1].Value = 0;
				}
			}
			else if (!string.IsNullOrEmpty(WinSettings.TTPeriodInfo.NgayLapKHBS))
			{
				sheetView.Cells[hTKKCMCCellInfo3.CellTemplate.RowIndex - 1, hTKKCMCCellInfo3.CellTemplate.ColIndex - 1].Value = WinSettings.TTPeriodInfo.NgayLapKHBS;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void TongHopKHBS(HTKKCMCTableInfo table, ref int iCountDieuChinh, double BS2, double BS1, HTKKCMCSheetInfo sInfo, int iActiveSheet, int iChiTieu, int tableIdx)
	{
		try
		{
			iCountDieuChinh++;
			HTKKCMCRowValue hTKKCMCRowValue = new HTKKCMCRowValue();
			hTKKCMCRowValue.RowValue.Add(WinSettings.KHBS[iChiTieu].TenChiTieu.ToString());
			hTKKCMCRowValue.RowValue.Add(WinSettings.KHBS[iChiTieu].MaChiTieu.ToString());
			hTKKCMCRowValue.RowValue.Add(iCountDieuChinh.ToString());
			hTKKCMCRowValue.RowValue.Add(BS1.ToString());
			hTKKCMCRowValue.RowValue.Add(BS2.ToString());
			hTKKCMCRowValue.RowValue.Add((BS2 - BS1).ToString());
			if (!table.RowsValues.Contains(hTKKCMCRowValue))
			{
				table.RowsValues.Add(hTKKCMCRowValue);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmWinToKhai_Resize(object sender, EventArgs e)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Expected O, but got Unknown
		//IL_0127: Unknown result type (might be due to invalid IL or missing references)
		//IL_012e: Expected O, but got Unknown
		int height = ((Control)this).Height;
		((Control)pnInsert).Location = new Point(((Control)pnWinSpread).Location.X, ((Control)pnButon).Location.Y);
		try
		{
			if (!File.Exists("Config.xml"))
			{
				return;
			}
			XmlDocument val = new XmlDocument();
			val.LoadXml(HTKKFile.LoadFile("Config.xml"));
			XmlNode val2 = ((XmlNode)val).SelectNodes("/Root/SizeAPP")[0];
			if (val2 != null)
			{
				string text = ((Control)this).Width.ToString();
				string text2 = ((Control)this).Height.ToString();
				val2.InnerText = ((Form)this).Location.X + "~" + ((Form)this).Location.Y + "~" + text + "~" + text2;
				XmlTextWriter val3 = new XmlTextWriter("Config.xml", Encoding.UTF8);
				try
				{
					val3.Formatting = (Formatting)1;
					val.Save((XmlWriter)(object)val3);
					return;
				}
				finally
				{
					((IDisposable)val3)?.Dispose();
				}
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_MouseHover(object sender, EventArgs e)
	{
		if (winFpSpread.IsLayoutSuspended)
		{
			winFpSpread.ResumeLayout();
		}
	}

	private void winFpSpread_MouseWheel(object sender, MouseEventArgs e)
	{
		winFpSpread.SuspendLayout();
		ac = winFpSpread.ActiveSheet.ActiveCell;
		scrollCurrent = ac.Row.Index;
		winFpSpread.ActiveSheet.SetActiveCell(scrollCurrent, 0);
		winFpSpread.ResumeLayout();
	}

	private void frmWinToKhai_FormClosed(object sender, FormClosedEventArgs e)
	{
		try
		{
			if (WinSettings.ToKhaiInfo.FolderName == "HEADER")
			{
				if (HTKKFile.IsEmptyFolder(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst))
				{
					HTKKFile.Delete_Folder(HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst);
					frmFormStart frmFormStart = new frmFormStart();
					((Control)frmFormStart).Show();
					return;
				}
				HTKKFactory.GetInstanceHTKKHeader().SetPropertyHeader(ref WinSettings.HTKKHeaderInfo, WinSettings.HTKKHeaderInfo.mst);
				if (WinSettings.myHTKKSystems != null)
				{
					((ToolStripItem)WinSettings.myHTKKSystems.toolStripStatusMST).Text = WinSettings.HTKKHeaderInfo.mst + " : " + WinSettings.HTKKHeaderInfo.tenNNT.Replace("&", "&&");
				}
			}
			WinSettings.ToKhaiInfo = new HTKKCMCToKhaiInfo();
			WinSettings.TTPeriodInfo = new HTKKPeriodInfo();
			FormProcess.ReLoadfrmSystems();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_CellClick(object sender, CellClickEventArgs e)
	{
		try
		{
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
			{
				return;
			}
			string text = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
			if (WinSettings.MenuID == "33" && winFpSpread.ActiveSheet.SheetName == "05-1BK-QTT-TNCN" && e.Column >= HTKKFile.ColumnLettersToIndex("T") - 1 && e.Row > 20)
			{
				object value = winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("S") - 1].Value;
				if (ConvertHelpers.ConvertToBool(value))
				{
					winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = false;
					winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = false;
					winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = false;
				}
				else
				{
					winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("T") - 1].Locked = true;
					winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("U") - 1].Locked = true;
					winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex("V") - 1].Locked = true;
				}
			}
			if (winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Locked)
			{
				return;
			}
			winFpSpread.ActiveSheet.Cells[e.Row, e.Column].ResetFormula();
			MaskCellType maskCellType = new MaskCellType();
			maskCellType.Mask = "00/00/0000";
			maskCellType.MaskChar = '.';
			HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
			if (staticCell != null)
			{
				if (staticCell.Controltype == "14")
				{
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
				}
				if (staticCell.Controltype == "104")
				{
					maskCellType.Mask = "00/00";
					maskCellType.MaskChar = '.';
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
				}
				if (staticCell.Controltype == "109")
				{
					maskCellType.Mask = "0000";
					maskCellType.MaskChar = '.';
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
				}
				if (staticCell.Controltype.IndexOf("_M_") >= 0 && winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType is CheckedComboBoxCellType checkedComboBoxCellType)
				{
					checkedComboBoxCellType.ShowMultiCheckDropdown(winFpSpread.ActiveSheet, e.Row, e.Column, staticCell.Controltype, staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1);
				}
			}
			else
			{
				int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(e.Row + 1, e.Column + 1);
				if (tableIndex >= 0)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
					staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Header.StaticCells);
					if (staticCell != null)
					{
						if (staticCell.Controltype == "14")
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
						}
						if (staticCell.Controltype == "104")
						{
							maskCellType.Mask = "00/00";
							maskCellType.MaskChar = '.';
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
						}
						if (staticCell.Controltype == "109")
						{
							maskCellType.Mask = "0000";
							maskCellType.MaskChar = '.';
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
						}
						if (staticCell.Controltype.IndexOf("_M_") >= 0 && winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType is CheckedComboBoxCellType checkedComboBoxCellType2)
						{
							checkedComboBoxCellType2.ShowMultiCheckDropdown(winFpSpread.ActiveSheet, e.Row, e.Column, staticCell.Controltype, staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1);
						}
					}
					else
					{
						staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Footer.StaticCells);
						if (staticCell != null)
						{
							if (staticCell.Controltype == "14")
							{
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
							}
							if (staticCell.Controltype == "104")
							{
								maskCellType.Mask = "00/00";
								maskCellType.MaskChar = '.';
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
							}
							if (staticCell.Controltype == "109")
							{
								maskCellType.Mask = "0000";
								maskCellType.MaskChar = '.';
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
							}
							if (staticCell.Controltype.IndexOf("_M_") >= 0 && winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType is CheckedComboBoxCellType checkedComboBoxCellType3)
							{
								checkedComboBoxCellType3.ShowMultiCheckDropdown(winFpSpread.ActiveSheet, e.Row, e.Column, staticCell.Controltype, staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1);
							}
						}
						if ((e.Row + 1 >= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart) & (e.Row + 1 <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count))
						{
							for (int i = 0; i < hTKKCMCTableInfo.RowsInfo.Cells.Count; i++)
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[i].CellTemplate.ColIndex == e.Column + 1)
								{
									if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "14")
									{
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
									}
									if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "104")
									{
										maskCellType.Mask = "00/00";
										maskCellType.MaskChar = '.';
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
									}
									if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "109")
									{
										maskCellType.Mask = "0000";
										maskCellType.MaskChar = '.';
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType = maskCellType;
									}
									if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype.IndexOf("_M_") >= 0 && winFpSpread.ActiveSheet.Cells[e.Row, e.Column].CellType is CheckedComboBoxCellType checkedComboBoxCellType4)
									{
										checkedComboBoxCellType4.ShowMultiCheckDropdown(winFpSpread.ActiveSheet, e.Row, e.Column, hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype, e.Row, HTKKFile.ColumnLettersToIndex(hTKKCMCTableInfo.RowsInfo.Cells[i].indexColumnSelectedValue) - 1);
									}
								}
							}
						}
					}
					try
					{
						HTKKCMCTableCellInfo tableCellInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableCellInfo(e.Row + 1, e.Column + 1);
						if (tableCellInfo.Controltype == "101" && string.IsNullOrEmpty(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text))
						{
							maTinhTP = winFpSpread.ActiveSheet.Cells[e.Row, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
							if (!string.IsNullOrEmpty(maTinhTP))
							{
								WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, e.Row + 1, e.Column + 1, HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(maTinhTP));
							}
						}
					}
					catch (Exception ex)
					{
						HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
					}
				}
			}
			if (cellOld == null)
			{
				cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
			}
			if (isErrCellOld)
			{
				winFpSpread.ActiveSheet.SetActiveCell(cellOld.Row.Index, cellOld.Column.Index);
				winFpSpread.EditModePermanent = true;
			}
			else
			{
				WinDoWorkEvent.CellClick(winFpSpread, winFpSpread.ActiveSheetIndex, e.Row + 1, e.Column + 1);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void winFpSpread_LeaveCell(object sender, LeaveCellEventArgs e)
	{
		try
		{
			if (WinSettings.ToKhaiInfo.TaiDuLieuRatLon == "1")
			{
				return;
			}
			string text = "";
			if (cellOld == null)
			{
				cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
			}
			HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
			if (staticCell != null)
			{
				if (staticCell.Controltype == "14")
				{
					text = staticCell.Value;
					string text2 = WinDoWorkEvent.CheckFormatDateOnChange(text);
					if (!HTKKFormater.isNullString(text))
					{
						if (HTKKFormater.isNullString(text2))
						{
							isErrCellOld = true;
							cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
							e.Cancel = true;
							return;
						}
						winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
						staticCell.Value = text2;
					}
				}
				if (staticCell.Controltype == "19")
				{
					string text2 = WinDoWorkEvent.CheckFormatMMyyyy(winFpSpread, winFpSpread.ActiveSheetIndex, e, new Dictionary<int, int>());
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
					if (HTKKFormater.isNullString(text2))
					{
						isErrCellOld = true;
						cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
						return;
					}
				}
				else if (staticCell.Controltype == "16" && string.IsNullOrEmpty(staticCell.Value) && string.IsNullOrEmpty(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text))
				{
					winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = (string.IsNullOrEmpty(staticCell.DefaultValue) ? "0" : staticCell.DefaultValue);
				}
			}
			else
			{
				int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(winFpSpread.ActiveSheet.ActiveRow.Index + 1, winFpSpread.ActiveSheet.ActiveColumn.Index + 1);
				if (tableIndex >= 0)
				{
					HTKKCMCTableInfo hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
					staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Header.StaticCells);
					if (staticCell != null)
					{
						if (staticCell.Controltype == "14")
						{
							text = staticCell.Value;
							string text2 = WinDoWorkEvent.CheckFormatDateOnChange(text);
							if (!HTKKFormater.isNullString(text))
							{
								if (HTKKFormater.isNullString(text2))
								{
									isErrCellOld = true;
									cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
									e.Cancel = true;
									return;
								}
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
							}
							else
							{
								winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
							}
						}
						else if (staticCell.Controltype == "16" && string.IsNullOrEmpty(staticCell.Value))
						{
							winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = (string.IsNullOrEmpty(staticCell.DefaultValue) ? "0" : staticCell.DefaultValue);
						}
					}
					else
					{
						staticCell = WinFnProcess.GetStaticCell(e.Row, e.Column, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex].Footer.StaticCells);
						if (staticCell != null)
						{
							if (staticCell.Controltype == "14")
							{
								text = staticCell.Value;
								string text2 = WinDoWorkEvent.CheckFormatDateOnChange(text);
								if (!HTKKFormater.isNullString(text))
								{
									if (HTKKFormater.isNullString(text2))
									{
										isErrCellOld = true;
										cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
										e.Cancel = true;
										return;
									}
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
								}
							}
							else if (staticCell.Controltype == "16" && string.IsNullOrEmpty(staticCell.Value))
							{
								staticCell.Value = (string.IsNullOrEmpty(staticCell.DefaultValue) ? "0" : staticCell.DefaultValue);
							}
						}
						if ((e.Row + 1 >= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart) & (e.Row + 1 <= hTKKCMCTableInfo.Location.DataLocationTemplateRowStart + hTKKCMCTableInfo.RowsValues.Count))
						{
							for (int i = 0; i < hTKKCMCTableInfo.RowsInfo.Cells.Count; i++)
							{
								if (hTKKCMCTableInfo.RowsInfo.Cells[i].CellTemplate.ColIndex != e.Column + 1)
								{
									continue;
								}
								if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "14")
								{
									text = winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text;
									string text2 = WinDoWorkEvent.CheckFormatDateOnChange(text);
									if (!HTKKFormater.isNullString(text))
									{
										if (HTKKFormater.isNullString(text2))
										{
											isErrCellOld = true;
											cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
											e.Cancel = true;
											return;
										}
										winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
									}
								}
								if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "19")
								{
									string text2 = WinDoWorkEvent.CheckFormatMMyyyy(winFpSpread, winFpSpread.ActiveSheetIndex, e, new Dictionary<int, int>());
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Text = text2;
									if (HTKKFormater.isNullString(text2))
									{
										isErrCellOld = true;
										cellOld = winFpSpread.ActiveSheet.Cells[e.Row, e.Column];
										return;
									}
								}
								else if (hTKKCMCTableInfo.RowsInfo.Cells[i].Controltype == "16" && string.IsNullOrEmpty(ConvertHelpers.ConvertToString(winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value)))
								{
									winFpSpread.ActiveSheet.Cells[e.Row, e.Column].Value = (string.IsNullOrEmpty(hTKKCMCTableInfo.RowsInfo.Cells[i].DefaultValue) ? "0" : hTKKCMCTableInfo.RowsInfo.Cells[i].DefaultValue);
								}
							}
						}
					}
				}
			}
			isErrCellOld = false;
			WinDoWorkEvent.LeaveCell(winFpSpread, winFpSpread.ActiveSheetIndex, e.Row + 1, e.Column + 1, e.NewRow + 1, e.NewColumn + 1);
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void btn_custom_Click(object sender, EventArgs e)
	{
		//IL_00be: Unknown result type (might be due to invalid IL or missing references)
		//IL_018a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0165: Unknown result type (might be due to invalid IL or missing references)
		//IL_0129: Unknown result type (might be due to invalid IL or missing references)
		//IL_0104: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			if (!(WinSettings.MenuID == "30") || !(WinSettings.TTPeriodInfo.KieuKy == "Q"))
			{
				return;
			}
			WinSettings.TTPeriodInfo.KieuKy = "Y";
			string fileNameBySheet = WinDoWorkEvent.GetFileNameBySheet(WinSettings.TTPeriodInfo, WinSettings.ToKhaiInfo.SheetInfos[0].DataFile, HTKKConfig.DATAFILES + "\\" + WinSettings.HTKKHeaderInfo.mst + "\\", WinSettings.ToKhaiInfo.VersionXML);
			if (File.Exists(fileNameBySheet))
			{
				frmMessage frmMessage = new frmMessage("0372", "5", "2", "Thông báo");
				((Form)frmMessage).ShowDialog();
				if (frmMessage.Yes)
				{
					if (WinDoWorkEvent.WinDoWorkSave(winFpSpread))
					{
						frmMessage frmMessage2 = new frmMessage("0370", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
					}
					else
					{
						frmMessage frmMessage2 = new frmMessage("0003", "3", "1", "Thông báo");
						((Form)frmMessage2).ShowDialog();
					}
				}
			}
			else if (WinDoWorkEvent.WinDoWorkSave(winFpSpread))
			{
				frmMessage frmMessage2 = new frmMessage("0370", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
			}
			else
			{
				frmMessage frmMessage2 = new frmMessage("0003", "3", "1", "Thông báo");
				((Form)frmMessage2).ShowDialog();
			}
			WinSettings.TTPeriodInfo.KieuKy = "Q";
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void SetActiveCellToObject(int sheetIndex)
	{
		WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].ActiveRowIndex = winFpSpread.Sheets[sheetIndex].ActiveCell.Row.Index;
		WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].ActiveColIndex = winFpSpread.Sheets[sheetIndex].ActiveCell.Column.Index;
	}

	private int[] GetActiveCellFromObject(int sheetIndex)
	{
		return new int[2]
		{
			WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].ActiveRowIndex,
			WinSettings.ToKhaiInfo.SheetInfos[sheetIndex].ActiveColIndex
		};
	}

	private void winFpSpread_DialogKey(object sender, DialogKeyEventArgs e)
	{
	}

	private void winFpSpread_TopChange(object sender, TopChangeEventArgs e)
	{
	}

	private void frmWinToKhai_KeyPress(object sender, KeyPressEventArgs e)
	{
		//IL_0009: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		((Control)this).KeyDown += new KeyEventHandler(frmWinToKhai_KeyDown);
	}

	private void frmWinToKhai_KeyDown(object sender, KeyEventArgs e)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Invalid comparison between Unknown and I4
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Invalid comparison between Unknown and I4
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Invalid comparison between Unknown and I4
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f3: Invalid comparison between Unknown and I4
		try
		{
			if ((e.Control && (int)e.KeyCode == 34) || (e.Control && (int)e.KeyCode == 33))
			{
				winFpSpread.EditMode = false;
			}
			if ((int)e.KeyCode == 40 && winFpSpread.ActiveSheet.ActiveCell.Row.Index > 0 && winFpSpread.ActiveSheet.ActiveCell.Row.Index < winFpSpread.ActiveSheet.NonEmptyRowCount)
			{
				winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveCell.Row.Index + 1, winFpSpread.ActiveSheet.ActiveCell.Column.Index);
			}
			if ((int)e.KeyCode == 38 && winFpSpread.ActiveSheet.ActiveCell.Row.Index > 0 && winFpSpread.ActiveSheet.ActiveCell.Row.Index < winFpSpread.ActiveSheet.NonEmptyRowCount)
			{
				winFpSpread.ActiveSheet.SetActiveCell(winFpSpread.ActiveSheet.ActiveCell.Row.Index - 1, winFpSpread.ActiveSheet.ActiveCell.Column.Index);
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void frmWinToKhai_FormClosing(object sender, FormClosingEventArgs e)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		if (isDataChanged && !WinSettings.MenuID.Equals("57") && !isDeleteEvent)
		{
			frmMessage frmMessage = new frmMessage("0013", "6", "2", "Cảnh báo");
			((Form)frmMessage).ShowDialog();
			if (frmMessage.Yes)
			{
				cmdSave_Click(sender, (EventArgs)(object)e);
				isDataChanged = false;
				isDataChangedCallTHKhbs = false;
			}
			else if (frmMessage.Cancel)
			{
				((CancelEventArgs)(object)e).Cancel = true;
				((Control)frmMessage).Hide();
			}
			else
			{
				isDataChanged = false;
				isDataChangedCallTHKhbs = false;
			}
		}
	}

	private void winFpSpread_MouseMove(object sender, MouseEventArgs e)
	{
		cellActiveWhileTab = winFpSpread.ActiveSheet.ActiveCell;
		if (WinSettings.MenuID == "A4")
		{
			if (WinSettings.TTPeriodInfo.dicVar.ContainsKey("CELL_CURRENT"))
			{
				WinSettings.TTPeriodInfo.dicVar.Remove("CELL_CURRENT");
			}
			WinSettings.TTPeriodInfo.dicVar.Add("CELL_CURRENT", cellActiveWhileTab);
		}
	}

	private void cbbCQTCuc_SelectedValueChanged(object sender, EventArgs e)
	{
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a8: Expected O, but got Unknown
		try
		{
			if (!((Control)panelCQT).Visible)
			{
				return;
			}
			string text = "";
			text = ((ListControl)cbbCQTCuc).SelectedValue.ToString();
			Dictionary<string, string> dictionary = new Dictionary<string, string>();
			HTKKFile.AddOrUpdateDicStr(dictionary, "", "");
			foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(text))
			{
				HTKKFile.AddOrUpdateDicStr(dictionary, item.Key, item.Value);
			}
			cbbCQTChiCuc.DataSource = (object)new BindingSource((object)dictionary, (string)null);
			((ListControl)cbbCQTChiCuc).DisplayMember = "Value";
			((ListControl)cbbCQTChiCuc).ValueMember = "Key";
			((Control)cbbCQTChiCuc).Refresh();
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private Dictionary<int, List<string>> saveHeaderIfDeleteF6F8(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndex, int iRow)
	{
		SheetView sheetView = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[ToKhaiInfo.SheetInfos.Count - 1];
		int num = 12;
		int nonEmptyRowCount = sheetView.NonEmptyRowCount;
		Dictionary<int, List<string>> dictionary = new Dictionary<int, List<string>>();
		int[] array = GeneralFunctions.ArrayCellId("BA_1");
		string text = sheetView.Cells[array[1], array[0]].Text.Trim();
		if (text != "1")
		{
			return dictionary;
		}
		for (int i = num; i < nonEmptyRowCount; i++)
		{
			string formula = sheetView.Cells[i - 1, 0].Formula;
			if (string.IsNullOrEmpty(formula))
			{
				continue;
			}
			string[] array2 = GeneralFunctions.GetCellValild(formula).Split(new char[1] { '#' });
			if (array2.Length == 3 && !(winFpSpread.Sheets[sheetIndex].SheetName != array2[0]) && winFpSpread.Sheets[sheetIndex].Visible)
			{
				int num2 = HTKKFile.ColumnLettersToIndex(array2[2]);
				int num3 = int.Parse(array2[1]);
				if (iRow == num3)
				{
					List<string> list = new List<string>();
					string formula2 = sheetView.Cells[i - 1, 1].Formula;
					string formula3 = sheetView.Cells[i - 1, 2].Formula;
					list.Add(formula);
					list.Add(formula2);
					list.Add(formula3);
					dictionary.Add(i, list);
				}
			}
		}
		return dictionary;
	}

	private void revertHeaderIfDeleteF6(FpSpread winFpSpread, HTKKCMCToKhaiInfo ToKhaiInfo, int sheetIndex, Dictionary<int, List<string>> dic)
	{
		if (dic.Count == 0)
		{
			return;
		}
		SheetView sheetView = winFpSpread.Sheets[winFpSpread.Sheets.Count - 1];
		HTKKCMCSheetInfo hTKKCMCSheetInfo = ToKhaiInfo.SheetInfos[ToKhaiInfo.SheetInfos.Count - 1];
		foreach (KeyValuePair<int, List<string>> item in dic)
		{
			int key = item.Key;
			if (item.Value.Count == 3)
			{
				sheetView.Cells[key - 1, 0].Formula = item.Value[0];
				sheetView.Cells[key - 1, 1].Formula = item.Value[1];
				sheetView.Cells[key - 1, 2].Formula = item.Value[2];
			}
		}
	}

	private void cellChange(int iRow, int iCol)
	{
		//IL_6e87: Unknown result type (might be due to invalid IL or missing references)
		//IL_7056: Unknown result type (might be due to invalid IL or missing references)
		//IL_74ad: Unknown result type (might be due to invalid IL or missing references)
		if (string.IsNullOrEmpty(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text))
		{
			return;
		}
		try
		{
			string value = "";
			selectValueOld.TryGetValue(iRow.ToString() + iCol, out value);
			if (value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text && WinSettings.TTPeriodInfo.menuID != "91" && WinSettings.TTPeriodInfo.menuID != "79" && WinSettings.TTPeriodInfo.menuID != "G1" && WinSettings.TTPeriodInfo.menuID != "G4" && WinSettings.TTPeriodInfo.menuID != "71" && WinSettings.TTPeriodInfo.menuID != "01")
			{
				return;
			}
			HTKKFile.AddOrUpdateDicStr(selectValueOld, iRow.ToString() + iCol, winFpSpread.ActiveSheet.Cells[iRow, iCol].Text);
			HTKKCMCCellInfo staticCell = WinFnProcess.GetStaticCell(iRow, iCol, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
			if (staticCell != null)
			{
				string selectedValues = string.Empty;
				if (staticCell.Controltype == "8")
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower() || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						return;
					}
					if (HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsKey(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text))
					{
						selectedValues = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
						if (staticCell.ChildCell != null)
						{
							WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, staticCell.ChildCell.RowIndex, staticCell.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getCoQuanThueQuanLybyMaCoQuanThueCapCuc(selectedValues));
							winFpSpread.ActiveSheet.Cells[staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex - 1].Text = "";
						}
					}
				}
				else if (staticCell.Controltype == "9")
				{
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower() || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						return;
					}
					if (staticCell.ParentCell != null && HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsKey(winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text))
					{
						winFpSpread.ActiveSheet.Cells[iRow, iCol].CanFocus = false;
						string strMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc[winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text];
						selectedValues = HTKKWinDictionaryDiaDanh.getMaCoQuanThueQuanLy(strMaCoQuanThueCapCuc, winFpSpread.ActiveSheet.Cells[iRow, iCol].Text);
						winFpSpread.ActiveSheet.Cells[iRow, iCol].CanFocus = true;
					}
					if (WinSettings.MenuID == "100_1" && staticCell.ParentCell != null && HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc.ContainsKey(winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, iCol].Text))
					{
						winFpSpread.ActiveSheet.Cells[iRow, iCol].CanFocus = false;
						string strMaCoQuanThueCapCuc = HTKKWinDictionaryDiaDanh.CoQuanThueCapCuc[winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, iCol].Text];
						selectedValues = HTKKWinDictionaryDiaDanh.getMaCoQuanThueQuanLy(strMaCoQuanThueCapCuc, winFpSpread.ActiveSheet.Cells[iRow, iCol].Text);
						winFpSpread.ActiveSheet.Cells[iRow, iCol].CanFocus = true;
					}
				}
				else if (staticCell.Controltype == "29")
				{
					Dictionary<string, string> dictionary = Functions.getDictionary("Catalogue_DM_Lydo_GH_TT16.xml");
					selectedValues = dictionary[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
				}
				else if (staticCell.Controltype == "11")
				{
					foreach (KeyValuePair<string, string> item in HTKKWinDictionaryDiaDanh.DMNganhNgheCoDoanhThuCaoNhat)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item.Value)
						{
							selectedValues = item.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "98")
				{
					foreach (KeyValuePair<string, string> item2 in WinSpreadProcess.GetDMNganHangKBNN())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item2.Value)
						{
							selectedValues = item2.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "23")
				{
					foreach (KeyValuePair<string, string> item3 in WinSpreadProcess.DMTaiNguyen())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item3.Value)
						{
							selectedValues = item3.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "10")
				{
					foreach (KeyValuePair<string, string> item4 in HTKKWinDictionaryDiaDanh.LyDoGiaHan)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item4.Key)
						{
							selectedValues = item4.Value;
							break;
						}
					}
				}
				if (staticCell.Controltype == "59")
				{
					string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
					if (WinSpreadProcess.GetDanhMuc("GetDMMucLePhiMonBaiCT22").ContainsValue(text))
					{
						foreach (KeyValuePair<string, string> item5 in WinSpreadProcess.GetDanhMuc("GetDMMucLePhiMonBaiCT22"))
						{
							if (item5.Value.Equals(text))
							{
								selectedValues = item5.Key;
								break;
							}
						}
					}
				}
				else if (staticCell.Controltype == "54")
				{
					foreach (KeyValuePair<string, string> item6 in HTKKWinDictionaryDiaDanh.GioiTinh)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item6.Key)
						{
							selectedValues = item6.Value;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "37")
				{
					foreach (KeyValuePair<string, string> item7 in WinSpreadProcess.GetDanhMuc("GetDMNganhNgheKD"))
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item7.Value)
						{
							selectedValues = item7.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "55")
				{
					foreach (KeyValuePair<string, string> item8 in HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item8.Key)
						{
							selectedValues = item8.Value;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "64")
				{
					foreach (KeyValuePair<string, string> item9 in HTKKWinDictionaryDiaDanh.LinhVucHD)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item9.Key)
						{
							selectedValues = item9.Value;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "61")
				{
					foreach (KeyValuePair<string, string> item10 in HTKKWinDictionaryDiaDanh.NoiCap)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item10.Key)
						{
							selectedValues = item10.Value;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "70")
				{
					foreach (KeyValuePair<string, string> item11 in HTKKWinDictionaryDiaDanh.QuocTichLangGieng)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item11.Value)
						{
							selectedValues = item11.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "74")
				{
					foreach (KeyValuePair<string, string> item12 in WinSpreadProcess.GetDMLoaiDiaDiem())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item12.Value)
						{
							selectedValues = item12.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "75")
				{
					foreach (KeyValuePair<string, string> item13 in WinSpreadProcess.GetDMMienPhiMonBai())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item13.Value)
						{
							selectedValues = item13.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "79")
				{
					foreach (KeyValuePair<string, string> dMMaChiTieuKHB in WinSpreadProcess.GetDMMaChiTieuKHBS())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == dMMaChiTieuKHB.Value)
						{
							selectedValues = dMMaChiTieuKHB.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "77")
				{
					foreach (KeyValuePair<string, string> dMToKhaiPhuLucKHB in WinSpreadProcess.GetDMToKhaiPhuLucKHBS())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == dMToKhaiPhuLucKHB.Value)
						{
							selectedValues = dMToKhaiPhuLucKHB.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "78")
				{
					foreach (KeyValuePair<string, string> dMTieuMucKHB in WinSpreadProcess.GetDMTieuMucKHBS(winFpSpread))
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == dMTieuMucKHB.Value)
						{
							selectedValues = dMTieuMucKHB.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "53")
				{
					foreach (KeyValuePair<string, string> item14 in WinSpreadProcess.GetDMYKienKiemToan())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item14.Value)
						{
							selectedValues = item14.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "56")
				{
					foreach (KeyValuePair<string, string> dMQuocGium in WinSpreadProcess.GetDMQuocGia())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == dMQuocGium.Value)
						{
							selectedValues = dMQuocGium.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "4")
				{
					foreach (KeyValuePair<string, string> item15 in WinSpreadProcess.GetDMQuocGiaVietNamKhac())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item15.Value)
						{
							selectedValues = item15.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "67")
				{
					foreach (KeyValuePair<string, string> item16 in WinSpreadProcess.GetDMQuocTichFull())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item16.Value)
						{
							selectedValues = item16.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "57")
				{
					foreach (KeyValuePair<string, string> item17 in WinSpreadProcess.GetDMPPLK())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item17.Value)
						{
							selectedValues = item17.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "65")
				{
					foreach (KeyValuePair<string, string> item18 in WinSpreadProcess.GetDMPPLKND132())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item18.Value)
						{
							selectedValues = item18.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "66")
				{
					foreach (KeyValuePair<string, string> item19 in WinSpreadProcess.GetDMTM())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item19.Value)
						{
							selectedValues = item19.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "83")
				{
					foreach (KeyValuePair<string, string> item20 in WinSpreadProcess.GetDMHangHoaGTGT())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item20.Value)
						{
							selectedValues = item20.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "84")
				{
					foreach (KeyValuePair<string, string> item21 in WinSpreadProcess.GetDMThueSuatGTGT("71"))
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item21.Value)
						{
							selectedValues = item21.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "85")
				{
					foreach (KeyValuePair<string, string> item22 in WinSpreadProcess.GetDMThueSuatGTGT("01"))
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item22.Value)
						{
							selectedValues = item22.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "89")
				{
					foreach (KeyValuePair<string, string> item23 in WinSpreadProcess.DMLoaiTienDk())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item23.Value)
						{
							selectedValues = item23.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype != null && staticCell.Controltype.Length > 3 && staticCell.Controltype.Substring(0, 3) == "DC_")
				{
					foreach (KeyValuePair<string, string> item24 in WinSpreadProcess.GetDMDungChung(staticCell.Controltype))
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item24.Value)
						{
							selectedValues = item24.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "49")
				{
					foreach (KeyValuePair<string, string> item25 in HTKKWinDictionaryDiaDanh.LyDoHoanThue)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item25.Value)
						{
							selectedValues = item25.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "48")
				{
					foreach (KeyValuePair<string, string> item26 in HTKKWinDictionaryDiaDanh.LoaiHoanThue)
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item26.Value)
						{
							selectedValues = item26.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "91")
				{
					foreach (KeyValuePair<string, string> item27 in WinSpreadProcess.GetDMNganhCoThuNhapCap())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item27.Value)
						{
							selectedValues = item27.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "92")
				{
					foreach (KeyValuePair<string, string> item28 in WinSpreadProcess.GetDMChuong())
					{
						if (winFpSpread.ActiveSheet.Cells[iRow, iCol].Text == item28.Value)
						{
							selectedValues = item28.Key;
							break;
						}
					}
				}
				else if (staticCell.Controltype == "5")
				{
					if (HTKKWinDictionaryDiaDanh.dicTinhTP.ContainsKey(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text))
					{
						selectedValues = HTKKWinDictionaryDiaDanh.dicTinhTP[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
					}
					if (staticCell.ChildCell != null)
					{
						WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, staticCell.ChildCell.RowIndex, staticCell.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[selectedValues]);
						winFpSpread.ActiveSheet.Cells[staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex - 1].Text = "";
						HTKKCMCCellInfo staticCell2 = WinFnProcess.GetStaticCell(staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
						if (staticCell2 != null)
						{
							winFpSpread.ActiveSheet.Cells[staticCell2.CellSelectedValue.RowIndex - 1, staticCell2.CellSelectedValue.ColIndex - 1].Text = "";
						}
					}
				}
				if (staticCell.Controltype == "100")
				{
					if (HTKKWinDictionaryDiaDanh.dicTinhTPMoi.ContainsKey(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text))
					{
						selectedValues = HTKKWinDictionaryDiaDanh.dicTinhTPMoi[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
					}
					if (staticCell.ChildCell != null)
					{
						WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, staticCell.ChildCell.RowIndex, staticCell.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(selectedValues));
						winFpSpread.ActiveSheet.Cells[staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex - 1].Text = "";
						HTKKCMCCellInfo staticCell2 = WinFnProcess.GetStaticCell(staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
						if (staticCell2 != null && staticCell2.CellSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[staticCell2.CellSelectedValue.RowIndex - 1, staticCell2.CellSelectedValue.ColIndex - 1].Text = "";
						}
					}
				}
				if (staticCell.Controltype == "103")
				{
					if (HTKKWinDictionaryDiaDanh.dicThueTinh.ContainsKey(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text))
					{
						selectedValues = HTKKWinDictionaryDiaDanh.dicThueTinh[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
					}
					if (staticCell.ChildCell != null)
					{
						if (WinSettings.MenuID.Equals("Y9"))
						{
							string[] values = (from x in HTKKWinDictionaryDiaDanh.coQuanThus
								where x.MaCQT == selectedValues
								select x.Ten).ToArray();
							WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, staticCell.ChildCell.RowIndex, staticCell.ChildCell.ColIndex, values);
						}
						else
						{
							string[] values = (from x in HTKKWinDictionaryDiaDanh.thueTinhs
								where x.Ma == selectedValues
								select x.Ten).ToArray();
							WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, staticCell.ChildCell.RowIndex, staticCell.ChildCell.ColIndex, values);
						}
						winFpSpread.ActiveSheet.Cells[staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex - 1].Text = "";
						HTKKCMCCellInfo staticCell2 = WinFnProcess.GetStaticCell(staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex, WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].StaticCells);
						if (staticCell2 != null && staticCell2.CellSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[staticCell2.CellSelectedValue.RowIndex - 1, staticCell2.CellSelectedValue.ColIndex - 1].Text = "";
						}
					}
				}
				if (staticCell.Controltype == "6")
				{
					if (staticCell.ParentCell != null)
					{
						maTinhTP = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
					}
					selectedValues = HTKKWinDictionaryDiaDanh.getMaQuanHuyen(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text, maTinhTP);
					if (staticCell.ChildCell != null)
					{
						WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, staticCell.ChildCell.RowIndex, staticCell.ChildCell.ColIndex, HTKKWinDictionaryDiaDanh.getPhuongXa(maTinhTP, winFpSpread.ActiveSheet.Cells[iRow, iCol].Text));
						winFpSpread.ActiveSheet.Cells[staticCell.ChildCell.RowIndex - 1, staticCell.ChildCell.ColIndex - 1].Text = "";
					}
				}
				if (staticCell.Controltype == "7")
				{
					if (staticCell.ParentCell != null)
					{
						maQuanHuyen = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
					}
					selectedValues = HTKKWinDictionaryDiaDanh.getMaPhuongXa(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text, maQuanHuyen);
				}
				if (staticCell.Controltype == "101")
				{
					if (staticCell.ParentCell != null)
					{
						maTinhTP = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
					}
					selectedValues = HTKKWinDictionaryDiaDanh.getMaPhuongXaMoi(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text, maTinhTP);
					HTKKCMCSheetInfo hTKKCMCSheetInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex];
					if ("D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
					{
						GeneralFunctions.setSviewCqtCapCucTheoMaXa_Static(winFpSpread.ActiveSheet, hTKKCMCSheetInfo, selectedValues);
					}
					if (hTKKCMCSheetInfo.DataFile.ToLower() == "Header".ToLower())
					{
						HTKKCMCCellInfo hTKKCMCCellInfo = hTKKCMCSheetInfo.StaticCells.Find((HTKKCMCCellInfo x) => x.Controltype == "8");
						if (hTKKCMCCellInfo != null && staticCell.CellTemplate.RowIndex <= hTKKCMCCellInfo.CellTemplate.RowIndex)
						{
							GeneralFunctions.setSviewCqtCapCucTheoMaXa_Static(winFpSpread.ActiveSheet, hTKKCMCSheetInfo, selectedValues);
						}
					}
				}
				if (staticCell.Controltype == "102")
				{
					if (staticCell.ParentCell != null)
					{
						maTinhTP = winFpSpread.ActiveSheet.Cells[staticCell.ParentCell.RowIndex - 1, staticCell.ParentCell.ColIndex - 1].Text;
					}
					selectedValues = HTKKWinDictionaryDiaDanh.coQuanThus.FirstOrDefault((CoQuanThu x) => x.MaCQT == maTinhTP).Ma;
				}
				if (staticCell.CellSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[staticCell.CellSelectedValue.RowIndex - 1, staticCell.CellSelectedValue.ColIndex - 1].Text = selectedValues;
				}
				if (!(staticCell.Controltype == "25"))
				{
				}
				return;
			}
			HTKKCMCTableCellInfo tableCellInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableCellInfo(iRow + 1, iCol + 1);
			int tableIndex = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableIndex(iRow + 1, iCol + 1);
			HTKKCMCTableInfo hTKKCMCTableInfo = null;
			if (tableIndex >= 0)
			{
				hTKKCMCTableInfo = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].Tables[tableIndex];
			}
			if (tableCellInfo.Controltype == "3" && tableCellInfo.indexColumnSelectedValue != null)
			{
				string text2 = HTKKWinDictionaryDiaDanh.QuanHe[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
				winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
			}
			if (tableCellInfo.Controltype == "4")
			{
				string value2 = "";
				GeneralFunctions.dicVar.TryGetValue("QuocTich", out value2);
				if (ConvertHelpers.ConvertToBool(value2))
				{
					return;
				}
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					string text2 = HTKKWinDictionaryDiaDanh.QuocTich[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
					if (tableCellInfo.ChildValue != null)
					{
						WinSpreadProcess.WinInsertComboxKey(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getTinhTP(text2));
					}
					if ((WinSettings.MenuID == "37" || WinSettings.MenuID == "95") && text2 == "02")
					{
						winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue)].Text = "99999";
						WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) + 2, HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP["ALL"]);
						winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) + 2].Text = "9999999";
						WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) + 4, HTKKWinDictionaryDiaDanh.getPhuongXa("999", "Khác"));
					}
				}
			}
			if (tableCellInfo.Controltype == "56")
			{
				string value2 = "";
				GeneralFunctions.dicVar.TryGetValue("QuocTich", out value2);
				if (ConvertHelpers.ConvertToBool(value2))
				{
					return;
				}
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					string text2 = HTKKWinDictionaryDiaDanh.QuocGia.FirstOrDefault((KeyValuePair<string, string> p) => p.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text).Key;
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
					if (tableCellInfo.ChildValue != null)
					{
						WinSpreadProcess.WinInsertComboxKey(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getTinhTP(text2));
						if (!(WinSettings.MenuID == "B2") && !(WinSettings.MenuID == "O9"))
						{
						}
					}
				}
			}
			if (tableCellInfo.Controltype == "67")
			{
				string value2 = "";
				GeneralFunctions.dicVar.TryGetValue("QuocTich", out value2);
				if (ConvertHelpers.ConvertToBool(value2))
				{
					return;
				}
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					string text2 = HTKKWinDictionaryDiaDanh.QuocTichFull.FirstOrDefault((KeyValuePair<string, string> p) => p.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text).Key;
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
					if (tableCellInfo.ChildValue != null)
					{
						WinSpreadProcess.WinInsertComboxKey(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getTinhTP(text2));
						if (!(WinSettings.MenuID == "B2") && !(WinSettings.MenuID == "O9"))
						{
						}
					}
				}
			}
			if (tableCellInfo.Controltype == "5")
			{
				maTinhTP = HTKKWinDictionaryDiaDanh.dicTinhTP[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maTinhTP;
				}
				if (tableCellInfo.ChildValue != null)
				{
					WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.dicQuanHuyenByTinhTP[maTinhTP]);
				}
			}
			if (tableCellInfo.Controltype == "100")
			{
				maTinhTP = HTKKWinDictionaryDiaDanh.dicTinhTPMoi[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maTinhTP;
				}
				if (tableCellInfo.ChildValue != null)
				{
					WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getDanhMucPhuongXaMoi(maTinhTP));
				}
				if (WinSettings.MenuID == "N2" || WinSettings.MenuID == "Q4")
				{
					HTKKCMCTableCellInfo hTKKCMCTableCellInfo = hTKKCMCTableInfo.RowsInfo.Cells.Find((HTKKCMCTableCellInfo x) => x.Controltype == "107");
					if (hTKKCMCTableCellInfo != null)
					{
						WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow + 1, hTKKCMCTableCellInfo.CellTemplate.ColIndex, WinSpreadProcess.GetDMDynamicByParent(maTinhTP, "DM_Kho_Bac.xml"));
					}
				}
			}
			if (tableCellInfo.Controltype == "6")
			{
				if (tableCellInfo.ParentValue != null)
				{
					maTinhTP = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
				}
				maQuanHuyen = HTKKWinDictionaryDiaDanh.getMaQuanHuyen(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text, maTinhTP);
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maQuanHuyen;
				}
				if (tableCellInfo.ChildValue != null && !tableCellInfo.ChildValue.Equals(""))
				{
					WinSpreadProcess.WinInsertCombox(winFpSpread.ActiveSheet, iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue), HTKKWinDictionaryDiaDanh.getPhuongXa(maTinhTP, winFpSpread.ActiveSheet.Cells[iRow, iCol].Text));
					if (WinSettings.MenuID == "M6" || WinSettings.MenuID == "O8")
					{
						HTKKCMCTableCellInfo tableCellInfo2 = WinSettings.ToKhaiInfo.SheetInfos[winFpSpread.ActiveSheetIndex].GetTableCellInfo(iRow + 1, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue));
						winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo2.indexColumnSelectedValue) - 1].Text = "";
					}
					else
					{
						winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 2].Text = "";
					}
				}
			}
			if (tableCellInfo.Controltype == "7")
			{
				if (tableCellInfo.ParentValue != null)
				{
					maQuanHuyen = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
				}
				string maPhuongXa = HTKKWinDictionaryDiaDanh.getMaPhuongXa(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text, maQuanHuyen);
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maPhuongXa;
				}
			}
			if (tableCellInfo.Controltype == "101")
			{
				if (tableCellInfo.ParentValue != null)
				{
					maTinhTP = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
				}
				string maPhuongXa = HTKKWinDictionaryDiaDanh.getMaPhuongXaMoi(winFpSpread.ActiveSheet.Cells[iRow, iCol].Text, maTinhTP);
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maPhuongXa;
				}
			}
			if (tableCellInfo.Controltype == "102")
			{
				if (tableCellInfo.ParentValue != null)
				{
					maTinhTP = winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.ParentValue) - 1].Text;
				}
				string maPhuongXa = HTKKWinDictionaryDiaDanh.coQuanThus.FirstOrDefault((CoQuanThu x) => x.MaCQT == maTinhTP).Ma;
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maPhuongXa;
				}
			}
			if (tableCellInfo.Controltype == "8")
			{
				if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					return;
				}
				Dictionary<string, string> dMFromCombobox = GeneralFunctions.GetDMFromCombobox(winFpSpread.ActiveSheet, iRow + 1, iCol + 1);
				string tenCqt = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (tenCqt == "Chọn chi cục thuế khu vực")
				{
					return;
				}
				string text2 = dMFromCombobox.First((KeyValuePair<string, string> x) => x.Value == tenCqt).Key;
				if (tableCellInfo.indexColumnSelectedValue != null)
				{
					winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
				}
				if (tableCellInfo.ChildValue != null)
				{
					GeneralFunctions.setSviewCqtCapChiCucTheoMaCqtCuc_Dinamic(winFpSpread.ActiveSheet, iRow + 1, hTKKCMCTableInfo.RowsInfo, text2);
				}
			}
			if (tableCellInfo.Controltype == "9")
			{
				if (WinSettings.ToKhaiInfo.NopTheoTinhCQT == "3" || "D1,N8,G9,H2,H4,T8,E8,L8,L7,N2,G2,D2,Q5,Q6,U6,Y6,Z8,Z5".Contains(WinSettings.MenuID))
				{
					return;
				}
				foreach (KeyValuePair<string, string> item29 in HTKKWinDictionaryDiaDanh.CoQuanThueQuanLy)
				{
					if (item29.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						string text2 = item29.Key;
						if (tableCellInfo.indexColumnSelectedValue != null)
						{
							winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text2;
						}
					}
				}
			}
			if (tableCellInfo.Controltype == "20")
			{
				string text3 = "";
				string text4 = "";
				foreach (KeyValuePair<string, string> item30 in WinSpreadProcess.DMKhoangSan("GetDMKhoangSan"))
				{
					if (item30.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text3 = item30.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text3))
				{
					text4 = text3;
					if (text3.Split(new char[1] { '_' }).Length > 3)
					{
						text4 = text3.Split(new char[1] { '_' })[2];
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text4;
				}
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
				{
					Dictionary<string, string> lstDonVi = new Dictionary<string, string>();
					Dictionary<string, string> dictionary2 = WinSpreadProcess.DMKhoangSanDonVi(ref lstDonVi, "GetDMKhoangSan");
					if (dictionary2 != null && dictionary2.ContainsKey(text3))
					{
						string text5 = WinSpreadProcess.DMKhoangSanDonVi(ref lstDonVi, "GetDMKhoangSan")[text3];
						winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text5;
					}
				}
			}
			if (tableCellInfo.Controltype == "62")
			{
				string text3 = "";
				foreach (KeyValuePair<string, string> item31 in WinSpreadProcess.DMKhoangSan("GetDMKSBVMT"))
				{
					if (item31.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text3 = item31.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text3))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text3;
				}
				if (!string.IsNullOrEmpty(text3) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
				{
					Dictionary<string, string> lstDonVi = new Dictionary<string, string>();
					string text5 = WinSpreadProcess.DMKhoangSanDonVi(ref lstDonVi, "GetDMKSBVMT")[text3];
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text5;
				}
			}
			if (tableCellInfo.Controltype == "53")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item32 in WinSpreadProcess.GetDMYKienKiemToan())
				{
					if (item32.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item32.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "80" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text7 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMThueSuat().ContainsValue(text7))
				{
					string text8 = "";
					foreach (KeyValuePair<string, string> item33 in WinSpreadProcess.DMThueSuat())
					{
						if (item33.Value.Equals(text7))
						{
							text8 = item33.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text8;
				}
			}
			if (tableCellInfo.Controltype == "72" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text9 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMMo().ContainsValue(text9))
				{
					string text10 = "";
					foreach (KeyValuePair<string, string> item34 in WinSpreadProcess.DMMo())
					{
						if (item34.Value.Equals(text9))
						{
							text10 = item34.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text10;
				}
			}
			if (tableCellInfo.Controltype == "81" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text11 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiDat().ContainsValue(text11))
				{
					string text12 = "";
					foreach (KeyValuePair<string, string> item35 in WinSpreadProcess.DMLoaiDat())
					{
						if (item35.Value.Equals(text11))
						{
							text12 = item35.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text12;
				}
			}
			if (tableCellInfo.Controltype == "82" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text11 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMHangDat().ContainsValue(text11))
				{
					string text13 = "";
					foreach (KeyValuePair<string, string> item36 in WinSpreadProcess.DMHangDat())
					{
						if (item36.Value.Equals(text11))
						{
							text13 = item36.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text13;
				}
			}
			if (tableCellInfo.Controltype == "87" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text14 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMNnkdNtnn().ContainsValue(text14))
				{
					string text15 = "";
					foreach (KeyValuePair<string, string> item37 in WinSpreadProcess.DMNnkdNtnn())
					{
						if (item37.Value.Equals(text14))
						{
							text15 = item37.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text15;
				}
			}
			if (tableCellInfo.Controltype == "89" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text16 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiTienDk().ContainsValue(text16))
				{
					string text17 = "";
					foreach (KeyValuePair<string, string> item38 in WinSpreadProcess.DMLoaiTienDk())
					{
						if (item38.Value.Equals(text16))
						{
							text17 = item38.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text17;
				}
			}
			if (tableCellInfo.Controltype != null && tableCellInfo.Controltype.Length > 3 && tableCellInfo.Controltype.Substring(0, 3) == "DC_" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text18 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiTienDk().ContainsValue(text18))
				{
					string text19 = "";
					foreach (KeyValuePair<string, string> item39 in WinSpreadProcess.GetDMDungChung(tableCellInfo.Controltype))
					{
						if (item39.Value.Equals(text18))
						{
							text19 = item39.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text19;
				}
			}
			if (tableCellInfo.Controltype == "56")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> dMQuocGium2 in WinSpreadProcess.GetDMQuocGia())
				{
					if (dMQuocGium2.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = dMQuocGium2.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "67")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item40 in WinSpreadProcess.GetDMQuocTichFull())
				{
					if (item40.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item40.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "57")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item41 in WinSpreadProcess.GetDMPPLK())
				{
					if (item41.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item41.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "65")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item42 in WinSpreadProcess.GetDMPPLKND132())
				{
					if (item42.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item42.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "66")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item43 in WinSpreadProcess.GetDMTM())
				{
					if (item43.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item43.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "83")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item44 in WinSpreadProcess.GetDMHangHoaGTGT())
				{
					if (item44.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item44.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "84")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item45 in WinSpreadProcess.GetDMThueSuatGTGT("71"))
				{
					if (item45.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item45.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "85")
			{
				string text6 = "";
				foreach (KeyValuePair<string, string> item46 in WinSpreadProcess.GetDMThueSuatGTGT("01"))
				{
					if (item46.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text6 = item46.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text6))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text6;
				}
			}
			if (tableCellInfo.Controltype == "23")
			{
				string text3 = "";
				foreach (KeyValuePair<string, string> item47 in WinSpreadProcess.DMTaiNguyen())
				{
					if (item47.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text3 = item47.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text3))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text3;
				}
			}
			if (tableCellInfo.Controltype == "22")
			{
				string text20 = "";
				foreach (KeyValuePair<string, string> item48 in WinSpreadProcess.DMPhiLephi())
				{
					if (item48.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text20 = item48.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text20))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text20;
				}
				if (!string.IsNullOrEmpty(text20) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
				{
					string text21 = WinSpreadProcess.DMPhiLephi_TieuMuc()[text20];
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text21;
				}
			}
			if (tableCellInfo.Controltype == "88")
			{
				string text20 = "";
				foreach (KeyValuePair<string, string> item49 in WinSpreadProcess.DMPhiLephiTt63())
				{
					if (item49.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text20 = item49.Key;
						break;
					}
				}
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text20))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text20;
				}
				if (!string.IsNullOrEmpty(text20) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
				{
					string text21 = WinSpreadProcess.DMPhiLephiTt63_TieuMuc()[text20];
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text21;
				}
			}
			if (tableCellInfo.Controltype == "86")
			{
				string text22 = "";
				Dictionary<string, string> dMFromCombobox = WinSpreadProcess.DMPHLPNG();
				foreach (KeyValuePair<string, string> item50 in dMFromCombobox)
				{
					if (item50.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
					{
						text22 = item50.Key;
						break;
					}
				}
				item itemDMPHLPNGById = WinSpreadProcess.GetItemDMPHLPNGById(text22);
				if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text22))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = itemDMPHLPNGById.ma;
				}
				if (!string.IsNullOrEmpty(text22) && !string.IsNullOrEmpty(tableCellInfo.ChildValue))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = itemDMPHLPNGById.matieumuc;
				}
			}
			if (tableCellInfo.Controltype == "25")
			{
				string text23 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiTemRuou().ContainsKey(text23) && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text23))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = WinSpreadProcess.DMLoaiTemRuou()[text23];
				}
				if (!string.IsNullOrEmpty(tableCellInfo.ChildValue))
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.ChildValue) - 1].Text = text23;
				}
			}
			if (tableCellInfo.Controltype == "28" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text24 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMGhiChuTemRuou().ContainsValue(text24))
				{
					string text25 = "";
					foreach (KeyValuePair<string, string> item51 in WinSpreadProcess.DMGhiChuTemRuou())
					{
						if (item51.Value.Equals(text24))
						{
							text25 = item51.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text25;
				}
			}
			if (tableCellInfo.Controltype == "73" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text16 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiTienTe().ContainsValue(text16))
				{
					string text17 = "";
					foreach (KeyValuePair<string, string> item52 in WinSpreadProcess.DMLoaiTienTe())
					{
						if (item52.Value.Equals(text16))
						{
							text17 = item52.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text17;
				}
			}
			if (tableCellInfo.Controltype == "80" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text7 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMThueSuat().ContainsValue(text7))
				{
					string text8 = "";
					foreach (KeyValuePair<string, string> item53 in WinSpreadProcess.DMThueSuat())
					{
						if (item53.Value.Equals(text7))
						{
							text8 = item53.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text8;
				}
			}
			if (tableCellInfo.Controltype == "72" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text9 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMMo().ContainsValue(text9))
				{
					string text10 = "";
					foreach (KeyValuePair<string, string> item54 in WinSpreadProcess.DMMo())
					{
						if (item54.Value.Equals(text9))
						{
							text10 = item54.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text10;
				}
			}
			if (tableCellInfo.Controltype == "81" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text11 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiDat().ContainsValue(text11))
				{
					string text8 = "";
					foreach (KeyValuePair<string, string> item55 in WinSpreadProcess.DMLoaiDat())
					{
						if (item55.Value.Equals(text11))
						{
							text8 = item55.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text8;
				}
			}
			if (tableCellInfo.Controltype == "82" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text26 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMHangDat().ContainsValue(text26))
				{
					string text8 = "";
					foreach (KeyValuePair<string, string> item56 in WinSpreadProcess.DMHangDat())
					{
						if (item56.Value.Equals(text26))
						{
							text8 = item56.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text8;
				}
			}
			if (tableCellInfo.Controltype == "87" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text14 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMNnkdNtnn().ContainsValue(text14))
				{
					string text15 = "";
					foreach (KeyValuePair<string, string> item57 in WinSpreadProcess.DMNnkdNtnn())
					{
						if (item57.Value.Equals(text14))
						{
							text15 = item57.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text15;
				}
			}
			if (tableCellInfo.Controltype == "89" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text16 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.DMLoaiTienDk().ContainsValue(text16))
				{
					string text17 = "";
					foreach (KeyValuePair<string, string> item58 in WinSpreadProcess.DMLoaiTienDk())
					{
						if (item58.Value.Equals(text16))
						{
							text17 = item58.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text17;
				}
			}
			if (tableCellInfo.Controltype != null && tableCellInfo.Controltype.Length > 3 && tableCellInfo.Controltype.Substring(0, 3) == "DC_" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text18 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.GetDMDungChung(tableCellInfo.Controltype).ContainsValue(text18))
				{
					string text19 = "";
					foreach (KeyValuePair<string, string> item59 in WinSpreadProcess.DMLoaiTienDk())
					{
						if (item59.Value.Equals(text18))
						{
							text19 = item59.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text19;
				}
			}
			if (tableCellInfo.Controltype == "35" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.GetDanhMuc("GetDMThueBVMT").ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item60 in WinSpreadProcess.GetDanhMuc("GetDMThueBVMT"))
					{
						if (item60.Value.Equals(text))
						{
							text27 = item60.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "36" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.GetDanhMuc("GetDMThan").ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item61 in WinSpreadProcess.GetDanhMuc("GetDMThan"))
					{
						if (item61.Value.Equals(text))
						{
							text27 = item61.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "71" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.GetDanhMuc("GetDMXangDau").ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item62 in WinSpreadProcess.GetDanhMuc("GetDMXangDau"))
					{
						if (item62.Value.Equals(text))
						{
							text27 = item62.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "54" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (HTKKWinDictionaryDiaDanh.GioiTinh.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item63 in HTKKWinDictionaryDiaDanh.GioiTinh)
					{
						if (item63.Value.Equals(text))
						{
							text27 = item63.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "55" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item64 in HTKKWinDictionaryDiaDanh.LoaiGiayToTuyThan)
					{
						if (item64.Value.Equals(text))
						{
							text27 = item64.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "64" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (HTKKWinDictionaryDiaDanh.LinhVucHD.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item65 in HTKKWinDictionaryDiaDanh.LinhVucHD)
					{
						if (item65.Value.Equals(text))
						{
							text27 = item65.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "61" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (HTKKWinDictionaryDiaDanh.NoiCap.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item66 in HTKKWinDictionaryDiaDanh.NoiCap)
					{
						if (item66.Value.Equals(text))
						{
							text27 = item66.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "50" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> khoanNopNSSHoanThue = HTKKWinDictionaryDiaDanh.KhoanNopNSSHoanThue;
				if (khoanNopNSSHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item67 in khoanNopNSSHoanThue)
					{
						if (item67.Value.Equals(text))
						{
							text27 = item67.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "51" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = HTKKWinDictionaryDiaDanh.KyPhatSinhHoanThue;
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item68 in kyPhatSinhHoanThue)
					{
						if (item68.Value.Equals(text))
						{
							text27 = item68.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "90" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetMaTieuMuc();
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item69 in kyPhatSinhHoanThue)
					{
						if (item69.Value.Equals(text))
						{
							text27 = item69.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "75" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetDMMienPhiMonBai();
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item70 in kyPhatSinhHoanThue)
					{
						if (item70.Value.Equals(text))
						{
							text27 = item70.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "79" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetDMMaChiTieuKHBS();
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item71 in kyPhatSinhHoanThue)
					{
						if (item71.Value.Equals(text))
						{
							text27 = item71.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "74" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetDMLoaiDiaDiem();
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item72 in kyPhatSinhHoanThue)
					{
						if (item72.Value.Equals(text))
						{
							text27 = item72.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "37" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> danhMuc = WinSpreadProcess.GetDanhMuc("GetDMNganhNgheKD");
				if (danhMuc.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item73 in danhMuc)
					{
						if (item73.Value.Equals(text))
						{
							text27 = item73.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "77" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetDMToKhaiPhuLucKHBS();
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item74 in kyPhatSinhHoanThue)
					{
						if (item74.Value.Equals(text))
						{
							text27 = item74.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "78" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetDMTieuMucKHBS(winFpSpread);
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item75 in kyPhatSinhHoanThue)
					{
						if (item75.Value.Equals(text))
						{
							text27 = item75.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype != null && tableCellInfo.Controltype.Length > 3 && tableCellInfo.Controltype == "DC_" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				Dictionary<string, string> kyPhatSinhHoanThue = WinSpreadProcess.GetDMDungChung(tableCellInfo.Controltype);
				if (kyPhatSinhHoanThue.ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item76 in kyPhatSinhHoanThue)
					{
						if (item76.Value.Equals(text))
						{
							text27 = item76.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "39")
			{
			}
			if (tableCellInfo.Controltype == "43")
			{
				HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
				SheetView sheetView = winFpSpread.Sheets[0];
				string value3 = sheetView.Cells[iRow, iCol].Value.ToString();
				HTKKCMCSheetInfo hTKKCMCSheetInfo2 = toKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				int num = dataLocationTemplateRowStart + hTKKCMCTableInfo2.RowsValues.Count - 1;
				Dictionary<string, string> dictionary3 = new Dictionary<string, string>();
				string empty = string.Empty;
				for (int num2 = dataLocationTemplateRowStart; num2 <= num; num2++)
				{
					empty = sheetView.Cells[num2 - 1, iCol].Text;
					if (empty.Equals(value3) && num2 - 1 != iRow)
					{
						frmMessage frmMessage = new frmMessage("0445", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						sheetView.Cells[iRow, iCol].Text = "";
						sheetView.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
						sheetView.Cells[iRow, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = "";
						break;
					}
				}
			}
			if (tableCellInfo.Controltype == "64")
			{
			}
			if (tableCellInfo.Controltype == "58")
			{
				HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
				SheetView sheetView = winFpSpread.Sheets[0];
				string value3 = sheetView.Cells[iRow, iCol].Value.ToString();
				HTKKCMCSheetInfo hTKKCMCSheetInfo2 = toKhaiInfo.SheetInfos[0];
				HTKKCMCTableInfo hTKKCMCTableInfo2 = hTKKCMCSheetInfo2.Tables[0];
				int dataLocationTemplateRowStart = hTKKCMCTableInfo2.Location.DataLocationTemplateRowStart;
				int num = dataLocationTemplateRowStart + hTKKCMCTableInfo2.RowsValues.Count - 1;
				Dictionary<string, string> dictionary3 = new Dictionary<string, string>();
				string empty = string.Empty;
				for (int num2 = dataLocationTemplateRowStart; num2 <= num; num2++)
				{
					empty = sheetView.Cells[num2 - 1, iCol].Text;
					if (empty.Equals(value3) && num2 - 1 != iRow)
					{
						frmMessage frmMessage = new frmMessage("0445", "3", "1", "Thông báo");
						((Form)frmMessage).ShowDialog();
						sheetView.Cells[iRow, iCol].Text = "";
						sheetView.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
						sheetView.Cells[iRow, HTKKFile.ColumnLettersToIndex("AA") - 1].Text = "";
						break;
					}
				}
			}
			if (tableCellInfo.Controltype == "59" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.GetDanhMuc("GetDMMucLePhiMonBaiCT22").ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item77 in WinSpreadProcess.GetDanhMuc("GetDMMucLePhiMonBaiCT22"))
					{
						if (item77.Value.Equals(text))
						{
							text27 = item77.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "60" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				string text = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (WinSpreadProcess.GetDanhMuc("GetDMMucLePhiMonBaiCT23").ContainsValue(text))
				{
					string text27 = "";
					foreach (KeyValuePair<string, string> item78 in WinSpreadProcess.GetDanhMuc("GetDMMucLePhiMonBaiCT23"))
					{
						if (item78.Value.Equals(text))
						{
							text27 = item78.Key;
							break;
						}
					}
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text27;
				}
			}
			if (tableCellInfo.Controltype == "42")
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string value3 = sheetView.Cells[iRow, iCol].Value.ToString();
				string[] array = new string[1] { "01_TTDB_II" };
				string[] array2 = array;
				foreach (string tableName in array2)
				{
					HTKKCMCToKhaiInfo toKhaiInfo = WinSettings.ToKhaiInfo;
					HTKKCMCSheetInfo hTKKCMCSheetInfo2 = toKhaiInfo.SheetInfos[0];
					List<HTKKCMCTableInfo> tables = hTKKCMCSheetInfo2.Tables;
					Predicate<HTKKCMCTableInfo> match = (HTKKCMCTableInfo table) => table.TableName == tableName;
					HTKKCMCTableInfo hTKKCMCTableInfo3 = tables.Find(match);
					int dataLocationTemplateRowStart2 = hTKKCMCTableInfo3.Location.DataLocationTemplateRowStart;
					int num4 = dataLocationTemplateRowStart2 + hTKKCMCTableInfo3.RowsValues.Count - 1;
					string empty2 = string.Empty;
					string empty3 = string.Empty;
					Dictionary<string, string> dictionary3 = new Dictionary<string, string>();
					for (int num2 = dataLocationTemplateRowStart2; num2 <= num4; num2++)
					{
						empty2 = sheetView.Cells[num2 - 1, iCol].Text;
						if (empty2.Equals(value3) && num2 - 1 != iRow && iRow >= dataLocationTemplateRowStart2 && iRow <= num4)
						{
							frmMessage frmMessage = new frmMessage("0261", "3", "1", "Thông báo");
							((Form)frmMessage).ShowDialog();
							sheetView.Cells[iRow, iCol].Text = "";
							sheetView.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = "";
							break;
						}
					}
				}
			}
			if ((WinSettings.MenuID == "B2" || WinSettings.MenuID == "O9") && iCol == 21 && iRow >= 42)
			{
				SheetView sheetView = winFpSpread.Sheets[0];
				string text28 = sheetView.Cells[iRow, HTKKFile.ColumnLettersToIndex("R") - 1].Text;
				if (text28 == "1020")
				{
					string text29 = "";
					foreach (KeyValuePair<string, string> dMQuocGium3 in WinSpreadProcess.GetDMQuocGia())
					{
						if (dMQuocGium3.Value == winFpSpread.ActiveSheet.Cells[iRow, iCol].Text)
						{
							text29 = dMQuocGium3.Key;
							break;
						}
					}
					if (!string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue) && !string.IsNullOrEmpty(text29))
					{
						winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text29;
					}
				}
				else
				{
					maTinhTP = HTKKWinDictionaryDiaDanh.dicTinhTP[winFpSpread.ActiveSheet.Cells[iRow, iCol].Text];
					if (tableCellInfo.indexColumnSelectedValue != null)
					{
						winFpSpread.ActiveSheet.Cells[winFpSpread.ActiveSheet.ActiveCell.Row.Index, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = maTinhTP;
					}
				}
			}
			if (tableCellInfo.Controltype != null && tableCellInfo.Controltype.Length > 3 && tableCellInfo.Controltype.Substring(0, 3) == "DC_" && !string.IsNullOrEmpty(tableCellInfo.indexColumnSelectedValue))
			{
				Dictionary<string, string> dictionary4 = new Dictionary<string, string>();
				if (WinSettings.MenuID != "N8" && winFpSpread.ActiveSheetIndex != 0 && winFpSpread.ActiveSheetIndex != 5 && winFpSpread.ActiveSheetIndex != 6)
				{
					dictionary4 = WinSpreadProcess.GetDMDungChung(tableCellInfo.Controltype);
				}
				else
				{
					int dataLocationTemplateRowStart3 = hTKKCMCTableInfo.Location.DataLocationTemplateRowStart;
					dictionary4 = GeneralFunctions.GetDMFromCombobox(winFpSpread.ActiveSheet, dataLocationTemplateRowStart3, tableCellInfo.CellTemplate.ColIndex);
				}
				string text18 = winFpSpread.ActiveSheet.Cells[iRow, iCol].Text;
				if (text18.Length > 300)
				{
					_tenDai = text18;
				}
				else
				{
					_tenDai = "";
				}
				string text19 = "";
				foreach (KeyValuePair<string, string> item79 in dictionary4)
				{
					if (item79.Value.Equals(text18))
					{
						text19 = item79.Key;
						break;
					}
				}
				if (string.IsNullOrEmpty(text19))
				{
					winFpSpread.ActiveSheet.Cells[iRow, iCol].Text = "";
				}
				else
				{
					winFpSpread.ActiveSheet.Cells[iRow, HTKKFile.ColumnLettersToIndex(tableCellInfo.indexColumnSelectedValue) - 1].Text = text19;
				}
			}
			Row row = winFpSpread.ActiveSheet.Rows[iRow];
			float preferredHeight = row.GetPreferredHeight();
			if (row.Height < preferredHeight - 5f || row.Height > preferredHeight + 5f)
			{
				row.Height = preferredHeight;
			}
		}
		catch (Exception ex)
		{
			HTKKFile.Log(MethodBase.GetCurrentMethod().Name, ex);
		}
	}

	private void txtPage_TextChanged(object sender, EventArgs e)
	{
		GeneralFunctions._page = Convert.ToInt32(((Control)txtPage).Text) - 1;
		WinDoWorkEvent.UpadatePage(winFpSpread, winFpSpread.ActiveSheetIndex, GeneralFunctions._page);
	}

	private void btnPrivPage_Click(object sender, EventArgs e)
	{
		if (Convert.ToInt32(((Control)txtPage).Text) > 1)
		{
			((Control)txtPage).Text = Convert.ToString(Convert.ToInt32(((Control)txtPage).Text) - 1);
		}
	}

	private void btnFirstPage_Click(object sender, EventArgs e)
	{
		((Control)txtPage).Text = "1";
	}

	private void btnNextPage_Click(object sender, EventArgs e)
	{
		if (Convert.ToInt32(((Control)txtPage).Text) < GeneralFunctions._totalPage)
		{
			((Control)txtPage).Text = Convert.ToString(Convert.ToInt32(((Control)txtPage).Text) + 1);
		}
	}

	private void btnEndPage_Click(object sender, EventArgs e)
	{
		((Control)txtPage).Text = GeneralFunctions._totalPage.ToString();
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
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b6: Expected O, but got Unknown
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c1: Expected O, but got Unknown
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cc: Expected O, but got Unknown
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Expected O, but got Unknown
		//IL_00d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e2: Expected O, but got Unknown
		//IL_00e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ed: Expected O, but got Unknown
		//IL_00ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f8: Expected O, but got Unknown
		//IL_00f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_0103: Expected O, but got Unknown
		//IL_0104: Unknown result type (might be due to invalid IL or missing references)
		//IL_010e: Expected O, but got Unknown
		//IL_010f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0119: Expected O, but got Unknown
		//IL_011a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0124: Expected O, but got Unknown
		//IL_0125: Unknown result type (might be due to invalid IL or missing references)
		//IL_012f: Expected O, but got Unknown
		//IL_0130: Unknown result type (might be due to invalid IL or missing references)
		//IL_013a: Expected O, but got Unknown
		//IL_013b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0145: Expected O, but got Unknown
		//IL_0146: Unknown result type (might be due to invalid IL or missing references)
		//IL_0150: Expected O, but got Unknown
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_015b: Expected O, but got Unknown
		//IL_015c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0166: Expected O, but got Unknown
		//IL_0167: Unknown result type (might be due to invalid IL or missing references)
		//IL_0171: Expected O, but got Unknown
		//IL_0172: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Expected O, but got Unknown
		//IL_017d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0187: Expected O, but got Unknown
		//IL_0188: Unknown result type (might be due to invalid IL or missing references)
		//IL_0192: Expected O, but got Unknown
		//IL_0193: Unknown result type (might be due to invalid IL or missing references)
		//IL_019d: Expected O, but got Unknown
		//IL_019e: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a8: Expected O, but got Unknown
		//IL_01a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b3: Expected O, but got Unknown
		//IL_01b4: Unknown result type (might be due to invalid IL or missing references)
		//IL_01be: Expected O, but got Unknown
		//IL_01bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_01c9: Expected O, but got Unknown
		//IL_01ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d4: Expected O, but got Unknown
		//IL_01d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01df: Expected O, but got Unknown
		//IL_01e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ea: Expected O, but got Unknown
		//IL_01eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f5: Expected O, but got Unknown
		//IL_049e: Unknown result type (might be due to invalid IL or missing references)
		//IL_04a8: Expected O, but got Unknown
		//IL_04ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_04d8: Expected O, but got Unknown
		//IL_207b: Unknown result type (might be due to invalid IL or missing references)
		//IL_2085: Expected O, but got Unknown
		//IL_20be: Unknown result type (might be due to invalid IL or missing references)
		//IL_20c8: Expected O, but got Unknown
		//IL_20d1: Unknown result type (might be due to invalid IL or missing references)
		//IL_20db: Expected O, but got Unknown
		//IL_20f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_2101: Expected O, but got Unknown
		//IL_210a: Unknown result type (might be due to invalid IL or missing references)
		//IL_2114: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(frmWinToKhai));
		opfdlg = new OpenFileDialog();
		winFpSpread = new FpSpread();
		winFpSpread_Sheet1 = new SheetView();
		pnWinSpread = new Panel();
		pnTimKiem = new Panel();
		cbbTimKiem = new ComboBox();
		txtTimKiem = new TextBox();
		btnTimKiem = new Button();
		cmdClear = new Button();
		cmdLoadToKhai = new Button();
		cmdSave = new Button();
		cmdPrint = new Button();
		cmdDelete = new Button();
		cmdKiemTra = new Button();
		btnThemPhuLuc = new Button();
		cmdExport = new Button();
		cmdImportXML = new Button();
		cmdExit = new Button();
		btnTaiBangKe = new Button();
		btnTongHopKHBS = new Button();
		pnButon = new Panel();
		btnCustom = new Button();
		pnF7 = new Panel();
		label1 = new Label();
		label2 = new Label();
		lblInfo = new Label();
		panelChild = new Panel();
		pnF5 = new Panel();
		lblThemDong = new Label();
		lblXoaDong = new Label();
		pnInsert = new Panel();
		panelCQT = new Panel();
		cbbCQTChiCuc = new ComboBox();
		label3 = new Label();
		lbCQTCapCuc = new Label();
		cbbCQTCuc = new ComboBox();
		panelPage = new Panel();
		btnFirstPage = new Button();
		btnPrivPage = new Button();
		txtPage = new TextBox();
		label4 = new Label();
		labelTotalPage = new Label();
		btnNextPage = new Button();
		btnEndPage = new Button();
		((ISupportInitialize)winFpSpread).BeginInit();
		((ISupportInitialize)winFpSpread_Sheet1).BeginInit();
		((Control)pnWinSpread).SuspendLayout();
		((Control)pnTimKiem).SuspendLayout();
		((Control)pnButon).SuspendLayout();
		((Control)pnF7).SuspendLayout();
		((Control)pnF5).SuspendLayout();
		((Control)pnInsert).SuspendLayout();
		((Control)panelCQT).SuspendLayout();
		((Control)panelPage).SuspendLayout();
		((Control)this).SuspendLayout();
		((FileDialog)opfdlg).FileName = "openFileDialog1";
		((Control)winFpSpread).AccessibleDescription = "winFpSpread, Sheet1, Row 0, Column 0, ";
		((Control)winFpSpread).Anchor = (AnchorStyles)15;
		((Control)winFpSpread).BackColor = SystemColors.Control;
		((Control)winFpSpread).BackgroundImageLayout = (ImageLayout)0;
		((Control)winFpSpread).Location = new Point(0, -2);
		winFpSpread.Name = "winFpSpread";
		winFpSpread.ScrollBarTrackPolicy = ScrollBarTrackPolicy.Both;
		winFpSpread.ScrollTipPolicy = ScrollTipPolicy.Both;
		winFpSpread.Sheets.AddRange(new SheetView[1] { winFpSpread_Sheet1 });
		((Control)winFpSpread).Size = new Size(1266, 505);
		((Control)winFpSpread).TabIndex = 25;
		winFpSpread.TextTipPolicy = TextTipPolicy.Floating;
		winFpSpread.EditModeOn += winFpSpread_EditModeOn;
		winFpSpread.EditModeOff += winFpSpread_EditModeOff;
		winFpSpread.DialogKey += winFpSpread_DialogKey;
		winFpSpread.SheetTabClick += winFpSpread_SheetTabClick;
		winFpSpread.Change += winFpSpread_CellChange;
		winFpSpread.LeaveCell += winFpSpread_LeaveCell;
		winFpSpread.EnterCell += winFpSpread_EnterCell;
		winFpSpread.CellClick += winFpSpread_CellClick;
		winFpSpread.ButtonClicked += winFpSpread_ButtonClicked;
		winFpSpread.ComboCloseUp += winFpSpread_ComboCloseUp;
		winFpSpread.ComboDropDown += winFpSpread_ComboCloseDown;
		winFpSpread.EditChange += winFpSpread_editchange;
		winFpSpread.TopChange += winFpSpread_TopChange;
		((Control)winFpSpread).KeyPress += new KeyPressEventHandler(winFpSpread_KeyPress);
		((Control)winFpSpread).MouseHover += winFpSpread_MouseHover;
		((Control)winFpSpread).MouseMove += new MouseEventHandler(winFpSpread_MouseMove);
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
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerFooterDefaultEnhanced";
		winFpSpread_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced";
		winFpSpread_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.DefaultStyle.BackColor = Color.White;
		winFpSpread_Sheet1.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.DefaultStyle.Parent = "DataAreaDefault";
		winFpSpread_Sheet1.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefaultEnhanced";
		winFpSpread_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
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
		winFpSpread_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefaultEnhanced";
		winFpSpread_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.SelectionBackColor = SystemColors.Control;
		winFpSpread_Sheet1.SelectionForeColor = SystemColors.Control;
		winFpSpread_Sheet1.SheetCornerStyle.HorizontalAlignment = CellHorizontalAlignment.General;
		winFpSpread_Sheet1.SheetCornerStyle.NoteIndicatorColor = Color.Red;
		winFpSpread_Sheet1.SheetCornerStyle.Parent = "CornerDefaultEnhanced";
		winFpSpread_Sheet1.SheetCornerStyle.VerticalAlignment = CellVerticalAlignment.General;
		winFpSpread_Sheet1.ReferenceStyle = ReferenceStyle.A1;
		((Control)pnWinSpread).Anchor = (AnchorStyles)15;
		((Control)pnWinSpread).BackgroundImageLayout = (ImageLayout)3;
		pnWinSpread.BorderStyle = (BorderStyle)2;
		((Control)pnWinSpread).Controls.Add((Control)(object)winFpSpread);
		((Control)pnWinSpread).Location = new Point(12, 41);
		((Control)pnWinSpread).Name = "pnWinSpread";
		((Control)pnWinSpread).Size = new Size(1281, 505);
		((Control)pnWinSpread).TabIndex = 26;
		((Control)pnTimKiem).Controls.Add((Control)(object)cbbTimKiem);
		((Control)pnTimKiem).Controls.Add((Control)(object)txtTimKiem);
		((Control)pnTimKiem).Controls.Add((Control)(object)btnTimKiem);
		((Control)pnTimKiem).Location = new Point(12, 5);
		((Control)pnTimKiem).Name = "pnTimKiem";
		((Control)pnTimKiem).Size = new Size(408, 30);
		((Control)pnTimKiem).TabIndex = 27;
		cbbTimKiem.DropDownStyle = (ComboBoxStyle)2;
		((ListControl)cbbTimKiem).FormattingEnabled = true;
		((Control)cbbTimKiem).Location = new Point(6, 4);
		((Control)cbbTimKiem).Name = "cbbTimKiem";
		((Control)cbbTimKiem).Size = new Size(146, 21);
		((Control)cbbTimKiem).TabIndex = 2;
		((ListControl)cbbTimKiem).SelectedValueChanged += cbbTimKiem_SelectedValueChanged;
		((Control)txtTimKiem).Location = new Point(156, 4);
		((Control)txtTimKiem).Name = "txtTimKiem";
		((Control)txtTimKiem).Size = new Size(164, 20);
		((Control)txtTimKiem).TabIndex = 1;
		((Control)btnTimKiem).Location = new Point(326, 3);
		((Control)btnTimKiem).Name = "btnTimKiem";
		((Control)btnTimKiem).Size = new Size(75, 23);
		((Control)btnTimKiem).TabIndex = 0;
		((Control)btnTimKiem).Text = "Tìm kiếm";
		((ButtonBase)btnTimKiem).UseVisualStyleBackColor = true;
		((Control)btnTimKiem).Click += btnTimKiem_Click;
		((Control)cmdClear).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdClear).ImageAlign = (ContentAlignment)16;
		((Control)cmdClear).Location = new Point(363, 6);
		((Control)cmdClear).Name = "cmdClear";
		((Control)cmdClear).Size = new Size(80, 25);
		((Control)cmdClear).TabIndex = 13;
		((Control)cmdClear).Text = "N&hập lại";
		((ButtonBase)cmdClear).UseVisualStyleBackColor = true;
		((Control)cmdClear).Click += cmdClear_Click;
		((Control)cmdLoadToKhai).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdLoadToKhai).ImageAlign = (ContentAlignment)16;
		((Control)cmdLoadToKhai).Location = new Point(190, 6);
		((Control)cmdLoadToKhai).Name = "cmdLoadToKhai";
		((Control)cmdLoadToKhai).Size = new Size(80, 25);
		((Control)cmdLoadToKhai).TabIndex = 12;
		((Control)cmdLoadToKhai).Text = "Tải tờ kh&ai";
		((ButtonBase)cmdLoadToKhai).UseVisualStyleBackColor = true;
		((Control)cmdLoadToKhai).Click += cmdLoadToKhai_Click;
		((Control)cmdSave).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdSave).ImageAlign = (ContentAlignment)16;
		((Control)cmdSave).Location = new Point(522, 6);
		((Control)cmdSave).Name = "cmdSave";
		((Control)cmdSave).Size = new Size(80, 25);
		((Control)cmdSave).TabIndex = 14;
		((Control)cmdSave).Text = "&Ghi";
		((ButtonBase)cmdSave).UseVisualStyleBackColor = true;
		((Control)cmdSave).Click += cmdSave_Click;
		((Control)cmdPrint).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdPrint).ImageAlign = (ContentAlignment)16;
		((Control)cmdPrint).Location = new Point(601, 6);
		((Control)cmdPrint).Name = "cmdPrint";
		((Control)cmdPrint).Size = new Size(80, 25);
		((Control)cmdPrint).TabIndex = 15;
		((Control)cmdPrint).Text = "&In";
		((ButtonBase)cmdPrint).UseVisualStyleBackColor = true;
		((Control)cmdPrint).Click += cmdPrint_Click;
		((Control)cmdDelete).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdDelete).ImageAlign = (ContentAlignment)16;
		((Control)cmdDelete).Location = new Point(659, 6);
		((Control)cmdDelete).Name = "cmdDelete";
		((Control)cmdDelete).Size = new Size(80, 25);
		((Control)cmdDelete).TabIndex = 11;
		((Control)cmdDelete).Text = "&Xóa";
		((ButtonBase)cmdDelete).UseVisualStyleBackColor = true;
		((Control)cmdDelete).Click += cmdDelete_Click;
		((Control)cmdKiemTra).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdKiemTra).ImageAlign = (ContentAlignment)16;
		((Control)cmdKiemTra).Location = new Point(442, 6);
		((Control)cmdKiemTra).Name = "cmdKiemTra";
		((Control)cmdKiemTra).Size = new Size(80, 25);
		((Control)cmdKiemTra).TabIndex = 16;
		((Control)cmdKiemTra).Text = "Kiểm t&ra";
		((ButtonBase)cmdKiemTra).UseVisualStyleBackColor = true;
		((Control)cmdKiemTra).Click += cmdKiemTra_Click;
		((Control)btnThemPhuLuc).Anchor = (AnchorStyles)10;
		((ButtonBase)btnThemPhuLuc).ImageAlign = (ContentAlignment)16;
		((Control)btnThemPhuLuc).Location = new Point(91, 6);
		((Control)btnThemPhuLuc).Name = "btnThemPhuLuc";
		((Control)btnThemPhuLuc).Size = new Size(80, 25);
		((Control)btnThemPhuLuc).TabIndex = 24;
		((Control)btnThemPhuLuc).Text = "&Thêm phụ lục";
		((ButtonBase)btnThemPhuLuc).UseVisualStyleBackColor = true;
		((Control)btnThemPhuLuc).Click += btnThemPhuLuc_Click;
		((Control)cmdExport).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdExport).ImageAlign = (ContentAlignment)16;
		((Control)cmdExport).Location = new Point(724, 6);
		((Control)cmdExport).Name = "cmdExport";
		((Control)cmdExport).Size = new Size(80, 25);
		((Control)cmdExport).TabIndex = 17;
		((Control)cmdExport).Text = "&Kết xuất";
		((ButtonBase)cmdExport).UseVisualStyleBackColor = true;
		((Control)cmdExport).Click += cmdExport_Click;
		((Control)cmdImportXML).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdImportXML).ImageAlign = (ContentAlignment)16;
		((Control)cmdImportXML).Location = new Point(804, 6);
		((Control)cmdImportXML).Name = "cmdImportXML";
		((Control)cmdImportXML).Size = new Size(80, 25);
		((Control)cmdImportXML).TabIndex = 20;
		((Control)cmdImportXML).Text = "Nhậ&p từ XML";
		((ButtonBase)cmdImportXML).UseVisualStyleBackColor = true;
		((Control)cmdImportXML).Click += cmdImportXML_Click;
		((Control)cmdExit).Anchor = (AnchorStyles)10;
		((ButtonBase)cmdExit).ImageAlign = (ContentAlignment)16;
		((Control)cmdExit).Location = new Point(902, 6);
		((Control)cmdExit).Name = "cmdExit";
		((Control)cmdExit).Size = new Size(80, 25);
		((Control)cmdExit).TabIndex = 18;
		((Control)cmdExit).Text = "Đó&ng";
		((ButtonBase)cmdExit).UseVisualStyleBackColor = true;
		((Control)cmdExit).Click += cmdExit_Click;
		((Control)btnTaiBangKe).Anchor = (AnchorStyles)10;
		((ButtonBase)btnTaiBangKe).ImageAlign = (ContentAlignment)16;
		((Control)btnTaiBangKe).Location = new Point(275, 6);
		((Control)btnTaiBangKe).Name = "btnTaiBangKe";
		((Control)btnTaiBangKe).Size = new Size(80, 25);
		((Control)btnTaiBangKe).TabIndex = 28;
		((Control)btnTaiBangKe).Text = "Tải &bảng kê";
		((ButtonBase)btnTaiBangKe).UseVisualStyleBackColor = true;
		((Control)btnTaiBangKe).Click += btnTaiBangKe_Click;
		((Control)btnTongHopKHBS).Anchor = (AnchorStyles)10;
		((ButtonBase)btnTongHopKHBS).ImageAlign = (ContentAlignment)16;
		((Control)btnTongHopKHBS).Location = new Point(152, 6);
		((Control)btnTongHopKHBS).Name = "btnTongHopKHBS";
		((Control)btnTongHopKHBS).Size = new Size(104, 25);
		((Control)btnTongHopKHBS).TabIndex = 29;
		((Control)btnTongHopKHBS).Text = "Tổng hợp KHB&S";
		((ButtonBase)btnTongHopKHBS).UseVisualStyleBackColor = true;
		((Control)btnTongHopKHBS).Click += btnTongHopKHBS_Click;
		((Control)pnButon).Anchor = (AnchorStyles)10;
		((Control)pnButon).Controls.Add((Control)(object)btnCustom);
		((Control)pnButon).Controls.Add((Control)(object)btnTongHopKHBS);
		((Control)pnButon).Controls.Add((Control)(object)btnTaiBangKe);
		((Control)pnButon).Controls.Add((Control)(object)cmdExit);
		((Control)pnButon).Controls.Add((Control)(object)cmdImportXML);
		((Control)pnButon).Controls.Add((Control)(object)cmdExport);
		((Control)pnButon).Controls.Add((Control)(object)btnThemPhuLuc);
		((Control)pnButon).Controls.Add((Control)(object)cmdKiemTra);
		((Control)pnButon).Controls.Add((Control)(object)cmdDelete);
		((Control)pnButon).Controls.Add((Control)(object)cmdPrint);
		((Control)pnButon).Controls.Add((Control)(object)cmdSave);
		((Control)pnButon).Controls.Add((Control)(object)cmdLoadToKhai);
		((Control)pnButon).Controls.Add((Control)(object)cmdClear);
		((Control)pnButon).Location = new Point(296, 549);
		((Control)pnButon).Name = "pnButon";
		((Control)pnButon).Size = new Size(984, 36);
		((Control)pnButon).TabIndex = 29;
		((Control)btnCustom).Anchor = (AnchorStyles)10;
		((ButtonBase)btnCustom).ImageAlign = (ContentAlignment)16;
		((Control)btnCustom).Location = new Point(33, 6);
		((Control)btnCustom).Name = "btnCustom";
		((Control)btnCustom).Size = new Size(138, 25);
		((Control)btnCustom).TabIndex = 30;
		((Control)btnCustom).Text = "&Chuyển sang tờ khai năm";
		((ButtonBase)btnCustom).UseVisualStyleBackColor = true;
		((Control)btnCustom).Visible = false;
		((Control)btnCustom).Click += btn_custom_Click;
		((Control)pnF7).Anchor = (AnchorStyles)10;
		((Control)pnF7).Controls.Add((Control)(object)label1);
		((Control)pnF7).Controls.Add((Control)(object)label2);
		((Control)pnF7).Location = new Point(96, 1);
		((Control)pnF7).Name = "pnF7";
		((Control)pnF7).Size = new Size(166, 35);
		((Control)pnF7).TabIndex = 40;
		((Control)label1).Anchor = (AnchorStyles)6;
		((Control)label1).AutoSize = true;
		((Control)label1).Location = new Point(4, 19);
		((Control)label1).Name = "label1";
		((Control)label1).Size = new Size(148, 13);
		((Control)label1).TabIndex = 39;
		((Control)label1).Text = "F8: Xóa bảng dữ liệu/ phụ lục";
		((Control)label2).Anchor = (AnchorStyles)6;
		((Control)label2).AutoSize = true;
		((Control)label2).Location = new Point(4, 3);
		((Control)label2).Name = "label2";
		((Control)label2).Size = new Size(156, 13);
		((Control)label2).TabIndex = 38;
		((Control)label2).Text = "F7: Thêm bảng dữ liệu/ phụ lục";
		((Control)lblInfo).Location = new Point(0, 0);
		((Control)lblInfo).Name = "lblInfo";
		((Control)lblInfo).Size = new Size(100, 23);
		((Control)lblInfo).TabIndex = 0;
		((Control)panelChild).Location = new Point(137, 587);
		((Control)panelChild).Name = "panelChild";
		((Control)panelChild).Size = new Size(17, 10);
		((Control)panelChild).TabIndex = 26;
		((Control)pnF5).Anchor = (AnchorStyles)10;
		((Control)pnF5).Controls.Add((Control)(object)lblThemDong);
		((Control)pnF5).Controls.Add((Control)(object)lblXoaDong);
		((Control)pnF5).Location = new Point(3, 1);
		((Control)pnF5).Name = "pnF5";
		((Control)pnF5).Size = new Size(91, 33);
		((Control)pnF5).TabIndex = 41;
		((Control)lblThemDong).Anchor = (AnchorStyles)6;
		((Control)lblThemDong).AutoSize = true;
		((Control)lblThemDong).Location = new Point(5, 0);
		((Control)lblThemDong).Name = "lblThemDong";
		((Control)lblThemDong).Size = new Size(79, 13);
		((Control)lblThemDong).TabIndex = 37;
		((Control)lblThemDong).Text = "F5: Thêm dòng";
		((Control)lblXoaDong).Anchor = (AnchorStyles)6;
		((Control)lblXoaDong).AutoSize = true;
		((Control)lblXoaDong).Location = new Point(5, 17);
		((Control)lblXoaDong).Name = "lblXoaDong";
		((Control)lblXoaDong).Size = new Size(71, 13);
		((Control)lblXoaDong).TabIndex = 36;
		((Control)lblXoaDong).Text = "F6: Xóa dòng";
		((Control)pnInsert).Controls.Add((Control)(object)pnF5);
		((Control)pnInsert).Controls.Add((Control)(object)pnF7);
		((Control)pnInsert).Location = new Point(14, 549);
		((Control)pnInsert).Name = "pnInsert";
		((Control)pnInsert).Size = new Size(262, 36);
		((Control)pnInsert).TabIndex = 42;
		((Control)panelCQT).Controls.Add((Control)(object)cbbCQTChiCuc);
		((Control)panelCQT).Controls.Add((Control)(object)label3);
		((Control)panelCQT).Controls.Add((Control)(object)lbCQTCapCuc);
		((Control)panelCQT).Controls.Add((Control)(object)cbbCQTCuc);
		((Control)panelCQT).Location = new Point(445, 5);
		((Control)panelCQT).Name = "panelCQT";
		((Control)panelCQT).Size = new Size(824, 27);
		((Control)panelCQT).TabIndex = 43;
		cbbCQTChiCuc.DropDownStyle = (ComboBoxStyle)2;
		((ListControl)cbbCQTChiCuc).FormattingEnabled = true;
		((Control)cbbCQTChiCuc).Location = new Point(517, 4);
		((Control)cbbCQTChiCuc).Name = "cbbCQTChiCuc";
		((Control)cbbCQTChiCuc).Size = new Size(300, 21);
		((Control)cbbCQTChiCuc).TabIndex = 3;
		((Control)label3).AutoSize = true;
		((Control)label3).Location = new Point(422, 7);
		((Control)label3).Name = "label3";
		((Control)label3).Size = new Size(89, 13);
		((Control)label3).TabIndex = 2;
		((Control)label3).Text = "CQT cấp chi Cục";
		((Control)lbCQTCapCuc).AutoSize = true;
		((Control)lbCQTCapCuc).Location = new Point(4, 5);
		((Control)lbCQTCapCuc).Name = "lbCQTCapCuc";
		((Control)lbCQTCapCuc).Size = new Size(72, 13);
		((Control)lbCQTCapCuc).TabIndex = 1;
		((Control)lbCQTCapCuc).Text = "CQT cấp Cục";
		cbbCQTCuc.DropDownStyle = (ComboBoxStyle)2;
		((ListControl)cbbCQTCuc).FormattingEnabled = true;
		((Control)cbbCQTCuc).Location = new Point(82, 3);
		((Control)cbbCQTCuc).Name = "cbbCQTCuc";
		((Control)cbbCQTCuc).Size = new Size(334, 21);
		((Control)cbbCQTCuc).TabIndex = 0;
		((ListControl)cbbCQTCuc).SelectedValueChanged += cbbCQTCuc_SelectedValueChanged;
		((Control)panelPage).Controls.Add((Control)(object)btnEndPage);
		((Control)panelPage).Controls.Add((Control)(object)btnNextPage);
		((Control)panelPage).Controls.Add((Control)(object)labelTotalPage);
		((Control)panelPage).Controls.Add((Control)(object)label4);
		((Control)panelPage).Controls.Add((Control)(object)txtPage);
		((Control)panelPage).Controls.Add((Control)(object)btnPrivPage);
		((Control)panelPage).Controls.Add((Control)(object)btnFirstPage);
		((Control)panelPage).Location = new Point(434, 5);
		((Control)panelPage).Name = "panelPage";
		((Control)panelPage).Size = new Size(305, 30);
		((Control)panelPage).TabIndex = 26;
		((Control)panelPage).Visible = false;
		((Control)btnFirstPage).Location = new Point(14, 5);
		((Control)btnFirstPage).Name = "btnFirstPage";
		((Control)btnFirstPage).Size = new Size(30, 23);
		((Control)btnFirstPage).TabIndex = 0;
		((Control)btnFirstPage).Text = "<<";
		((ButtonBase)btnFirstPage).UseVisualStyleBackColor = true;
		((Control)btnFirstPage).Click += btnFirstPage_Click;
		((Control)btnPrivPage).Location = new Point(50, 5);
		((Control)btnPrivPage).Name = "btnPrivPage";
		((Control)btnPrivPage).Size = new Size(30, 23);
		((Control)btnPrivPage).TabIndex = 1;
		((Control)btnPrivPage).Text = "<";
		((ButtonBase)btnPrivPage).UseVisualStyleBackColor = true;
		((Control)btnPrivPage).Click += btnPrivPage_Click;
		((Control)txtPage).Location = new Point(86, 6);
		((Control)txtPage).Name = "txtPage";
		((Control)txtPage).Size = new Size(54, 20);
		((Control)txtPage).TabIndex = 2;
		((Control)txtPage).Text = "1";
		txtPage.TextAlign = (HorizontalAlignment)2;
		((Control)txtPage).TextChanged += txtPage_TextChanged;
		((Control)label4).AutoSize = true;
		((Control)label4).Location = new Point(146, 10);
		((Control)label4).Name = "label4";
		((Control)label4).Size = new Size(12, 13);
		((Control)label4).TabIndex = 3;
		((Control)label4).Text = "/";
		((Control)labelTotalPage).AutoSize = true;
		((Control)labelTotalPage).Location = new Point(164, 9);
		((Control)labelTotalPage).Name = "labelTotalPage";
		((Control)labelTotalPage).Size = new Size(13, 13);
		((Control)labelTotalPage).TabIndex = 4;
		((Control)labelTotalPage).Text = "1";
		((Control)btnNextPage).Location = new Point(221, 5);
		((Control)btnNextPage).Name = "btnNextPage";
		((Control)btnNextPage).Size = new Size(30, 23);
		((Control)btnNextPage).TabIndex = 5;
		((Control)btnNextPage).Text = ">";
		((ButtonBase)btnNextPage).UseVisualStyleBackColor = true;
		((Control)btnNextPage).Click += btnNextPage_Click;
		((Control)btnEndPage).Location = new Point(257, 5);
		((Control)btnEndPage).Name = "btnEndPage";
		((Control)btnEndPage).Size = new Size(30, 23);
		((Control)btnEndPage).TabIndex = 6;
		((Control)btnEndPage).Text = ">>";
		((ButtonBase)btnEndPage).UseVisualStyleBackColor = true;
		((Control)btnEndPage).Click += btnEndPage_Click;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Form)this).ClientSize = new Size(1284, 587);
		((Control)this).Controls.Add((Control)(object)panelPage);
		((Control)this).Controls.Add((Control)(object)panelCQT);
		((Control)this).Controls.Add((Control)(object)panelChild);
		((Control)this).Controls.Add((Control)(object)pnInsert);
		((Control)this).Controls.Add((Control)(object)pnButon);
		((Control)this).Controls.Add((Control)(object)pnTimKiem);
		((Control)this).Controls.Add((Control)(object)pnWinSpread);
		((Form)this).Icon = (Icon)componentResourceManager.GetObject("$this.Icon");
		((Form)this).KeyPreview = true;
		((Control)this).Name = "frmWinToKhai";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Hỗ trợ kê khai";
		((Form)this).WindowState = (FormWindowState)2;
		((Form)this).FormClosing += new FormClosingEventHandler(frmWinToKhai_FormClosing);
		((Form)this).FormClosed += new FormClosedEventHandler(frmWinToKhai_FormClosed);
		((Form)this).Load += frmToKhai_Load;
		((Control)this).KeyDown += new KeyEventHandler(frmWinToKhai_KeyDown);
		((Control)this).KeyPress += new KeyPressEventHandler(frmWinToKhai_KeyPress);
		((Control)this).Resize += frmWinToKhai_Resize;
		((ISupportInitialize)winFpSpread).EndInit();
		((ISupportInitialize)winFpSpread_Sheet1).EndInit();
		((Control)pnWinSpread).ResumeLayout(false);
		((Control)pnTimKiem).ResumeLayout(false);
		((Control)pnTimKiem).PerformLayout();
		((Control)pnButon).ResumeLayout(false);
		((Control)pnF7).ResumeLayout(false);
		((Control)pnF7).PerformLayout();
		((Control)pnF5).ResumeLayout(false);
		((Control)pnF5).PerformLayout();
		((Control)pnInsert).ResumeLayout(false);
		((Control)panelCQT).ResumeLayout(false);
		((Control)panelCQT).PerformLayout();
		((Control)panelPage).ResumeLayout(false);
		((Control)panelPage).PerformLayout();
		((Control)this).ResumeLayout(false);
	}
}
