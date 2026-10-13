using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace CoolPrintPreview;

internal class CoolPrintPreviewDialog : Form
{
	private PrintDocument _doc;

	private IContainer components = null;

	private ToolStrip _toolStrip;

	private ToolStripButton _btnPrint;

	private ToolStripButton _btnPageSetup;

	private CoolPrintPreviewControl _preview;

	private ToolStripSplitButton _btnZoom;

	private ToolStripMenuItem _itemActualSize;

	private ToolStripMenuItem _itemFullPage;

	private ToolStripMenuItem _itemPageWidth;

	private ToolStripMenuItem _itemTwoPages;

	private ToolStripSeparator toolStripMenuItem1;

	private ToolStripMenuItem _item500;

	private ToolStripMenuItem _item200;

	private ToolStripMenuItem _item150;

	private ToolStripMenuItem _item100;

	private ToolStripMenuItem _item75;

	private ToolStripMenuItem _item50;

	private ToolStripMenuItem _item25;

	private ToolStripMenuItem _item10;

	private ToolStripButton _btnFirst;

	private ToolStripButton _btnPrev;

	private ToolStripTextBox _txtStartPage;

	private ToolStripLabel _lblPageCount;

	private ToolStripButton _btnNext;

	private ToolStripButton _btnLast;

	private ToolStripSeparator _separator;

	private ToolStripButton _btnCancel;

	private ToolStripSeparator toolStripSeparator2;

	public PrintDocument Document
	{
		get
		{
			return _doc;
		}
		set
		{
			//IL_001c: Unknown result type (might be due to invalid IL or missing references)
			//IL_0026: Expected O, but got Unknown
			//IL_0034: Unknown result type (might be due to invalid IL or missing references)
			//IL_003e: Expected O, but got Unknown
			//IL_0062: Unknown result type (might be due to invalid IL or missing references)
			//IL_006c: Expected O, but got Unknown
			//IL_007a: Unknown result type (might be due to invalid IL or missing references)
			//IL_0084: Expected O, but got Unknown
			if (_doc != null)
			{
				_doc.BeginPrint -= new PrintEventHandler(_doc_BeginPrint);
				_doc.EndPrint -= new PrintEventHandler(_doc_EndPrint);
			}
			_doc = value;
			if (_doc != null)
			{
				_doc.BeginPrint += new PrintEventHandler(_doc_BeginPrint);
				_doc.EndPrint += new PrintEventHandler(_doc_EndPrint);
			}
			if (((Control)this).Visible)
			{
				_preview.Document = Document;
			}
		}
	}

	public CoolPrintPreviewDialog()
		: this(null)
	{
	}

	public CoolPrintPreviewDialog(Control parentForm)
	{
		InitializeComponent();
		if (parentForm != null)
		{
			((Form)this).Size = parentForm.Size;
		}
	}

	protected override void OnShown(EventArgs e)
	{
		((Form)this).OnShown(e);
		_preview.Document = Document;
	}

	protected override void OnFormClosing(FormClosingEventArgs e)
	{
		((Form)this).OnFormClosing(e);
		if (_preview.IsRendering && !((CancelEventArgs)(object)e).Cancel)
		{
			_preview.Cancel();
		}
	}

	private void _btnPrint_Click(object sender, EventArgs e)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0064: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Invalid comparison between Unknown and I4
		PrintDialog val = new PrintDialog();
		try
		{
			val.AllowSomePages = true;
			val.AllowSelection = true;
			val.UseEXDialog = true;
			val.Document = Document;
			PrinterSettings printerSettings = val.PrinterSettings;
			int minimumPage = (printerSettings.FromPage = 1);
			printerSettings.MinimumPage = minimumPage;
			minimumPage = (printerSettings.ToPage = _preview.PageCount);
			printerSettings.MaximumPage = minimumPage;
			if ((int)((CommonDialog)val).ShowDialog((IWin32Window)(object)this) == 1)
			{
				_preview.Print();
			}
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
	}

	private void _btnPageSetup_Click(object sender, EventArgs e)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Invalid comparison between Unknown and I4
		PageSetupDialog val = new PageSetupDialog();
		try
		{
			val.Document = Document;
			if ((int)((CommonDialog)val).ShowDialog((IWin32Window)(object)this) == 1)
			{
				_preview.RefreshPreview();
			}
		}
		finally
		{
			((IDisposable)val)?.Dispose();
		}
	}

	private void _btnZoom_ButtonClick(object sender, EventArgs e)
	{
		_preview.ZoomMode = ((_preview.ZoomMode == ZoomMode.ActualSize) ? ZoomMode.FullPage : ZoomMode.ActualSize);
	}

	private void _btnZoom_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
	{
		if ((object)e.ClickedItem == _itemActualSize)
		{
			_preview.ZoomMode = ZoomMode.ActualSize;
		}
		else if ((object)e.ClickedItem == _itemFullPage)
		{
			_preview.ZoomMode = ZoomMode.FullPage;
		}
		else if ((object)e.ClickedItem == _itemPageWidth)
		{
			_preview.ZoomMode = ZoomMode.PageWidth;
		}
		else if ((object)e.ClickedItem == _itemTwoPages)
		{
			_preview.ZoomMode = ZoomMode.TwoPages;
		}
		if ((object)e.ClickedItem == _item10)
		{
			_preview.Zoom = 0.1;
		}
		else if ((object)e.ClickedItem == _item100)
		{
			_preview.Zoom = 1.0;
		}
		else if ((object)e.ClickedItem == _item150)
		{
			_preview.Zoom = 1.5;
		}
		else if ((object)e.ClickedItem == _item200)
		{
			_preview.Zoom = 2.0;
		}
		else if ((object)e.ClickedItem == _item25)
		{
			_preview.Zoom = 0.25;
		}
		else if ((object)e.ClickedItem == _item50)
		{
			_preview.Zoom = 0.5;
		}
		else if ((object)e.ClickedItem == _item500)
		{
			_preview.Zoom = 5.0;
		}
		else if ((object)e.ClickedItem == _item75)
		{
			_preview.Zoom = 0.75;
		}
	}

	private void _btnFirst_Click(object sender, EventArgs e)
	{
		_preview.StartPage = 0;
	}

	private void _btnPrev_Click(object sender, EventArgs e)
	{
		_preview.StartPage--;
	}

	private void _btnNext_Click(object sender, EventArgs e)
	{
		_preview.StartPage++;
	}

	private void _btnLast_Click(object sender, EventArgs e)
	{
		_preview.StartPage = _preview.PageCount - 1;
	}

	private void _txtStartPage_Enter(object sender, EventArgs e)
	{
		_txtStartPage.SelectAll();
	}

	private void _txtStartPage_Validating(object sender, CancelEventArgs e)
	{
		CommitPageNumber();
	}

	private void _txtStartPage_KeyPress(object sender, KeyPressEventArgs e)
	{
		char keyChar = e.KeyChar;
		if (keyChar == '\r')
		{
			CommitPageNumber();
			e.Handled = true;
		}
		else if (keyChar > ' ' && !char.IsDigit(keyChar))
		{
			e.Handled = true;
		}
	}

	private void CommitPageNumber()
	{
		if (int.TryParse(((ToolStripItem)_txtStartPage).Text, out var result))
		{
			_preview.StartPage = result - 1;
		}
	}

	private void _preview_StartPageChanged(object sender, EventArgs e)
	{
		int num = _preview.StartPage + 1;
		((ToolStripItem)_txtStartPage).Text = num.ToString();
	}

	private void _preview_PageCountChanged(object sender, EventArgs e)
	{
		((Control)this).Update();
		Application.DoEvents();
		((ToolStripItem)_lblPageCount).Text = $"of {_preview.PageCount}";
	}

	private void _btnCancel_Click(object sender, EventArgs e)
	{
		if (_preview.IsRendering)
		{
			_preview.Cancel();
		}
		else
		{
			((Form)this).Close();
		}
	}

	private void _doc_BeginPrint(object sender, PrintEventArgs e)
	{
		((ToolStripItem)_btnCancel).Text = "&Cancel";
		ToolStripButton btnPrint = _btnPrint;
		bool enabled = (((ToolStripItem)_btnPageSetup).Enabled = false);
		((ToolStripItem)btnPrint).Enabled = enabled;
	}

	private void _doc_EndPrint(object sender, PrintEventArgs e)
	{
		((ToolStripItem)_btnCancel).Text = "&Close";
		ToolStripButton btnPrint = _btnPrint;
		bool enabled = (((ToolStripItem)_btnPageSetup).Enabled = true);
		((ToolStripItem)btnPrint).Enabled = enabled;
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
		//IL_025b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0265: Expected O, but got Unknown
		//IL_02f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_02fb: Expected O, but got Unknown
		//IL_043f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0449: Expected O, but got Unknown
		//IL_04b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_04c1: Expected O, but got Unknown
		//IL_04d3: Unknown result type (might be due to invalid IL or missing references)
		//IL_04dd: Expected O, but got Unknown
		//IL_0529: Unknown result type (might be due to invalid IL or missing references)
		//IL_0533: Expected O, but got Unknown
		//IL_057f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0589: Expected O, but got Unknown
		//IL_05d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_05df: Expected O, but got Unknown
		//IL_0830: Unknown result type (might be due to invalid IL or missing references)
		//IL_083a: Expected O, but got Unknown
		//IL_08b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_08c3: Expected O, but got Unknown
		//IL_0989: Unknown result type (might be due to invalid IL or missing references)
		//IL_0993: Expected O, but got Unknown
		//IL_0a01: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a0b: Expected O, but got Unknown
		//IL_0a8a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a94: Expected O, but got Unknown
		//IL_0b45: Unknown result type (might be due to invalid IL or missing references)
		//IL_0b4f: Expected O, but got Unknown
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(CoolPrintPreviewDialog));
		_toolStrip = new ToolStrip();
		_btnPrint = new ToolStripButton();
		_btnPageSetup = new ToolStripButton();
		toolStripSeparator2 = new ToolStripSeparator();
		_btnZoom = new ToolStripSplitButton();
		_itemActualSize = new ToolStripMenuItem();
		_itemFullPage = new ToolStripMenuItem();
		_itemPageWidth = new ToolStripMenuItem();
		_itemTwoPages = new ToolStripMenuItem();
		toolStripMenuItem1 = new ToolStripSeparator();
		_item500 = new ToolStripMenuItem();
		_item200 = new ToolStripMenuItem();
		_item150 = new ToolStripMenuItem();
		_item100 = new ToolStripMenuItem();
		_item75 = new ToolStripMenuItem();
		_item50 = new ToolStripMenuItem();
		_item25 = new ToolStripMenuItem();
		_item10 = new ToolStripMenuItem();
		_btnFirst = new ToolStripButton();
		_btnPrev = new ToolStripButton();
		_txtStartPage = new ToolStripTextBox();
		_lblPageCount = new ToolStripLabel();
		_btnNext = new ToolStripButton();
		_btnLast = new ToolStripButton();
		_separator = new ToolStripSeparator();
		_btnCancel = new ToolStripButton();
		_preview = new CoolPrintPreviewControl();
		((Control)_toolStrip).SuspendLayout();
		((Control)this).SuspendLayout();
		_toolStrip.GripStyle = (ToolStripGripStyle)0;
		_toolStrip.Items.AddRange((ToolStripItem[])(object)new ToolStripItem[12]
		{
			(ToolStripItem)_btnPrint,
			(ToolStripItem)_btnPageSetup,
			(ToolStripItem)toolStripSeparator2,
			(ToolStripItem)_btnZoom,
			(ToolStripItem)_btnFirst,
			(ToolStripItem)_btnPrev,
			(ToolStripItem)_txtStartPage,
			(ToolStripItem)_lblPageCount,
			(ToolStripItem)_btnNext,
			(ToolStripItem)_btnLast,
			(ToolStripItem)_separator,
			(ToolStripItem)_btnCancel
		});
		((Control)_toolStrip).Location = new Point(0, 0);
		((Control)_toolStrip).Name = "_toolStrip";
		((Control)_toolStrip).Size = new Size(426, 25);
		((Control)_toolStrip).TabIndex = 0;
		((Control)_toolStrip).Text = "toolStrip1";
		((ToolStripItem)_btnPrint).DisplayStyle = (ToolStripItemDisplayStyle)2;
		((ToolStripItem)_btnPrint).Image = (Image)componentResourceManager.GetObject("_btnPrint.Image");
		((ToolStripItem)_btnPrint).ImageTransparentColor = Color.Magenta;
		((ToolStripItem)_btnPrint).Name = "_btnPrint";
		((ToolStripItem)_btnPrint).Size = new Size(23, 22);
		((ToolStripItem)_btnPrint).Text = "Print Document";
		((ToolStripItem)_btnPrint).Visible = false;
		((ToolStripItem)_btnPrint).Click += _btnPrint_Click;
		((ToolStripItem)_btnPageSetup).DisplayStyle = (ToolStripItemDisplayStyle)2;
		((ToolStripItem)_btnPageSetup).Image = (Image)componentResourceManager.GetObject("_btnPageSetup.Image");
		((ToolStripItem)_btnPageSetup).ImageTransparentColor = Color.Magenta;
		((ToolStripItem)_btnPageSetup).Name = "_btnPageSetup";
		((ToolStripItem)_btnPageSetup).Size = new Size(23, 22);
		((ToolStripItem)_btnPageSetup).Text = "Page Setup";
		((ToolStripItem)_btnPageSetup).Visible = false;
		((ToolStripItem)_btnPageSetup).Click += _btnPageSetup_Click;
		((ToolStripItem)toolStripSeparator2).Name = "toolStripSeparator2";
		((ToolStripItem)toolStripSeparator2).Size = new Size(6, 25);
		_btnZoom.AutoToolTip = false;
		((ToolStripDropDownItem)_btnZoom).DropDownItems.AddRange((ToolStripItem[])(object)new ToolStripItem[13]
		{
			(ToolStripItem)_itemActualSize,
			(ToolStripItem)_itemFullPage,
			(ToolStripItem)_itemPageWidth,
			(ToolStripItem)_itemTwoPages,
			(ToolStripItem)toolStripMenuItem1,
			(ToolStripItem)_item500,
			(ToolStripItem)_item200,
			(ToolStripItem)_item150,
			(ToolStripItem)_item100,
			(ToolStripItem)_item75,
			(ToolStripItem)_item50,
			(ToolStripItem)_item25,
			(ToolStripItem)_item10
		});
		((ToolStripItem)_btnZoom).Image = (Image)componentResourceManager.GetObject("_btnZoom.Image");
		((ToolStripItem)_btnZoom).ImageTransparentColor = Color.Magenta;
		((ToolStripItem)_btnZoom).Name = "_btnZoom";
		((ToolStripItem)_btnZoom).Size = new Size(71, 22);
		((ToolStripItem)_btnZoom).Text = "&Zoom";
		_btnZoom.ButtonClick += _btnZoom_ButtonClick;
		((ToolStripDropDownItem)_btnZoom).DropDownItemClicked += new ToolStripItemClickedEventHandler(_btnZoom_DropDownItemClicked);
		((ToolStripItem)_itemActualSize).Image = (Image)componentResourceManager.GetObject("_itemActualSize.Image");
		((ToolStripItem)_itemActualSize).Name = "_itemActualSize";
		((ToolStripItem)_itemActualSize).Size = new Size(135, 22);
		((ToolStripItem)_itemActualSize).Text = "Actual Size";
		((ToolStripItem)_itemFullPage).Image = (Image)componentResourceManager.GetObject("_itemFullPage.Image");
		((ToolStripItem)_itemFullPage).Name = "_itemFullPage";
		((ToolStripItem)_itemFullPage).Size = new Size(135, 22);
		((ToolStripItem)_itemFullPage).Text = "Full Page";
		((ToolStripItem)_itemPageWidth).Image = (Image)componentResourceManager.GetObject("_itemPageWidth.Image");
		((ToolStripItem)_itemPageWidth).Name = "_itemPageWidth";
		((ToolStripItem)_itemPageWidth).Size = new Size(135, 22);
		((ToolStripItem)_itemPageWidth).Text = "Page Width";
		((ToolStripItem)_itemTwoPages).Image = (Image)componentResourceManager.GetObject("_itemTwoPages.Image");
		((ToolStripItem)_itemTwoPages).Name = "_itemTwoPages";
		((ToolStripItem)_itemTwoPages).Size = new Size(135, 22);
		((ToolStripItem)_itemTwoPages).Text = "Two Pages";
		((ToolStripItem)toolStripMenuItem1).Name = "toolStripMenuItem1";
		((ToolStripItem)toolStripMenuItem1).Size = new Size(132, 6);
		((ToolStripItem)_item500).Name = "_item500";
		((ToolStripItem)_item500).Size = new Size(135, 22);
		((ToolStripItem)_item500).Text = "500%";
		((ToolStripItem)_item200).Name = "_item200";
		((ToolStripItem)_item200).Size = new Size(135, 22);
		((ToolStripItem)_item200).Text = "200%";
		((ToolStripItem)_item150).Name = "_item150";
		((ToolStripItem)_item150).Size = new Size(135, 22);
		((ToolStripItem)_item150).Text = "150%";
		((ToolStripItem)_item100).Name = "_item100";
		((ToolStripItem)_item100).Size = new Size(135, 22);
		((ToolStripItem)_item100).Text = "100%";
		((ToolStripItem)_item75).Name = "_item75";
		((ToolStripItem)_item75).Size = new Size(135, 22);
		((ToolStripItem)_item75).Text = "75%";
		((ToolStripItem)_item50).Name = "_item50";
		((ToolStripItem)_item50).Size = new Size(135, 22);
		((ToolStripItem)_item50).Text = "50%";
		((ToolStripItem)_item25).Name = "_item25";
		((ToolStripItem)_item25).Size = new Size(135, 22);
		((ToolStripItem)_item25).Text = "25%";
		((ToolStripItem)_item10).Name = "_item10";
		((ToolStripItem)_item10).Size = new Size(135, 22);
		((ToolStripItem)_item10).Text = "10%";
		((ToolStripItem)_btnFirst).DisplayStyle = (ToolStripItemDisplayStyle)2;
		((ToolStripItem)_btnFirst).Image = (Image)componentResourceManager.GetObject("_btnFirst.Image");
		((ToolStripItem)_btnFirst).ImageTransparentColor = Color.Red;
		((ToolStripItem)_btnFirst).Name = "_btnFirst";
		((ToolStripItem)_btnFirst).Size = new Size(23, 22);
		((ToolStripItem)_btnFirst).Text = "First Page";
		((ToolStripItem)_btnFirst).Click += _btnFirst_Click;
		((ToolStripItem)_btnPrev).DisplayStyle = (ToolStripItemDisplayStyle)2;
		((ToolStripItem)_btnPrev).Image = (Image)componentResourceManager.GetObject("_btnPrev.Image");
		((ToolStripItem)_btnPrev).ImageTransparentColor = Color.Red;
		((ToolStripItem)_btnPrev).Name = "_btnPrev";
		((ToolStripItem)_btnPrev).Size = new Size(23, 22);
		((ToolStripItem)_btnPrev).Text = "Previous Page";
		((ToolStripItem)_btnPrev).Click += _btnPrev_Click;
		((ToolStripItem)_txtStartPage).AutoSize = false;
		((ToolStripItem)_txtStartPage).Name = "_txtStartPage";
		((ToolStripItem)_txtStartPage).Size = new Size(32, 21);
		_txtStartPage.TextBoxTextAlign = (HorizontalAlignment)2;
		((ToolStripControlHost)_txtStartPage).Enter += _txtStartPage_Enter;
		((ToolStripControlHost)_txtStartPage).KeyPress += new KeyPressEventHandler(_txtStartPage_KeyPress);
		((ToolStripControlHost)_txtStartPage).Validating += _txtStartPage_Validating;
		((ToolStripItem)_lblPageCount).Name = "_lblPageCount";
		((ToolStripItem)_lblPageCount).Size = new Size(10, 22);
		((ToolStripItem)_lblPageCount).Text = " ";
		((ToolStripItem)_btnNext).DisplayStyle = (ToolStripItemDisplayStyle)2;
		((ToolStripItem)_btnNext).Image = (Image)componentResourceManager.GetObject("_btnNext.Image");
		((ToolStripItem)_btnNext).ImageTransparentColor = Color.Red;
		((ToolStripItem)_btnNext).Name = "_btnNext";
		((ToolStripItem)_btnNext).Size = new Size(23, 22);
		((ToolStripItem)_btnNext).Text = "Next Page";
		((ToolStripItem)_btnNext).Click += _btnNext_Click;
		((ToolStripItem)_btnLast).DisplayStyle = (ToolStripItemDisplayStyle)2;
		((ToolStripItem)_btnLast).Image = (Image)componentResourceManager.GetObject("_btnLast.Image");
		((ToolStripItem)_btnLast).ImageTransparentColor = Color.Red;
		((ToolStripItem)_btnLast).Name = "_btnLast";
		((ToolStripItem)_btnLast).Size = new Size(23, 22);
		((ToolStripItem)_btnLast).Text = "Last Page";
		((ToolStripItem)_btnLast).Click += _btnLast_Click;
		((ToolStripItem)_separator).Name = "_separator";
		((ToolStripItem)_separator).Size = new Size(6, 25);
		((ToolStripItem)_separator).Visible = false;
		_btnCancel.AutoToolTip = false;
		((ToolStripItem)_btnCancel).Image = (Image)componentResourceManager.GetObject("_btnCancel.Image");
		((ToolStripItem)_btnCancel).ImageTransparentColor = Color.Magenta;
		((ToolStripItem)_btnCancel).Name = "_btnCancel";
		((ToolStripItem)_btnCancel).Size = new Size(63, 22);
		((ToolStripItem)_btnCancel).Text = "Cancel";
		((ToolStripItem)_btnCancel).Click += _btnCancel_Click;
		((ScrollableControl)_preview).AutoScroll = true;
		((ScrollableControl)_preview).AutoScrollMinSize = new Size(0, 8);
		((Control)_preview).Dock = (DockStyle)5;
		_preview.Document = null;
		((Control)_preview).Location = new Point(0, 25);
		((Control)_preview).Name = "_preview";
		((Control)_preview).Size = new Size(426, 323);
		((Control)_preview).TabIndex = 1;
		_preview.ZoomMode = ZoomMode.PageWidth;
		_preview.StartPageChanged += _preview_StartPageChanged;
		_preview.PageCountChanged += _preview_PageCountChanged;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(96f, 96f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)2;
		((Form)this).ClientSize = new Size(426, 348);
		((Control)this).Controls.Add((Control)(object)_preview);
		((Control)this).Controls.Add((Control)(object)_toolStrip);
		((Control)this).Name = "CoolPrintPreviewDialog";
		((Form)this).ShowIcon = false;
		((Form)this).ShowInTaskbar = false;
		((Form)this).StartPosition = (FormStartPosition)4;
		((Control)this).Text = "Xem trước bản in";
		((Control)_toolStrip).ResumeLayout(false);
		((Control)_toolStrip).PerformLayout();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
