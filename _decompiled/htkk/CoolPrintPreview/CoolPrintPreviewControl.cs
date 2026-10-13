using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace CoolPrintPreview;

internal class CoolPrintPreviewControl : UserControl
{
	internal class DocumentPrinter : PrintDocument
	{
		private int _first;

		private int _last;

		private int _index;

		private PageImageList _imgList;

		public DocumentPrinter(CoolPrintPreviewControl preview, int first, int last)
		{
			_first = first;
			_last = last;
			_imgList = preview.PageImages;
			((PrintDocument)this).DefaultPageSettings = preview.Document.DefaultPageSettings;
			((PrintDocument)this).PrinterSettings = preview.Document.PrinterSettings;
		}

		protected override void OnBeginPrint(PrintEventArgs e)
		{
			_index = _first;
		}

		protected override void OnPrintPage(PrintPageEventArgs e)
		{
			e.Graphics.PageUnit = (GraphicsUnit)1;
			e.Graphics.DrawImage(_imgList[_index++], e.PageBounds);
			e.HasMorePages = _index <= _last;
		}
	}

	private const int MARGIN = 4;

	private PrintDocument _doc;

	private ZoomMode _zoomMode;

	private double _zoom;

	private int _startPage;

	private Brush _backBrush;

	private Point _ptLast;

	private PointF _himm2pix = new PointF(-1f, -1f);

	private PageImageList _img = new PageImageList();

	private bool _cancel;

	private bool _rendering;

	public PrintDocument Document
	{
		get
		{
			return _doc;
		}
		set
		{
			if (value != _doc)
			{
				_doc = value;
				RefreshPreview();
			}
		}
	}

	[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
	[Browsable(false)]
	public bool IsRendering => _rendering;

	[DefaultValue(ZoomMode.FullPage)]
	public ZoomMode ZoomMode
	{
		get
		{
			return _zoomMode;
		}
		set
		{
			if (value != _zoomMode)
			{
				_zoomMode = value;
				UpdateScrollBars();
				OnZoomModeChanged(EventArgs.Empty);
			}
		}
	}

	[Browsable(false)]
	[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
	public double Zoom
	{
		get
		{
			return _zoom;
		}
		set
		{
			if (value != _zoom || ZoomMode != ZoomMode.Custom)
			{
				ZoomMode = ZoomMode.Custom;
				_zoom = value;
				UpdateScrollBars();
				OnZoomModeChanged(EventArgs.Empty);
			}
		}
	}

	[Browsable(false)]
	[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
	public int StartPage
	{
		get
		{
			return _startPage;
		}
		set
		{
			if (value > PageCount - 1)
			{
				value = PageCount - 1;
			}
			if (value < 0)
			{
				value = 0;
			}
			if (value != _startPage)
			{
				_startPage = value;
				UpdateScrollBars();
				OnStartPageChanged(EventArgs.Empty);
			}
		}
	}

	[Browsable(false)]
	[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
	public int PageCount => _img.Count;

	[DefaultValue(typeof(Color), "AppWorkspace")]
	public override Color BackColor
	{
		get
		{
			return ((Control)this).BackColor;
		}
		set
		{
			//IL_000b: Unknown result type (might be due to invalid IL or missing references)
			//IL_0015: Expected O, but got Unknown
			((Control)this).BackColor = value;
			_backBrush = (Brush)new SolidBrush(value);
		}
	}

	[Browsable(false)]
	public PageImageList PageImages => _img;

	public event EventHandler StartPageChanged;

	public event EventHandler PageCountChanged;

	public event EventHandler ZoomModeChanged;

	public CoolPrintPreviewControl()
	{
		((Control)this).BackColor = SystemColors.AppWorkspace;
		ZoomMode = ZoomMode.PageWidth;
		StartPage = 0;
		((Control)this).SetStyle((ControlStyles)131072, true);
	}

	public void RefreshPreview()
	{
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00af: Expected O, but got Unknown
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c7: Expected O, but got Unknown
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Expected O, but got Unknown
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		//IL_0061: Expected O, but got Unknown
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Expected O, but got Unknown
		if (_doc != null)
		{
			_img.Clear();
			PrintController printController = _doc.PrintController;
			try
			{
				_cancel = false;
				_rendering = true;
				_doc.PrintController = (PrintController)new PreviewPrintController();
				_doc.PrintPage += new PrintPageEventHandler(_doc_PrintPage);
				_doc.EndPrint += new PrintEventHandler(_doc_EndPrint);
				_doc.Print();
			}
			finally
			{
				_cancel = false;
				_rendering = false;
				_doc.PrintPage -= new PrintPageEventHandler(_doc_PrintPage);
				_doc.EndPrint -= new PrintEventHandler(_doc_EndPrint);
				_doc.PrintController = printController;
			}
		}
		OnPageCountChanged(EventArgs.Empty);
		UpdatePreview();
		UpdateScrollBars();
	}

	public void Cancel()
	{
		_cancel = true;
	}

	public void Print()
	{
		//IL_0020: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Expected I4, but got Unknown
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0041: Invalid comparison between Unknown and I4
		PrinterSettings printerSettings = _doc.PrinterSettings;
		int num = printerSettings.MinimumPage - 1;
		int last = printerSettings.MaximumPage - 1;
		PrintRange printRange = printerSettings.PrintRange;
		switch ((int)printRange)
		{
		default:
			if ((int)printRange == 4194304)
			{
				num = (last = StartPage);
			}
			break;
		case 0:
			Document.Print();
			return;
		case 1:
			num = (last = StartPage);
			if (ZoomMode == ZoomMode.TwoPages)
			{
				last = Math.Min(num + 1, PageCount - 1);
			}
			break;
		case 2:
			num = printerSettings.FromPage - 1;
			last = printerSettings.ToPage - 1;
			break;
		}
		DocumentPrinter documentPrinter = new DocumentPrinter(this, num, last);
		((PrintDocument)documentPrinter).Print();
	}

	protected void OnStartPageChanged(EventArgs e)
	{
		if (this.StartPageChanged != null)
		{
			this.StartPageChanged(this, e);
		}
	}

	protected void OnPageCountChanged(EventArgs e)
	{
		if (this.PageCountChanged != null)
		{
			this.PageCountChanged(this, e);
		}
	}

	protected void OnZoomModeChanged(EventArgs e)
	{
		if (this.ZoomModeChanged != null)
		{
			this.ZoomModeChanged(this, e);
		}
	}

	protected override void OnPaintBackground(PaintEventArgs e)
	{
	}

	protected override void OnPaint(PaintEventArgs e)
	{
		Image image = GetImage(StartPage);
		if (image != null)
		{
			Rectangle imageRectangle = GetImageRectangle(image);
			if (imageRectangle.Width > 2 && imageRectangle.Height > 2)
			{
				imageRectangle.Offset(((ScrollableControl)this).AutoScrollPosition);
				if (_zoomMode != ZoomMode.TwoPages)
				{
					RenderPage(e.Graphics, image, imageRectangle);
				}
				else
				{
					imageRectangle.Width = (imageRectangle.Width - 4) / 2;
					RenderPage(e.Graphics, image, imageRectangle);
					image = GetImage(StartPage + 1);
					if (image != null)
					{
						imageRectangle = GetImageRectangle(image);
						imageRectangle.Width = (imageRectangle.Width - 4) / 2;
						imageRectangle.Offset(imageRectangle.Width + 4, 0);
						RenderPage(e.Graphics, image, imageRectangle);
					}
				}
			}
		}
		e.Graphics.FillRectangle(_backBrush, ((Control)this).ClientRectangle);
	}

	protected override void OnSizeChanged(EventArgs e)
	{
		UpdateScrollBars();
		((Control)this).OnSizeChanged(e);
	}

	protected override void OnMouseDown(MouseEventArgs e)
	{
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Invalid comparison between Unknown and I4
		((UserControl)this).OnMouseDown(e);
		if ((int)e.Button == 1048576 && ((ScrollableControl)this).AutoScrollMinSize != Size.Empty)
		{
			((Control)this).Cursor = Cursors.NoMove2D;
			_ptLast = new Point(e.X, e.Y);
		}
	}

	protected override void OnMouseUp(MouseEventArgs e)
	{
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Invalid comparison between Unknown and I4
		((Control)this).OnMouseUp(e);
		if ((int)e.Button == 1048576 && ((Control)this).Cursor == Cursors.NoMove2D)
		{
			((Control)this).Cursor = Cursors.Default;
		}
	}

	protected override void OnMouseMove(MouseEventArgs e)
	{
		((Control)this).OnMouseMove(e);
		if (((Control)this).Cursor == Cursors.NoMove2D)
		{
			int num = e.X - _ptLast.X;
			int num2 = e.Y - _ptLast.Y;
			if (num != 0 || num2 != 0)
			{
				Point autoScrollPosition = ((ScrollableControl)this).AutoScrollPosition;
				((ScrollableControl)this).AutoScrollPosition = new Point(-(autoScrollPosition.X + num), -(autoScrollPosition.Y + num2));
				_ptLast = new Point(e.X, e.Y);
			}
		}
	}

	protected override bool IsInputKey(Keys keyData)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Expected I4, but got Unknown
		//IL_0033: Unknown result type (might be due to invalid IL or missing references)
		switch (keyData - 33)
		{
		case 0:
		case 1:
		case 2:
		case 3:
		case 4:
		case 5:
		case 6:
		case 7:
			return true;
		default:
			return ((Control)this).IsInputKey(keyData);
		}
	}

	protected override void OnKeyDown(KeyEventArgs e)
	{
		//IL_001c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0022: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected I4, but got Unknown
		//IL_00bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00de: Expected I4, but got Unknown
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_008e: Expected I4, but got Unknown
		((Control)this).OnKeyDown(e);
		if (e.Handled)
		{
			return;
		}
		Keys keyCode = e.KeyCode;
		switch (keyCode - 33)
		{
		default:
			return;
		case 4:
		case 5:
		case 6:
		case 7:
		{
			if (ZoomMode == ZoomMode.FullPage || ZoomMode == ZoomMode.TwoPages)
			{
				keyCode = e.KeyCode;
				switch (keyCode - 37)
				{
				case 0:
				case 1:
					StartPage--;
					break;
				case 2:
				case 3:
					StartPage++;
					break;
				}
				break;
			}
			Point autoScrollPosition = ((ScrollableControl)this).AutoScrollPosition;
			keyCode = e.KeyCode;
			switch (keyCode - 37)
			{
			case 0:
				autoScrollPosition.X += 20;
				break;
			case 2:
				autoScrollPosition.X -= 20;
				break;
			case 1:
				autoScrollPosition.Y += 20;
				break;
			case 3:
				autoScrollPosition.Y -= 20;
				break;
			}
			((ScrollableControl)this).AutoScrollPosition = new Point(-autoScrollPosition.X, -autoScrollPosition.Y);
			break;
		}
		case 0:
			StartPage--;
			break;
		case 1:
			StartPage++;
			break;
		case 3:
			((ScrollableControl)this).AutoScrollPosition = Point.Empty;
			StartPage = 0;
			break;
		case 2:
			((ScrollableControl)this).AutoScrollPosition = Point.Empty;
			StartPage = PageCount - 1;
			break;
		}
		e.Handled = true;
	}

	private void _doc_PrintPage(object sender, PrintPageEventArgs e)
	{
		SyncPageImages(lastPageReady: false);
		if (_cancel)
		{
			e.Cancel = true;
		}
	}

	private void _doc_EndPrint(object sender, PrintEventArgs e)
	{
		SyncPageImages(lastPageReady: true);
	}

	private void SyncPageImages(bool lastPageReady)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		PreviewPrintController val = (PreviewPrintController)_doc.PrintController;
		if (val == null)
		{
			return;
		}
		PreviewPageInfo[] previewPageInfo = val.GetPreviewPageInfo();
		int num = (lastPageReady ? previewPageInfo.Length : (previewPageInfo.Length - 1));
		for (int i = _img.Count; i < num; i++)
		{
			Image image = previewPageInfo[i].Image;
			_img.Add(image);
			OnPageCountChanged(EventArgs.Empty);
			if (StartPage < 0)
			{
				StartPage = 0;
			}
			if (i == StartPage || i == StartPage + 1)
			{
				((Control)this).Refresh();
			}
			Application.DoEvents();
		}
	}

	private Image GetImage(int page)
	{
		return (page > -1 && page < PageCount) ? _img[page] : null;
	}

	private Rectangle GetImageRectangle(Image img)
	{
		Size imageSizeInPixels = GetImageSizeInPixels(img);
		Rectangle result = new Rectangle(0, 0, imageSizeInPixels.Width, imageSizeInPixels.Height);
		Rectangle clientRectangle = ((Control)this).ClientRectangle;
		switch (_zoomMode)
		{
		case ZoomMode.ActualSize:
			_zoom = 1.0;
			break;
		case ZoomMode.TwoPages:
			result.Width *= 2;
			goto case ZoomMode.FullPage;
		case ZoomMode.FullPage:
		{
			double val = ((result.Width > 0) ? ((double)clientRectangle.Width / (double)result.Width) : 0.0);
			double val2 = ((result.Height > 0) ? ((double)clientRectangle.Height / (double)result.Height) : 0.0);
			_zoom = Math.Min(val, val2);
			break;
		}
		case ZoomMode.PageWidth:
			_zoom = ((result.Width > 0) ? ((double)clientRectangle.Width / (double)result.Width) : 0.0);
			break;
		}
		result.Width = (int)((double)result.Width * _zoom);
		result.Height = (int)((double)result.Height * _zoom);
		int num = (clientRectangle.Width - result.Width) / 2;
		if (num > 0)
		{
			result.X += num;
		}
		int num2 = (clientRectangle.Height - result.Height) / 2;
		if (num2 > 0)
		{
			result.Y += num2;
		}
		result.Inflate(-4, -4);
		if (_zoomMode == ZoomMode.TwoPages)
		{
			result.Inflate(-2, 0);
		}
		return result;
	}

	private Size GetImageSizeInPixels(Image img)
	{
		SizeF physicalDimension = img.PhysicalDimension;
		if (img is Metafile)
		{
			if (_himm2pix.X < 0f)
			{
				Graphics val = ((Control)this).CreateGraphics();
				try
				{
					_himm2pix.X = val.DpiX / 2540f;
					_himm2pix.Y = val.DpiY / 2540f;
				}
				finally
				{
					((IDisposable)val)?.Dispose();
				}
			}
			physicalDimension.Width *= _himm2pix.X;
			physicalDimension.Height *= _himm2pix.Y;
		}
		return Size.Truncate(physicalDimension);
	}

	private void RenderPage(Graphics g, Image img, Rectangle rc)
	{
		rc.Offset(1, 1);
		g.DrawRectangle(Pens.Black, rc);
		rc.Offset(-1, -1);
		g.FillRectangle(Brushes.White, rc);
		g.DrawImage(img, rc);
		g.DrawRectangle(Pens.Black, rc);
		rc.Width++;
		rc.Height++;
		g.ExcludeClip(rc);
		rc.Offset(1, 1);
		g.ExcludeClip(rc);
	}

	private void UpdateScrollBars()
	{
		Rectangle rectangle = Rectangle.Empty;
		Image image = GetImage(StartPage);
		if (image != null)
		{
			rectangle = GetImageRectangle(image);
		}
		Size size = new Size(0, 0);
		switch (_zoomMode)
		{
		case ZoomMode.PageWidth:
			size = new Size(0, rectangle.Height + 8);
			break;
		case ZoomMode.ActualSize:
		case ZoomMode.Custom:
			size = new Size(rectangle.Width + 8, rectangle.Height + 8);
			break;
		}
		if (size != ((ScrollableControl)this).AutoScrollMinSize)
		{
			((ScrollableControl)this).AutoScrollMinSize = size;
		}
		UpdatePreview();
	}

	private void UpdatePreview()
	{
		if (_startPage < 0)
		{
			_startPage = 0;
		}
		if (_startPage > PageCount - 1)
		{
			_startPage = PageCount - 1;
		}
		((Control)this).Invalidate();
	}
}
