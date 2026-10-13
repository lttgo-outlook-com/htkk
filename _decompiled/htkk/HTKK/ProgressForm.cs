using System;
using System.ComponentModel;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using HTKK.Properties;

namespace HTKK;

public class ProgressForm : Form
{
	public delegate void DoWorkEventHandler(ProgressForm sender, DoWorkEventArgs e);

	private BackgroundWorker worker;

	public string _strTitle = "";

	public bool _finish = false;

	private IContainer components = null;

	private Label labelStatus;

	private PictureBox pictureBox1;

	public object Argument { get; set; }

	public RunWorkerCompletedEventArgs Result { get; private set; }

	public event DoWorkEventHandler DoWork;

	public ProgressForm(string strTitle)
	{
		InitializeComponent();
		_finish = false;
		_strTitle = strTitle;
		((Control)labelStatus).Text = strTitle;
		worker = new BackgroundWorker();
		worker.WorkerReportsProgress = true;
		worker.WorkerSupportsCancellation = true;
		worker.ProgressChanged += worker_ProgressChanged;
		worker.DoWork += worker_DoWork;
		worker.RunWorkerCompleted += worker_RunWorkerCompleted;
	}

	private void ProgressForm_Load(object sender, EventArgs e)
	{
		Result = null;
		worker.RunWorkerAsync(Argument);
	}

	private void worker_DoWork(object sender, DoWorkEventArgs e)
	{
		new Thread((ThreadStart)delegate
		{
			Thread.CurrentThread.IsBackground = true;
			for (int i = 0; i < 20000; i++)
			{
				Thread.Sleep(1000);
				if (!_finish)
				{
					try
					{
						worker.ReportProgress(i);
					}
					catch (Exception)
					{
						break;
					}
				}
				else if (i > 0)
				{
					break;
				}
			}
		}).Start();
		if (this.DoWork != null)
		{
			this.DoWork(this, e);
		}
	}

	private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		string text = ((Control)labelStatus).Text;
		string[] array = text.Split(new char[1] { '.' });
		if (array.Length == 1)
		{
			((Control)labelStatus).Text = ((Control)labelStatus).Text + ".";
		}
		if (array.Length == 2)
		{
			((Control)labelStatus).Text = ((Control)labelStatus).Text + ".";
		}
		if (array.Length == 3)
		{
			((Control)labelStatus).Text = ((Control)labelStatus).Text + ".";
		}
		if (array.Length == 4)
		{
			((Control)labelStatus).Text = ((Control)labelStatus).Text.Replace(".", "");
		}
	}

	private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		_finish = true;
		Result = e;
		if (e.Error != null)
		{
			((Form)this).DialogResult = (DialogResult)3;
		}
		else if (e.Cancelled)
		{
			((Form)this).DialogResult = (DialogResult)2;
		}
		else
		{
			((Form)this).DialogResult = (DialogResult)1;
		}
		((Form)this).Close();
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
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Expected O, but got Unknown
		labelStatus = new Label();
		pictureBox1 = new PictureBox();
		((ISupportInitialize)pictureBox1).BeginInit();
		((Control)this).SuspendLayout();
		((Control)labelStatus).AutoSize = true;
		((Control)labelStatus).ForeColor = Color.White;
		((Control)labelStatus).Location = new Point(12, 9);
		((Control)labelStatus).Name = "labelStatus";
		((Control)labelStatus).Size = new Size(66, 13);
		((Control)labelStatus).TabIndex = 0;
		((Control)labelStatus).Text = "Đang xử lý...";
		pictureBox1.Image = (Image)(object)Resources.ProgressBarImage;
		((Control)pictureBox1).Location = new Point(15, 25);
		((Control)pictureBox1).Name = "pictureBox1";
		((Control)pictureBox1).Size = new Size(242, 18);
		pictureBox1.TabIndex = 1;
		pictureBox1.TabStop = false;
		((ContainerControl)this).AutoScaleDimensions = new SizeF(6f, 13f);
		((ContainerControl)this).AutoScaleMode = (AutoScaleMode)1;
		((Control)this).BackColor = Color.FromArgb(255, 128, 0);
		((Form)this).ClientSize = new Size(271, 56);
		((Form)this).ControlBox = false;
		((Control)this).Controls.Add((Control)(object)pictureBox1);
		((Control)this).Controls.Add((Control)(object)labelStatus);
		((Form)this).FormBorderStyle = (FormBorderStyle)0;
		((Control)this).Name = "ProgressForm";
		((Form)this).StartPosition = (FormStartPosition)1;
		((Control)this).Text = "Operation in progress";
		((Form)this).Load += ProgressForm_Load;
		((ISupportInitialize)pictureBox1).EndInit();
		((Control)this).ResumeLayout(false);
		((Control)this).PerformLayout();
	}
}
