using System.Threading;
using Foundation.Base;

namespace Foundation.Utility;

internal class ThreadProgram : BaseObject
{
	protected int stop;

	protected int work_interval;

	public ThreadProgram(int work_interval)
	{
		if (work_interval <= 0)
		{
			work_interval = 1;
		}
		this.work_interval = work_interval;
		stop = 0;
	}

	public void DoWork()
	{
		if (InitProgram() == 0)
		{
			stop = 0;
			while (stop != 1)
			{
				MainProgram();
				Thread.Sleep(work_interval);
			}
			FinishingProgram();
		}
	}

	protected virtual void MainProgram()
	{
	}

	protected virtual int InitProgram()
	{
		return 0;
	}

	protected virtual void FinishingProgram()
	{
	}

	public void RequestStop()
	{
		stop = 1;
	}
}
