using System;
using System.Threading;

namespace HTKK;

public class Forker
{
	private int running;

	private readonly object joinLock = new object();

	private readonly object eventLock = new object();

	private EventHandler allComplete;

	private EventHandler<ParallelEventArgs> itemComplete;

	public event EventHandler AllComplete
	{
		add
		{
			lock (eventLock)
			{
				allComplete = (EventHandler)Delegate.Combine(allComplete, value);
			}
		}
		remove
		{
			lock (eventLock)
			{
				allComplete = (EventHandler)Delegate.Remove(allComplete, value);
			}
		}
	}

	public event EventHandler<ParallelEventArgs> ItemComplete
	{
		add
		{
			lock (eventLock)
			{
				itemComplete = (EventHandler<ParallelEventArgs>)Delegate.Combine(itemComplete, value);
			}
		}
		remove
		{
			lock (eventLock)
			{
				itemComplete = (EventHandler<ParallelEventArgs>)Delegate.Remove(itemComplete, value);
			}
		}
	}

	private void OnItemComplete(object state, Exception exception)
	{
		itemComplete?.Invoke(this, new ParallelEventArgs(state, exception));
		if (Interlocked.Decrement(ref running) == 0)
		{
			allComplete?.Invoke(this, EventArgs.Empty);
			lock (joinLock)
			{
				Monitor.PulseAll(joinLock);
			}
		}
	}

	public Forker OnItemComplete(EventHandler<ParallelEventArgs> handler)
	{
		if (handler == null)
		{
			throw new ArgumentNullException("handler");
		}
		ItemComplete += handler;
		return this;
	}

	public Forker OnAllComplete(EventHandler handler)
	{
		if (handler == null)
		{
			throw new ArgumentNullException("handler");
		}
		AllComplete += handler;
		return this;
	}

	public void Join()
	{
		Join(-1);
	}

	public bool Join(int millisecondsTimeout)
	{
		lock (joinLock)
		{
			if (CountRunning() == 0)
			{
				return true;
			}
			Thread.SpinWait(1);
			return CountRunning() == 0 || Monitor.Wait(joinLock, millisecondsTimeout);
		}
	}

	public int CountRunning()
	{
		return Interlocked.CompareExchange(ref running, 0, 0);
	}

	public Forker Fork(ThreadStart action)
	{
		return Fork(action, null);
	}

	public Forker Fork(ThreadStart action, object state)
	{
		if (action == null)
		{
			throw new ArgumentNullException("action");
		}
		Interlocked.Increment(ref running);
		ThreadPool.QueueUserWorkItem(delegate
		{
			Exception exception = null;
			try
			{
				action();
			}
			catch (Exception ex)
			{
				exception = ex;
			}
			OnItemComplete(state, exception);
		});
		return this;
	}
}
