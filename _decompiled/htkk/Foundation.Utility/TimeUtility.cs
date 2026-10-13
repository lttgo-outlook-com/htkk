using System;

namespace Foundation.Utility;

internal class TimeUtility
{
	public static TimeSpan MultiplySpan(TimeSpan span, int n)
	{
		if (n <= 0)
		{
			return TimeSpan.MinValue;
		}
		TimeSpan result = span;
		for (int i = 0; i < n - 1; i++)
		{
			result += span;
		}
		return result;
	}
}
