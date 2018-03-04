using System;
using System.Collections.Generic;
using System.Web;
using System.Threading;
using System.IO;

/// <summary>
/// BackgroundWork 的摘要描述
/// </summary>
public class BackgroundWork
{
	//public static object oLock = new object();
	private static Timer timer;

	// 開始背景作業
	public void StartWork() {
		TimeSpan delayTime = new TimeSpan(0, 0, 5); // 應用程式起動後5秒開始執行
		TimeSpan intervalTime = new TimeSpan(0, 0, 2); // 應用程式起動後間隔2秒重複執行
		TimerCallback timerDelegate = new TimerCallback(BatchMethod);  // 委派呼叫方法
		timer = new Timer(timerDelegate, null, delayTime, intervalTime);  // 產生 timer
	}

	// 背景批次方法
	private void BatchMethod(object pStatus) {
		//lock (oLock) {
			using (StreamWriter sw = new StreamWriter(System.Web.Hosting.HostingEnvironment.MapPath("~/TimeLog.txt"), true)) {
				sw.WriteLine(DateTime.Now);
			}
		//}
	}
}