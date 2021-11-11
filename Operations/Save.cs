using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace KnockOutlook.Operations
{
	class Save
	{
		private static string ExportPath;
		private static string ExceptionMessage;

		public static void Run(string id, bool bypass)
		{
			SaveMail(id, bypass);
			ShowResults();
		}

		private static void ShowResults()
		{
			string output = Program.Banner();

			if (string.IsNullOrEmpty(ExceptionMessage))
			{
				output += string.Format("Saved at {0}\r\n", ExportPath);
			}
			else
			{
				output += ExceptionMessage + "\r\n";
			}

			Console.Write(output);
		}

		private static void SaveMail(string id, bool bypass)
		{
			try
			{
				Application Outlook = new Application();
				string majorVersion = Utilities.GetMajorVersion(Outlook.ProductCode);
				ExportPath = Utilities.GenerateExportPath(majorVersion);
				List<int> snapshot = null;

				if (bypass)
				{
					snapshot = Utilities.Bypass(majorVersion);
				}

				NameSpace ns = Outlook.GetNamespace("MAPI");
				MailItem mailItem = (MailItem)ns.GetItemFromID(id);
				mailItem.SaveAs(ExportPath, OlSaveAsType.olMSGUnicode);

				if (bypass)
				{
					Utilities.Revert(majorVersion, snapshot);
				}
			}
			catch (COMException exception)
			{
				if (exception.HResult == -2147221164)
				{
					ExceptionMessage = "Outlook is not installed on the system";
				}
				else if (exception.HResult == -2079195127)
				{
					ExceptionMessage = "Outlook is not connected";
				}
				else if (exception.HResult == -2147467260)
				{
					ExceptionMessage = "Operation aborted by the user";
				}
				else if (exception.HResult == -2147352567)
				{
					ExceptionMessage = "Invalid EntryID";
				}
				else
				{
					ExceptionMessage = exception.ToString();
				}
			}
		}
	}
}
