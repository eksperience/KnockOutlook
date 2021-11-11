using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace KnockOutlook.Operations
{
	class Search
	{
		private static readonly Dictionary<string, HashSet<string>> Findings = new Dictionary<string, HashSet<string>>();
		private static readonly List<string> FalsePositive = new List<string>();
		private static string ExceptionMessage;

		public static void Run(string keyword, bool bypass)
		{
			Find(keyword, bypass);

			foreach (string key in Findings.Keys)
			{
				if (Findings[key].Count == 0)
				{
					FalsePositive.Add(key);
				}
			}

			foreach(string falsePositive in FalsePositive)
			{
				Findings.Remove(falsePositive);
			}
			
			ShowResults(keyword);
		}

		private static void ShowResults(string keyword)
		{
			string output = Program.Banner();

			if (string.IsNullOrEmpty(ExceptionMessage))
			{
				if (Findings.Count > 0 || FalsePositive.Count > 0)
				{
					int index = 0;

					if (Findings.Count > 0)
					{
						output += string.Format("Hits for keyword '{0}': \r\n", keyword);

						foreach (string key in Findings.Keys)
						{
							index++;
							output += string.Format("\r\n{0}. {1}\r\n", index, key);

							using (HashSet<string>.Enumerator enumerator = Findings[key].GetEnumerator())
							{
								string current;
								string sign;
								bool hasNext = enumerator.MoveNext();

								while (hasNext)
								{
									current = enumerator.Current;
									sign = (hasNext = enumerator.MoveNext()) ? "\u251C\u2574" : "\u2514\u2574";
									output += string.Format(" {0}{1}\r\n", sign, current);
								}
							}
						}
					}
					else
					{
						if (FalsePositive.Count > 0)
						{
							output += string.Format("No hits for keyword '{0}'\r\n", keyword);
						}
					}
				}
				else
				{
					output += "No configured accounts were identified\r\n";
				}
			}
			else
			{
				output += ExceptionMessage + "\r\n";
			}

			Console.Write(output);
		}

		private static void Find(string keyword, bool bypass)
		{
			try
			{
				Application Outlook = new Application();
				string majorVersion = Utilities.GetMajorVersion(Outlook.ProductCode);
				List<int> snapshot = null;

				if (bypass)
				{
					snapshot = Utilities.Bypass(majorVersion);
				}

				foreach (Account sessionAccount in Outlook.Session.Accounts)
				{
					HashSet<string> hits = new HashSet<string>();
					List<Folder> mailFolders = Utilities.GetFolders((Folder)sessionAccount.DeliveryStore.GetRootFolder());
					string filter = string.Format("@SQL=\"urn:schemas:httpmail:textdescription\" like '%{0}%'", keyword);

					if (mailFolders.Count > 0)
					{
						foreach (Folder folder in mailFolders)
						{
							foreach (object item in folder.Items.Restrict(filter))
							{
								try
								{
									MailItem mailItem = (MailItem)item;
									hits.Add(mailItem.EntryID);
								}
								catch (InvalidCastException)
								{
									continue;
								}
							}
						}
					}

					Findings.Add(Utilities.GetAccountEmailAddress(sessionAccount), hits);
				}

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
				else
				{
					ExceptionMessage = exception.ToString();
				}
			}
		}
	}
}
