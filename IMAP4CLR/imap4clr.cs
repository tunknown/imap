// TUnknown 2020 license cc0/public domain
// https://en.wikipedia.org/wiki/Creative_Commons_license#Zero_/_public_domain

//https://tools.ietf.org/html/rfc3501
//https://tools.ietf.org/html/rfc2822
//https://tools.ietf.org/html/rfc2683
//https://tools.ietf.org/html/rfc5321#section-4.5.3
//http://regexstorm.net/reference		-- c# compatible

using System;
using Microsoft.SqlServer.Server;
using System.Net.Sockets;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;
using System.Threading;
using System.Data.SqlTypes;
using System.Net.Mail;

public class IMAP4CLR
{
	const UInt32 TimeOut = 500;
	const int CRLFOverhead = 1000;
	const int ResponseLength = 1000;

	private static bool GetCommandResult(ref TcpClient IMAP, ref StreamWriter sw, string CMDID, ref byte[] data)
	{
		data = null;
		data = new byte[IMAP.ReceiveBufferSize]; // commands that return no more than one buffer must come here

		bool Completed = false, OKFound = false;
		Match mFields;
		DateTime start = System.DateTime.Now;
		string Res = "";
		int readed, ptr = 0, remain = data.Length;

		do
		{
			readed = IMAP.GetStream().Read(data, ptr, remain);
			ptr = ptr + readed;
			remain = remain - readed;

			if (0 < readed) start = System.DateTime.Now;
			Res = Encoding.ASCII.GetString(data).TrimEnd('\0');

			if (!Completed)
			{
				mFields = Regex.Match(Res, CMDID + " ([OBN][KAO][D]{0,1}) ([\\s\\S]+)", RegexOptions.IgnoreCase);
				Completed = mFields.Success;
				if (Completed)
				{
					OKFound = (mFields.Groups[1].ToString() == "OK");
					if (!OKFound) throw new Exception(Res);
				}
			}
			if (IMAP.Available == 0 && !Completed) Thread.Sleep(20);
		} while (!Completed && (System.DateTime.Now - start).TotalMilliseconds < TimeOut);

		return OKFound;
	}

	private static bool ProtocolCommand(ref TcpClient IMAP, ref StreamWriter sw, string Command, ref byte[] data)
	{
		bool OKFound = false;
		if (string.IsNullOrEmpty(Command)) return OKFound;

		string CMDID = Guid.NewGuid().ToString();

		//try
		{
			// can send NOOP before any command to workaround system messages from server if exists

			sw.WriteLine(CMDID + " " + Command);
#if DEBUG
				//string FN = "C:\\TEMP\\" + System.DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + "_" + CMDID;
				//File.WriteAllText(FN + ".cm", CMDID + " " + Command);
#endif
			OKFound = GetCommandResult(ref IMAP, ref sw, CMDID, ref data);
#if DEBUG
				//File.WriteAllBytes(FN + ".rs", data);
#endif
		}
		/*catch (Exception e)
		{
			Res = e.Message;
		}*/
		return OKFound;
	}

	private static string ReadTillEnd(ref TcpClient IMAP)
	{
		string Res = "";
		if (0 < IMAP.Available)
		{
			DateTime start = System.DateTime.Now;
			byte[] Buffer = new byte[IMAP.Available];
			try
			{
				while (0 < IMAP.Available && (System.DateTime.Now - start).TotalMilliseconds < TimeOut)
					if (0 < IMAP.GetStream().Read(Buffer, 0, Buffer.Length))
					{
						start = System.DateTime.Now;
						Res = Res + Encoding.ASCII.GetString(Buffer, 0, Buffer.Length).TrimEnd('\0');
					}
					else
						Thread.Sleep(20);
			}
			finally
			{
				Buffer = null;
			}
		}
		return Res;
	}

	private static bool FetchBody(ref TcpClient IMAP, ref StreamWriter sw, int UID, ref MemoryStream MS)
	{
		bool Fetched = false;
		DateTime start = System.DateTime.Now;
		string Res, CMDID = Guid.NewGuid().ToString();
		int readed = 0, BodyLength = 0, remain, BeginMessage = 0, Avail, Pos;
		Match match;

		MS = null;

		byte[] Buf = new byte[IMAP.ReceiveBufferSize];
		try
		{
			sw.WriteLine(CMDID + " UID FETCH " + UID + " BODY.PEEK[]");
			do
			{
				Res = "";
				if (BodyLength == 0) // first step to detect mail length
				{
					readed = IMAP.GetStream().Read(Buf, 0, Buf.Length); // wait for FETCH only response, not a other nonreaded server response
					if (0 < readed) start = System.DateTime.Now;

					if (100 < readed) remain = 100; else remain = readed;

					//code page has no sense because of command response starts by ASCII
					match = Regex.Match(Encoding.ASCII.GetString(Buf, 0, remain).TrimEnd('\0'), "[*] \\d+ FETCH [(]BODY[[]] {(\\d+)}[\\s]+(\\S+)", RegexOptions.IgnoreCase); // get mail size
					if (match.Success)
						if (Int32.TryParse(match.Groups[1].ToString(), out BodyLength))
							if (0 < BodyLength)
							{
								BodyLength = BodyLength + CRLFOverhead; // reserve for difference between {Size} and real IMAP protocol fetched data

								BeginMessage = match.Groups[2].Index;

								if (readed - BeginMessage < BodyLength)
									remain = readed - BeginMessage;
								else
								{
									remain = BodyLength;
									Res = ReadTillEnd(ref IMAP);// read till end unnecessary data
								}
								MS = new MemoryStream(BodyLength); // instead multiple small expand do one large

								MS.Write(Buf, BeginMessage, remain);
							}
				}
				else
				{
					Avail = (IMAP.Available < Buf.Length ? IMAP.Available : Buf.Length);
					if (0 < Avail)
					{
						readed = IMAP.GetStream().Read(Buf, 0, Avail);

						if (0 < readed)
						{
							start = System.DateTime.Now;
							if (MS.Capacity < MS.Position + readed) MS.SetLength(MS.Position + readed);
							MS.Write(Buf, 0, readed);

							Pos = (int)MS.Position;
							try
							{
								MS.Position = MS.Position - ResponseLength;
								readed = MS.Read(Buf, 0, ResponseLength); // read back to obtain non breaked by readed beffer size end of buffer
							}
							finally
							{
								MS.Position = Pos;
							}
						}
					}
					else
						readed = 0;
				}

				if (0 < readed)
				{
					if (readed < ResponseLength)
					{
						Pos = 0;
						Avail = readed;
					}
					else
					{
						Pos = readed - ResponseLength;
						Avail = ResponseLength;
					}
					Res = Encoding.ASCII.GetString(Buf, Pos, Avail);

					if (Res != "")
					{
						match = Regex.Match(Res, CMDID + " ([OBN][KAO][D]{0,1}) [\\s\\S]+", RegexOptions.IgnoreCase);
						if (match.Success)
							if (match.Groups[1].ToString() != "OK")
								throw new Exception(match.ToString());
							else
							{
								match = Regex.Match(Res, " UID \\d+[)][\\n\\r]{1,2}" + CMDID, RegexOptions.IgnoreCase); // not enought get OK because of fetching for not existed uid returns OK too
								Fetched = match.Success;
								if (Fetched) MS.SetLength(MS.Position - ResponseLength + match.Index); // shorten mail buffer- trim IMAP ptotocol response string
								ReadTillEnd(ref IMAP);
							}
					}
				}

				if (IMAP.Available == 0 && !Fetched) Thread.Sleep(20);
			} while (!Fetched && (System.DateTime.Now - start).TotalMilliseconds < TimeOut); // Timeout after successful reading
		}
		finally
		{
			Buf = null;
		}

		return Fetched;
	}

	private static bool Autenticate(ref TcpClient IMAP, ref StreamWriter sw/*thread-safe*/ , string Server, int Port, string User, string Password, bool IsReadOnly)
	{
		bool Result = false;

		IMAP = new TcpClient(Server, Port); // IMAP.Connect not needed
		sw = new StreamWriter(IMAP.GetStream());
		sw.AutoFlush = true;

		try
		{
			byte[] data = null;
			try
			{
				if (!GetCommandResult(ref IMAP, ref sw, "[*]", ref data)) throw new Exception(Encoding.ASCII.GetString(data, 0, data.Length).TrimEnd('\0'));
			}
			finally
			{
				data = null;
			}

			try
			{
				if (!ProtocolCommand(ref IMAP, ref sw, "LOGIN " + User + " " + Password, ref data)) throw new Exception(Encoding.ASCII.GetString(data, 0, data.Length).TrimEnd('\0'));
			}
			finally
			{
				data = null;
			}

			try
			{
				Result = ProtocolCommand(ref IMAP, ref sw, (IsReadOnly ? "EXAMINE" : "SELECT") + " INBOX", ref data);
				if (!Result) throw new Exception(Encoding.ASCII.GetString(data, 0, data.Length).TrimEnd('\0'));
			}
			finally
			{
				data = null;
			}
		}
		catch
		{
			Disconnect(ref IMAP, ref sw);
		}

		return Result;
	}

	private static void Disconnect(ref TcpClient IMAP, ref StreamWriter sw)
	{
		byte[] data = null;
		try
		{
			ProtocolCommand(ref IMAP, ref sw, "LOGOUT", ref data);
		}
		finally
		{
			data = null;
			if (IMAP.Connected) IMAP.Close();

			sw = null;
			IMAP = null;
		}
	}

	private static string Decode(string str)
	{
		if (8 < str.Length)
		{
			string part, chr;
			foreach (Match enc in Regex.Matches(str, "=[?][a-z0-9_$.:-]+[?](B|Q)[?][\\S]+[?]=[\\s]{0,1}", RegexOptions.IgnoreCase)) // according to https://tools.ietf.org/html/rfc1342
			{
				part = enc.ToString();
				chr = enc.Groups[1].ToString();
				str = str.Replace(part, Attachment.CreateAttachmentFromString("", ((chr == "b" || chr == "q") ? part.Replace("?b?", "?B?").Replace("?q?", "?Q?") : part).TrimEnd(null)).Name); // CreateAttachmentFromString wants strict letter case
			}
		}
		return str;
	}

	private static bool GetMailFields(ref byte[] data, ref string MessageId, ref Nullable<DateTime> Date, ref string From, ref string To, ref string Cc, ref string Subject, ref string Importance)
	{
		string Res = Encoding.ASCII.GetString(data, 0, data.Length).TrimEnd('\0');

		Res = Regex.Replace(Res, "\\r\\n(\\s)", "$1"); // unfolding due to https://tools.ietf.org/html/rfc2822#section-2.2.3

		bool MailPresent = Regex.Match(Res, "[\\r\\n].+:\\s.+[\\r\\n]", RegexOptions.IgnoreCase).Success; // detect if OK present but UID not exist
		if (MailPresent)
		{
			Match mFields = Regex.Match(Res, "[\\r\\n]Date:\\s([0-9a-z ,:+-]+)(?: [(][A-Z]{2,6}[)]){0,1}[\\r\\n]", RegexOptions.IgnoreCase /*| RegexOptions.Multiline*/);
			if (mFields.Success)
			{
				DateTime dt;
				if (DateTime.TryParse(mFields.Groups[1].ToString(), out dt)) // ignore https://en.wikipedia.org/wiki/List_of_time_zone_abbreviations
					try
					{
						Date = (Nullable<DateTime>)dt;
					}
					catch
					{
						Date = null/*DateTime.Now*/; // System.Data.SqlTypes.SqlTypeException DateTime out of bound
					}
				else
					Date = null;
			}
			else
				Date = null;

			mFields = Regex.Match(Res, "[\\r\\n]Message-ID:\\s[<]{0,1}([^\\r\\n]+)[>]{0,1}[\\r\\n]", RegexOptions.IgnoreCase);
			if (mFields.Success) MessageId = mFields.Groups[1].ToString(); else MessageId = null;

			mFields = Regex.Match(Res, "[\\r\\n]From:\\s([^\\r\\n]+)[\\r\\n]", RegexOptions.IgnoreCase);
			if (mFields.Success) From = Decode(mFields.Groups[1].ToString()); else From = null;

			mFields = Regex.Match(Res, "[\\r\\n]To:\\s([^\\r\\n]+)[\\r\\n]", RegexOptions.IgnoreCase);
			if (mFields.Success) To = Decode(mFields.Groups[1].ToString()); else To = null;

			mFields = Regex.Match(Res, "[\\r\\n]Cc:\\s([^\\r\\n]+)[\\r\\n]", RegexOptions.IgnoreCase);
			if (mFields.Success) Cc = Decode(mFields.Groups[1].ToString()); else Cc = null;

			mFields = Regex.Match(Res, "[\\r\\n]Subject:\\s([^\\r\\n]+)[\\r\\n]", RegexOptions.IgnoreCase);
			if (mFields.Success) Subject = Decode(mFields.Groups[1].ToString()); else Subject = null;

			mFields = Regex.Match(Res, "[\\r\\n]Importance:\\s([^\\r\\n]{3,6})[\\r\\n]", RegexOptions.IgnoreCase);
			if (mFields.Success) Importance = mFields.Groups[1].ToString(); else Importance = null;
		}
		return MailPresent;
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////
	[SqlFunction(FillRowMethodName = "FillList", TableDefinition = "UID int , MessageId nvarchar ( 1000 ) , Date datetime , From nvarchar ( 256 ) , To nvarchar ( max ) , Cc nvarchar ( max ) , Subject nvarchar ( 4000 ) , Importance nvarchar ( 6 )")]
	public static IEnumerable GetEMailsAvailable(string sServer, int iPort, string sUser, string sPassword, DateTime dtMoment)
	{
		TcpClient IMAP = null;
		StreamWriter sw = null;

		UInt16 Records = 0;
		ArrayList rows = new ArrayList();

		if (Autenticate(ref IMAP, ref sw, sServer, iPort, sUser, sPassword, true))
			try
			{
				string cmd = "UID SEARCH 1:* UNDELETED", MessageId = "", From = "", To = "", Cc = "", Subject = "", Importance = "";
				byte[] data = null;


				if (dtMoment != null) cmd = cmd + " SINCE " + dtMoment.ToString("dd-MMM-yyyy", new System.Globalization.CultureInfo("en-US"));
				if (ProtocolCommand(ref IMAP, ref sw, cmd, ref data))
				{
					Int32 UID;
					Nullable<DateTime> Date = null;
					MatchCollection mUIDs;

					try
					{
						mUIDs = Regex.Matches(Encoding.ASCII.GetString(data, 0, data.Length).TrimEnd('\0'), "([ ]+\\d+)", RegexOptions.IgnoreCase);
					}
					finally
					{
						data = null;
					}

					foreach (Match match in mUIDs)
						foreach (Capture capture in match.Captures)
							if (Int32.TryParse(capture.Value.ToString(), out UID))
								try
								{
									data = null;
									if (ProtocolCommand(ref IMAP, ref sw, "UID FETCH " + UID + " body.peek[header.fields (message-id date from to cc subject importance)]", ref data))
										if (GetMailFields(ref data, ref MessageId, ref Date, ref From, ref To, ref Cc, ref Subject, ref Importance))
										{
											rows.Add(new object[] { UID, MessageId, Date, From, To, Cc, Subject, Importance });
											Records++;
										}
								}
								finally
								{
									data = null;
								}
				}
			}
			/*catch
			{
				if (Records == 0) throw; // hide error to show obtained records
			}*/
			finally
			{
				Disconnect(ref IMAP, ref sw);
			}
		return rows;
	}

	private static void FillList(Object row, out int UID, out string MessageId, out Nullable<DateTime> Date, out string From, out string To, out string Cc, out string Subject, out string Importance)
	{
		UID = (int)((object[])row)[0];
		MessageId = (string)((object[])row)[1];
		Date = (Nullable<DateTime>)((object[])row)[2];
		From = (string)((object[])row)[3];
		To = (string)((object[])row)[4];
		Cc = (string)((object[])row)[5];
		Subject = (string)((object[])row)[6];
		Importance = (string)((object[])row)[7];
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////
	[SqlFunction(FillRowMethodName = "FillMessages", TableDefinition = "UID int , MessageId nvarchar ( 1000 ) , Date datetime , From nvarchar ( 256 ) , To nvarchar ( max ) , Cc nvarchar ( max ) , Subject nvarchar ( 4000 ) , Importance nvarchar ( 6 ) , Body varbinary ( max )")]
	public static IEnumerable GetEMails(string sServer, int iPort, string sUser, string sPassword, [SqlFacet(MaxSize = 4000)]string sUIDs)
	{
		TcpClient IMAP = null;
		StreamWriter sw = null;

		ArrayList rows = new ArrayList();

		if (1 < sUIDs.Length)
			if (Autenticate(ref IMAP, ref sw, sServer, iPort, sUser, sPassword, true))
			{
				UInt16 Records = 0;

				try
				{
					Int32 UID;
					string MessageId = "", From = "", To = "", Cc = "", Subject = "", Importance = "";
					Nullable<DateTime> Date = null;
					byte[] body;
					MemoryStream MS;

					foreach (string sUID in sUIDs.Split(new char[] { sUIDs[0] }, System.StringSplitOptions.RemoveEmptyEntries))
						if (Int32.TryParse(sUID, out UID))
						{
							body = null;
							if (ProtocolCommand(ref IMAP, ref sw, "UID FETCH " + UID + " body.peek[header.fields (message-id date from to cc subject importance)]", ref body))
								if (GetMailFields(ref body, ref MessageId, ref Date, ref From, ref To, ref Cc, ref Subject, ref Importance))
									try
									{
										MS = null;
										if (FetchBody(ref IMAP, ref sw, UID, ref MS))
										{
											rows.Add(new object[] { UID, MessageId, Date, From, To, Cc, Subject, Importance, MS });
											Records++;
										}
									}
									finally
									{
										MS = null;
										body = null;
									}
						}
				}
				/*catch
				{
					if (Records == 0) throw; //hide error to show obtained records
				}*/
				finally
				{
					Disconnect(ref IMAP, ref sw);
				}
			}
		return rows;
	}

	private static void FillMessages(Object row, out int UID, out string MessageId, out Nullable<DateTime> Date, out string From, out string To, out string Cc, out string Subject, out string Importance, [SqlFacet(MaxSize = -1)]out SqlBytes Body)
	{
		UID = (int)((object[])row)[0];
		MessageId = (string)((object[])row)[1];
		Date = (Nullable<DateTime>)((object[])row)[2];
		From = (string)((object[])row)[3];
		To = (string)((object[])row)[4];
		Cc = (string)((object[])row)[5];
		Subject = (string)((object[])row)[6];
		Importance = (string)((object[])row)[7];
		Body = new SqlBytes((MemoryStream)((object[])row)[8]);
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////
	[Microsoft.SqlServer.Server.SqlProcedure]
	public static void RemoveEMails(string sServer, int iPort, string sUser, string sPassword, [SqlFacet(MaxSize = 4000)]string sUIDs)
	{
		TcpClient IMAP = null;
		StreamWriter sw = null;

		if (1 < sUIDs.Length)
			if (Autenticate(ref IMAP, ref sw, sServer, iPort, sUser, sPassword, false))
				try
				{
					UInt32 UID;
					bool Closing = false;
					byte[] body;

					foreach (string sUID in sUIDs.Split(new char[] { sUIDs[0] }, System.StringSplitOptions.RemoveEmptyEntries))
						if (UInt32.TryParse(sUID, out UID))
							try
							{
								body = null;
								if (ProtocolCommand(ref IMAP, ref sw, "UID STORE " + UID + " +FLAGS.SILENT (\\Deleted)", ref body)) Closing = true;
							}
							finally
							{
								body = null;
							}

					if (Closing)
						try
						{
							body = null;
							ProtocolCommand(ref IMAP, ref sw, "CLOSE", ref body);
						}
						finally
						{
							body = null;
						}
				}
				finally
				{
					Disconnect(ref IMAP, ref sw);
				}
		return;
	}
}