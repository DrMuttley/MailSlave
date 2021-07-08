using System;

using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit;
using MimeKit;
using MailKit.Search;

using System.IO;
using System.Diagnostics;

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


		static void Main(string[] args)
		{
			ImapClient client = new ImapClient();

			client.Connect(IMAP_SERVER_ADDRESS, IMAP_PORT_NUMBER, true);
			client.Authenticate(LOGIN, PASSWORD);

			client.Inbox.Open(FolderAccess.ReadWrite);

			foreach (var uid in client.Inbox.Search(SearchQuery.NotSeen))
			{
				var message = client.Inbox.GetMessage(uid);
				Console.WriteLine(message.Subject); // to delete

				if (message.Subject.Contains("SAVE"))
				{
					client.Inbox.AddFlags(uid, MessageFlags.Seen, true);

					string[] messageSubjectData = message.Subject.Split('#');

					foreach (var attachment in message.Attachments)
					{
						Stream stream = File.Create(messageSubjectData[1]);

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
					sendMail($"The attachment {messageSubjectData[1]} was saved correctly");
				}

				if (message.Subject.Contains("RUN"))
				{
					string[] messageSubjectData = message.Subject.Split('#');

					client.Inbox.AddFlags(uid, MessageFlags.Seen, true);
					Process.Start(messageSubjectData[1]);

					sendMail($"The app {messageSubjectData[1]} has started");
				}
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
