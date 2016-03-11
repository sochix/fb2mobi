using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using CommandLine.Utility;

namespace fb2mobi
{
	class FB2mobiMain
	{
		private static string kindlegen = "kindlegen.exe";
		private static string opfxsl = "FB2_2_opf.xsl";
		private static string bodyxsl = "FB2_2_xhtml.xsl";
		private static string ncxxsl = "FB2_2_ncx.xsl";

		static void print_usage()
		{
			Console.WriteLine("Usage: fb2mobi <file.fb2> [<output.mobi>] [{-,/,--}param]");
			Console.WriteLine("  -nc \t No compress output file. Increase speed and size");
			Console.WriteLine("  -cl \t Clean output dir after convert.");
			Console.WriteLine("  -v0 \t Suppress verbose.");
			Console.WriteLine("  -v1 \t Suppress verbose. Only output file name.");
			Console.WriteLine("  -kindle-mail \t Kindle email");
			Console.WriteLine("  -from-mail \t from email");
			Console.WriteLine("  -smtp-server \t smtp server");
			Console.WriteLine("  -smtp-port \t smtp port");
			Console.WriteLine("  -use-ssl \t use ssl connection");
			Console.WriteLine("  -uname \t your email login");
			Console.WriteLine("  -password \t your email password");
			Worker.print_usage();
			Console.ReadLine();
		}

		static void print_copyright()
		{
			Console.WriteLine("FB2mobi v 2.0.4 Copyright (c) 2008-2012 Rakunov Alexander 2012-01-07");
			Console.WriteLine("Project home: http://code.google.com/p/fb2mobi/\n");
		}

		[STAThread]
		static int Main(string[] args)
		{

			if (args.Length == 0)
			{
				print_copyright();
				print_usage();
				return 1;
			}

			var commandLine = new Arguments(args);

			// parsing of mail settings
			var kindleMail = commandLine["kindle-mail"];
			var fromMail = commandLine["from-mail"];

			var smtpServer = commandLine["smtp-server"];
			int smtpPort;
			if (!int.TryParse(commandLine["smtp-port"], out smtpPort))
			{
				Console.WriteLine("Can't parse smptServer address.");
				return 0;
			}

			var useSsl = commandLine["use-ssl"] == "true";
			var userName = commandLine["uname"];
			var password = commandLine["password"];

			var verbose = (commandLine["v0"] == "" && commandLine["v1"] == "");

			if (verbose)
				print_copyright();

			if (commandLine["?"] == "true" || commandLine["help"] == "true" || commandLine["h"] == "true")
			{
				print_usage();
				return 0;
			}

			var filename = commandLine[0];
			if (!File.Exists(filename))
			{
				Console.Error.WriteLine("File: \"" + filename + "\" not found\n");
				if (verbose)
					print_usage();
				return 1;
			}

			// SET CURRENT DIR TO FB2MOBI EXECUTE DIR

			var pathToExecute = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
			if (pathToExecute.Trim().Length > 0)
				Directory.SetCurrentDirectory(pathToExecute);

			if (!File.Exists(kindlegen))
			{
				Console.Error.WriteLine("File: \"" + kindlegen + "\" not found\n");
				return 1;
			}

			// PREPARE DATA


			Worker sp = new Worker(commandLine);

			if (sp.error())
			{
				Console.Error.WriteLine("Init error.\n");
				if (verbose)
					print_usage();
				return 1;
			}


			// GET SOURCE FILES FOR KINGLEGEN FROM FB2


			try
			{
				sp.saveImages();
				sp.transform(bodyxsl, "index.html");
				sp.transform(opfxsl, sp.getBookName(".opf"));
				sp.transform(ncxxsl, "book.ncx");
			}
			catch (Exception e)
			{
				Console.Error.WriteLine("error occured: " + e.Message);
				return 1;
			}


			// RUN KINDLEGEN


			System.Diagnostics.Process process = new System.Diagnostics.Process();
			process.StartInfo.UseShellExecute = false;
			process.StartInfo.RedirectStandardOutput = true;
			process.StartInfo.RedirectStandardError = true;
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.FileName = kindlegen;

			string KindleGenArguments = commandLine["nc"] == "true" ? " -c0" : " -c2";
			KindleGenArguments += " \"" + sp.getWorkDir() + sp.getBookName(".opf") + "\"";

			process.StartInfo.Arguments = KindleGenArguments;

			process.Start();

			string str;
			while ((str = process.StandardOutput.ReadLine()) != null)
				if (verbose && str.Length > 0)
					Console.WriteLine(str);

			process.Close();

			// CLEAN AND PUBLISH
			if (verbose)
				Console.WriteLine("");

			var bookname = sp.getBookName(".mobi");
			if (File.Exists(sp.getWorkDir() + bookname))
			{
				if (!string.IsNullOrEmpty(kindleMail))
				{
					if (!SendEmail(
						bookname, 
						sp.getWorkDir() + bookname,
						kindleMail,
						fromMail,
						smtpServer,
						smtpPort,
						userName,
						password,
						useSsl
						))
						return -1;
				}
					
				File.Move(sp.getWorkDir() + bookname, sp.getOutputDir() + bookname);

				if (commandLine["cl"] == "true")
				{
					try
					{
						Directory.Delete(sp.getWorkDir(), true);
					}
					catch (Exception) { }
				}
				else if (verbose)
					Console.WriteLine("Output: " + sp.getWorkDir());

				if (verbose)
					Console.WriteLine("Book: " + sp.getOutputDir() + bookname);
				else if (commandLine["v1"] == "true")
					Console.WriteLine(sp.getOutputDir() + bookname);

				return 0;

			}
			else
			{
				if (verbose)
					Console.WriteLine("The output file is missing.");
				try
				{
					Directory.Delete(sp.getWorkDir(), true);
				}
				catch (Exception) { }
			}

			return 1;
		}

		private static bool SendEmail(string bookName, 
			string filePath, 
			string kindleMail, 
			string fromMail, 
			string smtpServer, 
			int smtpPort,
			string userName,
			string password,
			bool useSsl)
		{
			var message = new MailMessage(fromMail, kindleMail);
			using (var client = new SmtpClient
			{
				Port = smtpPort,
				DeliveryMethod = SmtpDeliveryMethod.Network,
				UseDefaultCredentials = false,
				Host = smtpServer,
				EnableSsl = useSsl,
				Credentials = new NetworkCredential(userName, password)
			})
			{
				using (var mobiFile = new Attachment(filePath))
				{
					message.Subject = bookName;
					message.Body = $"{bookName} from fb2mobi";
					message.Attachments.Add(mobiFile);

					try
					{
						client.Send(message);
						Console.WriteLine("Message send succesfully");
						return true;
					}
					catch (Exception e)
					{
						Console.WriteLine("Error occured during sending message: " + e.Message);
						return false;
					}
					
				}
			}
		}
	}
}
