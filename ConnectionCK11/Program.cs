﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using Newtonsoft.Json;
using System.Linq;
using System.Net;
using OfficeOpenXml;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace WriteRequestCK11
{
	class Token
	{
		public string access_token { get; set; }
		public string token_type { get; set; }
		public string expires_in { get; set; }
		public string user_login { get; set; }
		public string user_host { get; set; }
	}

	class Value
	{
		public string uid { get; set; }
		public string timeStamp { get; set; }
		public string timeStamp2 { get; set; }
		public long qCode { get; set; }
		public double value { get; set; }
	}

	class WriteRequest
	{
		public string writeType { get; set; }
		public List<Value> values { get; set; }
	}

	class Program
	{
		static string ck11PolEP = ConfigurationManager.AppSettings["ck11EndPoint"];
		static string auth = ck11PolEP + ConfigurationManager.AppSettings["ck11TokenEndPoint"];
		static string measurWrite = ck11PolEP + ConfigurationManager.AppSettings["ck11MeasurWriteEndPoint"];

		static Token getToken()
		{
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;
			Console.WriteLine("Запрос токена в СК11");
			WebRequest webRequestToken = WebRequest.Create(auth);
			webRequestToken.Method = "POST";
			webRequestToken.Credentials = CredentialCache.DefaultCredentials;
			WebResponse WebResponseToken = webRequestToken.GetResponse();
			string tokenBody = "";
			using (Stream tokenStream = WebResponseToken.GetResponseStream())
			{
				using (StreamReader tokenStreamReader = new StreamReader(tokenStream))
				{
					tokenBody = tokenStreamReader.ReadToEnd();
				}
			}
			return JsonConvert.DeserializeObject<Token>(tokenBody);
		}

		public static void writeDataToCK11(string uid, string timeStamp, string timeStamp2, long qCode, double value)
		{
			ServicePointManager.Expect100Continue = true;
			ServicePointManager.DefaultConnectionLimit = 9999;
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

			Token token = getToken();

			Console.WriteLine("Запись данных в СК11");
			WebRequest webRequestWrite = WebRequest.Create(measurWrite);
			webRequestWrite.Method = "POST";
			webRequestWrite.ContentType = "application/json";
			webRequestWrite.Headers.Add("charset", "UTF-8");
			webRequestWrite.Headers.Add("Authorization", token.token_type + " " + token.access_token);

			WriteRequest bodyRequest = new WriteRequest
			{
				/*writeType = "forcedToArchive",*/
				values = new List<Value>
				{
					new Value
					{
						uid = uid,
						timeStamp = timeStamp,
						timeStamp2 = timeStamp2,
						qCode = qCode,
						value = value
					}
				}
			};

			string requestBodyJson = JsonConvert.SerializeObject(bodyRequest);
			Console.WriteLine("Отправляемое тело запроса: " + requestBodyJson);

			using (Stream requestStream = webRequestWrite.GetRequestStream())
			{
				using (StreamWriter requestStreamWriter = new StreamWriter(requestStream))
				{
					requestStreamWriter.Write(requestBodyJson);
				}
			}

			try
			{
				WebResponse webResponseWrite = webRequestWrite.GetResponse();
				if (((HttpWebResponse)webResponseWrite).StatusDescription == "OK")
				{
					Console.WriteLine("Данные успешно записаны.");
				}
				else
				{
					Console.WriteLine("Ошибка при записи данных в СК-11: " + ((HttpWebResponse)webResponseWrite).StatusDescription);
				}
			}
			catch (WebException ex)
			{
				using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
				{
					string responseText = reader.ReadToEnd();
					Console.WriteLine("Ошибка при записи данных в СК-11: " + responseText);
				}
			}
		}

		static void Main()
		{
			// Пример записи данных
			writeDataToCK11(
				"A879B6EB-F0B6-4708-A422-12E8890B1D4A",
				"2024-12-11T22:05:00Z",
				"2024-12-11T22:05:00Z",
				// нужно вводить в Dec, в навигаторе данных отображает в Hex
				805310466,
				60);

			Console.WriteLine("\nЗапись данных окончена.");
			Console.ReadLine();
		}
	}
}
