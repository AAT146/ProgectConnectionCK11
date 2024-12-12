using System;
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
	// Для обращения к большинству методов API СК-11 требуется, чтобы запрос был аутентифицирован. 
	// Аутентификация в запросах к API СК-11, осуществляется путем передачи токена доступа в виде Bearer-токена в заголовке каждого запроса и 
	// предназначена для аутентификации запросов из внешних серверных приложений. 
	// Токен доступа — это непрозрачная строка, которая соответствует идентификатору сессии аутентификации.
	class Token
	{
		// Токен доступа
		public string access_token { get; set; }
		// Признак, обозначающий, что данный идентификатор может быть использовать со схемой аутентификации Bearer
		public string token_type { get; set; }
		// Количество секунд, через которое время жизни сессии истечет
		public string expires_in { get; set; }
		// Логин пользователя
		public string user_login { get; set; }
		// Имя компьютера, с которого произведен вход (получен активный токен доступа), либо IP адресс, если получить имя не удалось
		public string user_host { get; set; }
	}

	// Внутренняя схема тела запроса (для формирования записи данных)
	class Value
	{
		// UID значения измерений
		public string uid { get; set; }
		// Левая граница интервала
		public string timeStamp { get; set; }
		// Правая граница интервала
		public string timeStamp2 { get; set; }
		// Код качества
		public long qCode { get; set; }
		// Значение измерения
		public double value { get; set; }
	}

	// Схема тела запроса
	class WriteRequest
	{
		// Тип записи данных
		public string writeType { get; set; }
		// Лист значений измерений
		public List<Value> values { get; set; }
	}

	class Program
	{
		// Определение пользовательских параметров приложения: строки соединения с БД или заголовки окна браузера
		// Один элемент - add, два атрибута: key и value
		static string ck11PolEP = ConfigurationManager.AppSettings["ck11EndPoint"];
		static string auth = ck11PolEP + ConfigurationManager.AppSettings["ck11TokenEndPoint"];
		static string measurWrite = ck11PolEP + ConfigurationManager.AppSettings["ck11MeasurWriteEndPoint"];

		// Получить токен доступа 
		static Token getToken()
		{
			// Запрос токена
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

		// Записать данные в СК-11
		public static void writeDataToCK11(string uid, string timeStamp, string timeStamp2, long qCode, double value)
		{
			ServicePointManager.Expect100Continue = true;
			ServicePointManager.DefaultConnectionLimit = 9999;
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

			// Получение токена
			Token token = getToken();

			// Запись значений
			Console.WriteLine("Запись данных в СК11");
			WebRequest webRequestWrite = WebRequest.Create(measurWrite);
			webRequestWrite.Method = "POST";
			webRequestWrite.ContentType = "application/json";
			webRequestWrite.Headers.Add("charset", "UTF-8");
			webRequestWrite.Headers.Add("Authorization", token.token_type + " " + token.access_token);

			// Тело запроса на запись
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
