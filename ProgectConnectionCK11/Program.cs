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

namespace ReadRequestCK11
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

    // Схема тела запроса
    class ReadRequest
    {
        // UID-ы значений измерений
        public string[] uids { get; set; }
        // Левая граница интервала
        public string fromTimeStamp { get; set; }
        // Правая граница интервала
        public string toTimeStamp { get; set; }
        // Единицы измерения шага времени между столбцами: секунды, дни, месяцы, года
        public string stepUnits { get; set; }
        // Значение шага времени между столбцами
        public uint stepValue { get; set; }
    }

    // Схема ответа
    // Содержит таблицу данных значений измерений.
    // Если измерение не найдено или у него отсутствуют данные за некоторый момент времени, соответствующий экземпляр данных будет содержать нулевой код качества.
    class ReadResponse
    {
        // Массив объектов
        public ReadResponseValue[] value { get; set; }
    }

    class ReadResponseValue
    {
        // UID значения измерения, данные которого находятся в строке таблицы
        public string uid { get; set; }
        // Массив объектов
        public Value[] value { get; set; }
    }

    class Value
    {
        // Глобально-уникальный идентификатор
        public string uid { get; set; }
        // Перва метка времени
        public string timeStamp { get; set; }
        // Вторая метка времени
        public string timeStamp2 { get; set; }
        // Коды качества (32-битное знаковое число)
        public long qCode { get; set; }
        // Фактическое значение экземпляра данных измерения. 
        public double value { get; set; }
    }

    class Program
    {
        /// <summary>
        /// Определение пользовательских параметров приложения: строки соединения с БД или заголовки окна браузера
        /// Один элемент - add, два атрибута: key и value
        /// </summary>
        static string ck11PolEP = ConfigurationManager.AppSettings["ck11EndPoint"];
        static string auth = ck11PolEP + ConfigurationManager.AppSettings["ck11TokenEndPoint"];
        static string measurRead = ck11PolEP + ConfigurationManager.AppSettings["ck11MeasurReadEndPoint"];

        static List<string> ck11Uids = new List<string>
            {
                "A879B6EB-F0B6-4708-A422-12E8890B1D4A"
            };

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
            Token token = JsonConvert.DeserializeObject<Token>(tokenBody);
            return token;
        }

        // Получить данные из СК11
        public static ReadResponse getDataFromCK11(string timeStart, string timeEnd, List<string> uids)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.DefaultConnectionLimit = 9999;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3;

            // Получение токена
            Token token = getToken();

            // Получение измерений
            Console.WriteLine("Чтение измерений из СК11");
            WebRequest webRequestMeasur = WebRequest.Create(measurRead);
            webRequestMeasur.Method = "POST";
            webRequestMeasur.ContentType = "application/json";
            webRequestMeasur.Headers.Add("charset", "UTF-8");
            webRequestMeasur.Headers.Add("Authorization", token.token_type + " " + token.access_token);

            // Формирование тела запроса измерений
            ReadRequest bodyRequest = new ReadRequest();
            bodyRequest.uids = uids.ToArray();
            bodyRequest.fromTimeStamp = timeStart;
            bodyRequest.toTimeStamp = timeEnd;
            bodyRequest.stepUnits = "seconds";
            bodyRequest.stepValue = 300;
            string requestBodyJson = JsonConvert.SerializeObject(bodyRequest);
            using (Stream requestStream = webRequestMeasur.GetRequestStream())
            {
                using (StreamWriter requestStreamWriter = new StreamWriter(requestStream))
                {
                    requestStreamWriter.Write(requestBodyJson);
                }
            }
            WebResponse webResponseMeasure = webRequestMeasur.GetResponse();
            if (((HttpWebResponse)webResponseMeasure).StatusDescription == "OK")
            {
                Console.WriteLine("Запрос успешно выполнен, но некоторые значения измерений могут не содержать данных в указанном интервале");
            }
            else
            {
                Console.WriteLine("Запрос на чтение данных из СК-11 не обработан");
            }
            string responseReadMeasureBody = "";
            using (Stream responseReadMeasureStream = webResponseMeasure.GetResponseStream())
            {
                using (StreamReader responseReadMeasureStreamReader = new StreamReader(responseReadMeasureStream))
                {
                    responseReadMeasureBody = responseReadMeasureStreamReader.ReadToEnd();
                }
            }
            ReadResponse readResponse = JsonConvert.DeserializeObject<ReadResponse>(responseReadMeasureBody);
            return readResponse;
        }

        static void Main()
        {
            ReadResponse myResponse = new ReadResponse();
            myResponse = getDataFromCK11("2023-11-13T09:00:00Z", "2023-11-13T21:00:00Z", ck11Uids);

            // Установка контекста лицензирования
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                // Добавление нового листа
                var worksheet = package.Workbook.Worksheets.Add("Sheet");

                int i = 1;
                foreach (ReadResponseValue responseValue in myResponse.value)
                {
                    foreach (Value val in responseValue.value)
                    {
                        // Запись в ячейки
                        worksheet.Cells[i, 1].Value = val.uid;
                        worksheet.Cells[i, 2].Value = val.value;
                        worksheet.Cells[i, 3].Value = val.timeStamp;
                        i++;
                        Console.WriteLine($"{val.uid} - {val.value} - {val.timeStamp}");
                    }
                }
                // Сохранение файла
                var file = new FileInfo(@"C:\Users\stdAdmin.PL23B\Documents\Result\file.xlsx");
                package.SaveAs(file);
            }

            Console.WriteLine("\nСчитывание параметром окончено.");
            Console.ReadLine();
        }
    }
}
