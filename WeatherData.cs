using System.Net;
using System.Xml;

namespace Voice_Bot
{
    class WeatherData
    {
        private string city;
        private float temp;
        private string condition;

        public WeatherData(string City)
        {
            city = City;
        }

        //Update weather
        public void CheckWeather()
        {
            WeatherAPI DataAPI = new WeatherAPI(City);
            temp = DataAPI.GetTemp();
            condition = DataAPI.GetCondition();
        }

        public string City { get => city; set => city = value; }
        public float Temperature { get => temp; set => temp = value; }
        public string Condition { get => condition; set => condition = value; }
    }

    class WeatherAPI
    {
        public WeatherAPI(string city)
        {
            SetCurrentURL(city);
            xmlDocument = GetXML(CurrentURL);
        }

        public float GetTemp()
        {
            //Get value in temperature
            XmlNode temp_node = xmlDocument.SelectSingleNode("//temperature");
            XmlAttribute temp_value = temp_node.Attributes["value"];
            string temp_string = temp_value.Value;

            return float.Parse(temp_string);
        }

        public string GetCondition()
        {
            //Get value in weather
            XmlNode condition_node = xmlDocument.SelectSingleNode("//weather");
            XmlAttribute condition_value = condition_node.Attributes["value"];
            string condition_string = condition_value.Value;

            return condition_string;
        }

        //You need to create your own APIKEY in Openweathermap if you want to use its API
        private const string APIKEY = "YOUR APIKEY HERE";
        private string CurrentURL;
        private XmlDocument xmlDocument;

        private void SetCurrentURL(string location)
        {
            CurrentURL = "http://api.openweathermap.org/data/2.5/weather?q=" 
                + location + "&mode=xml&units=metric&APPID=" + APIKEY;
        }

        private XmlDocument GetXML(string CurrentURL)
        {
            using (WebClient client = new WebClient())
            {
                string xmlContent = client.DownloadString(CurrentURL);
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlContent);
                return xmlDocument;
            }
        }
    }
}
