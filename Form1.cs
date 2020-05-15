using System;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.IO.Ports;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using VoiceBot.Constants;

/**
 * Hi, this is C#.
 * 
 * If you don't understand any method or class, you can 
 * move your mouse to the key word (eg. "SpeechSynthesizer") 
 * to see its function or Ctrl + click to navigate to its place.
 * 
 * You also need to add System.Speak reference for using speech:
 * 1/ In Solution Explorer window, under Voice Bot Solution, there are several files such as Properties, 
 *    References, App.config, Form1.cs, Program.cs, etc.
 * 2/ Right click on References -> Add Reference...
 * 3/ Search (Ctrl+E) Speech
 * 4/ Click System.Speech -> OK
 * 
 */
namespace VoiceBot
{
    public partial class Form1 : Form
    {
        #region Private members
        string name = "Taylor";     
        static bool wake = false;   
        bool isLiked = true;        
        bool search = false;       
        bool closeApp = false;      
        bool appIsRunning = false;     
        bool exitVoiceBot = false;          

        SpeechSynthesizer speech = new SpeechSynthesizer();
        Choices grammarList = new Choices();
        SpeechRecognitionEngine record = new SpeechRecognitionEngine();

        static string city = "Auckland";
        WeatherData weather = new WeatherData(city);

        //Define cursor position
        static int positionY;
        static int positionX;
        int moveArea = 50; 

        static SerialPort port = new SerialPort("COM5", 9600, Parity.None, 8, StopBits.One);
        static bool lightStatus = false;

        //Assign command and response
        string[] comList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Command.txt");
        string[] resList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Response.txt");
        string[] googleSearch = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Google search.txt");

        /**
         * Note: Command.txt and Response.txt are linking to each other by line number 
         *   (eg. line 9 in Command.txt is "how are you" and line 9 in Response.txt is "Good, and you").
         */

        public Form1()
        {
            //Define the contraints for speech recognition
            grammarList.Add(comList);
            grammarList.Add(googleSearch);
            Grammar grammar = new Grammar(new GrammarBuilder(grammarList));

            try
            {
                record.RequestRecognizerUpdate();
                record.LoadGrammar(grammar);
                record.SpeechRecognized += RecordSpeechRecognized;
                record.SetInputToDefaultAudioDevice();
                record.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch
            {
                return;
            }

            speech.SelectVoiceByHints(VoiceGender.Female);
            //speech.Speak("Hohohoho, I am Santa Claus.");

            InitializeComponent();
        }

        #endregion

        #region Methods
        public void Restart()
        {
            Process.Start(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\bin\Debug\Voice Bot.exe");
            Environment.Exit(0);
        }

        public void Say(string response)
        {
            speech.Speak(response); 
            textBox2.AppendText(response + "\n");
            ChangeState(false, stateLabel);
        }

        public static void ChangeState(bool nextState, Label stateLabel)
        {
            wake = nextState;

            if (nextState)
            {
                stateLabel.Text = "State: Listening";
            }
            else
            {
                stateLabel.Text = "State: Deaf";
            }
        }

        public static bool KillProgram(string processName)
        {
            //if this app is not running
            if (Process.GetProcessesByName(processName).Length == 0)
            {
                return false;
            }
            else
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    process.Kill(); 
                }

                return true;
            }
        }

        public static void UpdateMousePosition()
        {
            positionY = Cursor.Position.Y;
            positionX = Cursor.Position.X;
        }

        public static void ChangeLightStatus(string id)
        {
            port.Open();
            port.WriteLine(id);
            port.Close();

            if (id == "A")
            {
                lightStatus = true;
            }
            else
            {
                lightStatus = false;
            }
        }

        #endregion

        #region Private Method

        /**
         * Interacting method (getting command and responding)
         * When you say something assigned in grammarList, this method will be invoked.
         * Say "Wake" to make it work ("Listening")
         * Say "Sleep" to make it stop listening ("Deaf")
         * Whenever the voice bot talks something to you, its state will change
         * from "Listening" to "Deaf" -> Stop listening.
         */
        private void RecordSpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Saving your speech
            string record = e.Result.Text;

            //Find record's position in comList array to get resList[count]
            int count = Array.IndexOf(comList, record);

            if (search)
            {
                Process.Start("https://www.google.com/search?q=" + record); 
                ChangeState(false, stateLabel); 
                search = false;
            }

            if (closeApp)
            {
                if (record == ResponseConstant.Google)
                {
                    appIsRunning = KillProgram("chrome");
                }
                else
                {
                    appIsRunning = KillProgram(record);
                }

                closeApp = false;
            }

            if (record == ResponseConstant.Wake)
            {
                ChangeState(true, stateLabel);
            }

            /**
             * DOING TASKS when wake = true and search = false.
             * From here, every response will be processed according to its type
             * (Multi responses, Date and Time, No response, One response, Exit)
             * and its function (eg. Weather), identified through Response.txt.
             */
            if (wake == true && search == false)
            {
                if (resList[count].StartsWith("+"))
                {
                    HandleMultipleResponse(count);
                }
                else if (resList[count].StartsWith("-"))
                {
                    HandleDateTimeResponse(count);
                }
                //No response - the state will not automatically change to "Deaf"
                else if (resList[count].StartsWith("~"))
                {
                    HandleNoResponse(record);
                }
                else if (record != ResponseConstant.Exit && record != ResponseConstant.Yes && record != ResponseConstant.No)
                {
                    HandleOtherResponse(count);
                }
                else
                {
                    if (record == ResponseConstant.Exit)
                    {
                        HandleExit(count);
                    }

                    if (record == ResponseConstant.Yes && exitVoiceBot == true) //turn off the light and exit
                    {
                        HandleYes(count);

                    }

                    if (record == ResponseConstant.No && exitVoiceBot == true) //exit without turning off the light
                    {
                        HandleNo(count);
                    }
                }
            }

            textBox1.AppendText(record + "\n");
        }

        private void HandleMultipleResponse(int responsePosition)
        {
            //Seperate multi responses by '/'
            List<string> multiRes = resList[responsePosition].
                Substring(resList[responsePosition].LastIndexOf(']') + 1).Split('/').ToList();

            string response = resList[responsePosition];

            string responseResult = response.Substring(response.IndexOf('[') + 1, response.IndexOf(']') - response.IndexOf('[') - 1);

            switch (responseResult)
            {
                case ResponseConstant.Greetings:

                    Random random = new Random();
                    Say(multiRes[random.Next(multiRes.Count())]);

                    break;

                case ResponseConstant.IsLike:

                    if (isLiked)
                    {
                        Say(multiRes[0]);
                    }
                    else
                    {
                        Say(multiRes[1] + name + multiRes[2]);
                    }

                    break;

                case ResponseConstant.Delete:

                    Say(multiRes[0]);
                    speech.SelectVoiceByHints(VoiceGender.Male);

                    Say(multiRes[1]);
                    speech.SelectVoiceByHints(VoiceGender.Female);

                    Say(multiRes[2]);

                    break;

                case ResponseConstant.Weather:

                    Say(multiRes[0]);
                    weather.CheckWeather();
                    Say(multiRes[1] + weather.Condition + multiRes[2] + name);

                    break;

                case ResponseConstant.Temp:

                    Say(multiRes[0]);
                    weather.CheckWeather();
                    Say(multiRes[1] + weather.Temperature.ToString() + multiRes[2]);

                    break;

                case ResponseConstant.CloseApp:

                    if (appIsRunning)
                    {
                        Say(multiRes[0]);
                    }
                    else
                    {
                        Say(multiRes[1]);
                    }

                    break;

                default:
                    break;
            }
        }

        private void HandleDateTimeResponse(int responsePosition)
        {
            string response = resList[responsePosition];

            string responseResult = response.Substring(response.IndexOf('[') + 1, response.IndexOf(']') - response.IndexOf('[') - 1);

            switch (responseResult)
            {
                case ResponseConstant.Date:

                    Say(DateTime.Now.ToString("M/dd/yyyy"));

                    break;

                case ResponseConstant.Time:

                    Say(DateTime.Now.ToString("h:mm tt"));

                    break;

                default:
                    break;
            }
        }

        private void HandleNoResponse(string record)
        {
            switch(record)
            {
                case ResponseConstant.Restart:

                    Restart();

                    break;

                case ResponseConstant.Update:

                    Restart();

                    break;

                case ResponseConstant.Close:
                    
                    closeApp = true;

                    break;

                case ResponseConstant.Search:

                    search = true;

                    break;

                case ResponseConstant.Down:

                    UpdateMousePosition();
                    VoiceBot.Peripherals.Mouse.MoveToPoint(positionX, positionY += moveArea);

                    break;

                case ResponseConstant.Up:

                    UpdateMousePosition();
                    VoiceBot.Peripherals.Mouse.MoveToPoint(positionX, positionY -= moveArea);

                    break;

                case ResponseConstant.Left:

                    UpdateMousePosition();
                    VoiceBot.Peripherals.Mouse.MoveToPoint(positionX -= moveArea, positionY);

                    break;
                    
                case ResponseConstant.Right:

                    UpdateMousePosition();
                    VoiceBot.Peripherals.Mouse.MoveToPoint(positionX += moveArea, positionY);

                    break;

                case ResponseConstant.Click:

                    VoiceBot.Peripherals.Mouse.DoMouseClick();

                    break;

                case ResponseConstant.RightClick:

                    VoiceBot.Peripherals.Mouse.DoMouseRightClick();

                    break;

                case ResponseConstant.Play:
                
                    SendKeys.Send(" ");

                    break;

                case ResponseConstant.Pause:
                    
                    SendKeys.Send(" ");

                    break;

                case ResponseConstant.Next:
                
                    SendKeys.Send("^{RIGHT}");

                    break;
                    
                case ResponseConstant.Last:

                    SendKeys.Send("^{LEFT}");

                    break;
                    
                case ResponseConstant.Enter:

                    SendKeys.Send("^{ENTER}");

                    break;
        
                case ResponseConstant.ChangeTab:

                    SendKeys.Send("^{TAB}");

                    break;
        
                case ResponseConstant.NewTab:

                    SendKeys.Send("^T");

                    break;
                            
                case ResponseConstant.CloseTab:

                    SendKeys.Send("W");

                    break;

                case ResponseConstant.ClosePreviousTab:

                    SendKeys.Send("^+{TAB}");
                    SendKeys.Send("^W");

                    break;

               case ResponseConstant.CloseFollowingTab:

                    SendKeys.Send("^{TAB}");
                    SendKeys.Send("^W");

                    break;            
                
                case ResponseConstant.OpenSecretTab:

                    SendKeys.Send("^+N");

                    break;

                /**
                    * Open Arduino file (One_led) to see what will happen when 
                    * port.WriteLine("A") or port.WriteLine("B") is called.
                */
                case ResponseConstant.LightOn:

                    ChangeLightStatus("A");

                    break;

                case ResponseConstant.LightOff:

                    ChangeLightStatus("B");

                    break;

                default:
                    break;
            }
        }

        private void HandleOtherResponse(int responsePosition)
        {
            string response = resList[responsePosition];

            string responseResult = response.Substring(response.IndexOf('[') + 1, response.IndexOf(']') - response.IndexOf('[') - 1);

            switch (responseResult)
            {
                case ResponseConstant.Sleep:

                    ChangeState(false, stateLabel);

                    break;

                case ResponseConstant.Like:

                    isLiked = true;

                    break;
                
                case ResponseConstant.Hate:

                    isLiked = false;

                    break;

                case ResponseConstant.Nezuko:

                    //Open wav file
                    System.Media.SoundPlayer player = new System.Media.
                        SoundPlayer(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Media\nezuko chan.wav");
                    player.Play();

                    break;
                    
                case ResponseConstant.Minimise:

                    this.WindowState = FormWindowState.Normal;

                    break;

                case ResponseConstant.Maximise:

                    this.WindowState = FormWindowState.Maximized;

                    break;
                
                case ResponseConstant.Google:

                    Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");

                    break;

                case ResponseConstant.Spotify:

                    Process.Start(@"C:\Users\pc\AppData\Roaming\Spotify\Spotify.exe");

                    break;
                
                case ResponseConstant.Teams:

                    Process.Start(@"C:\Users\pc\AppData\Local\Microsoft\Teams\current\Teams.exe");

                    break;
                
                case ResponseConstant.YouTube:

                    SendKeys.Send("https://www.youtube.com");
                    SendKeys.Send("{ENTER}");

                    break;

                case ResponseConstant.MaximiseScreen:

                    SendKeys.Send("f");

                    break;

                default:
                    break;
            }

            //If response contains "["
            if (resList[responsePosition].Contains("["))
            {
                resList[responsePosition] = resList[responsePosition].Substring(resList[responsePosition].LastIndexOf(']') + 1);
            }

            Say(resList[responsePosition]);
        }

        private void HandleExit(int responsePosition)
        {
            string response = resList[responsePosition];

            //If the light is on, ask whether user wants to turn it off
            if (lightStatus)
            {
                Say(response.Substring(response.LastIndexOf(']') + 1));
                exitVoiceBot = true;
            }
            else
            {
                Say(resList[responsePosition + 1].Substring(resList[responsePosition + 1].LastIndexOf(']') + 1));
                Environment.Exit(0);
            }
        }

        private void HandleYes(int responsePosition)
        {
            string response = resList[responsePosition];

            Say(response.Substring(response.LastIndexOf(']') + 1));
            ChangeLightStatus("B");

            Environment.Exit(0);
        }

        private void HandleNo(int responsePosition)
        {
            string response = resList[responsePosition];

            Say(response.Substring(response.LastIndexOf(']') + 1));

            Environment.Exit(0);
        }

        #endregion
    }
}