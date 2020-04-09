using System;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.IO.Ports;
using System.IO;
using System.Collections.Generic;
using System.Linq;

/**
 * Hi, this is C#.
 * 
 * If you don't understand any method or class, you can 
 * move your mouse to the key word (eg. move your mouse to the word "SpeechSynthesizer" 
 * to see its function) or Ctrl + click to navigate to its location.
 * 
 */
namespace Voice_Bot
{
    public partial class Form1 : Form
    {
        string name = "Taylor";     //set my name
        bool wake = false;          //set status to "Deaf"
        bool isLiked = true;        //set if I like the voice bot
        bool search = false;        //set if I want to search something in Google
        bool close = false;         //set if I want to close application
        bool isRunning = false;     //set if the app is running

        //WeatherData
        static string city = "Auckland";   //set location for checking weather
        WeatherData weather = new WeatherData(city);

        //Declare cursor position
        int positionY = Cursor.Position.Y;
        int positionX = Cursor.Position.X;
        int moveArea = 50; //the amount of unit your mouse will move

        //Declare Arduino port
        SerialPort port = new SerialPort("COM5", 9600, Parity.None, 8, StopBits.One);

        /**
         * Assign command and response lists
         * You can add/delete/change any commands, reponses, and searching words in those text files.
         * Note: be careful when modifying command and response, because they are linking to 
         *       each other by line number. (eg. line 2 in Command.txt is "how are you" and 
         *       line 2 in Response.txt is "Good, and you")
         */
        string[] comList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Command.txt");
        string[] resList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Response.txt");
        string[] googleSearch = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Google search.txt");

        //SpeechSynthesizer 
        SpeechSynthesizer speech = new SpeechSynthesizer();

        //Speech record
        Choices grammarList = new Choices();
        SpeechRecognitionEngine record = new SpeechRecognitionEngine();

        public Form1()
        {
            grammarList.Add(comList);
            grammarList.Add(googleSearch);
            Grammar grammar = new Grammar(new GrammarBuilder(grammarList));

            try
            {
                record.RequestRecognizerUpdate();
                record.LoadGrammar(grammar);
                record.SpeechRecognized += record_SpeechRecognized;
                record.SetInputToDefaultAudioDevice();
                record.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch
            {
                return;
            }

            //Set voice gender to female
            speech.SelectVoiceByHints(VoiceGender.Female);

            //Whenever you open or reset the voice bot, this sentence will be spoken.
            speech.Speak("Hohohoho, I am Santa Claus.");

            InitializeComponent();
        }

        /**
         * Interacting method (getting command and responding)
         * Say "Wake" to make it work ("Listening")
         * Say "Sleep" to make it stop listening ("Deaf")
         * Whenever the voice bot talks something to you, its state will change
         * from "Listening" to "Deaf".
         */
        private void record_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Saving your speech
            String record = e.Result.Text;

            //Get comList's element position according to your speech
            int count = Array.IndexOf(comList, record);

            if(search)
            {
                Process.Start("https://www.google.com/search?q=" + record);
                wake = false;
                search = false;
            }

            if(close)
            {
                if(record == "Google")
                {
                    isRunning = killProgram("chrome");
                }
                else
                {
                    isRunning = killProgram(record);
                }

                close = false;
            }

            if(record == "wake")
            {
                wake = true;
            }

            if (wake == true && search == false)
            {
                //Multi responses
                if (resList[count].StartsWith("+"))
                {
                    //Seperate multi responses
                    List<string> multiRes = resList[count].
                        Substring(resList[count].LastIndexOf(']') + 1).Split('/').ToList();

                    if (resList[count].Contains("Greetings"))
                    {
                        Random random = new Random();
                        say(multiRes[random.Next(multiRes.Count())]);
                    }

                    if (resList[count].Contains("Like?"))
                    {
                        if (isLiked)
                        {
                            say(multiRes[0]);
                        }
                        else
                        {
                            say(multiRes[1] + name + multiRes[2]);
                        }
                    }

                    if (resList[count].Contains("Delete"))
                    {
                        say(multiRes[0]);
                        speech.SelectVoiceByHints(VoiceGender.Male);
                        say(multiRes[1]);
                        speech.SelectVoiceByHints(VoiceGender.Female);
                        say(multiRes[2]);
                    }

                    if (resList[count].Contains("Weather"))
                    {
                        say(multiRes[0]);
                        weather.CheckWeather();
                        say(multiRes[1] + weather.Condition + multiRes[2] + name);
                    }

                    if (resList[count].Contains("Temp"))
                    {
                        say(multiRes[0]);
                        weather.CheckWeather();
                        say(multiRes[1] + weather.Temperature.ToString() + multiRes[2]);
                    }

                    if (resList[count].Contains("Close"))
                    {
                        if(isRunning)
                        {
                            say(multiRes[0]);
                        }
                        else
                        {
                            say(multiRes[1]);
                        }
                    }
                }
                //Date and Time
                else if (resList[count].StartsWith("-"))
                {
                    if (resList[count].Contains("Date"))
                    {
                        say(DateTime.Now.ToString("M/dd/yyyy"));
                    }
                 
                    if (resList[count].Contains("Time"))
                    {
                        say(DateTime.Now.ToString("h:mm tt"));
                    }
                }
                /**
                  * No response
                  *     Note: The state will not automatically change to "Deaf" in "No response".
                  */
                else if (resList[count].StartsWith("~")) 
                {
                    if (record == "restart" || record == "update")
                    {
                        restart();
                    }

                    if (record == "wake")
                    {
                        wake = true;
                        label3.Text = "State: Listening"; //Change text in label3
                    }

                    if (record == "close")
                    {
                        close = true;
                    }

                    if(record == "search for")
                    {
                        search = true;
                    }

                    if (record == "play" || record == "pause")
                    {
                        SendKeys.Send(" ");
                    }

                    if (record == "next")
                    {
                        SendKeys.Send("^{RIGHT}");
                    }

                    if (record == "last")
                    {
                        SendKeys.Send("^{LEFT}");
                    }

                    if (record == "search for")
                    {
                        search = true;
                    }

                    //Changing cursor location
                    if (record == "down")
                    {
                        Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX, positionY += moveArea);
                    }

                    if (record == "up")
                    {
                        Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX, positionY -= moveArea);
                    }

                    if (record == "left")
                    {
                        Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX -= moveArea, positionY);
                    }

                    if (record == "right")
                    {
                        Voice_Bot.Peripherals.Mouse.MoveToPoint(positionX += moveArea, positionY);
                    }

                    if (record == "click")
                    {
                        Voice_Bot.Peripherals.Mouse.DoMouseClick();
                    }

                    //Send Keys
                    if (record == "enter")
                    {
                        SendKeys.Send("{ENTER}");
                    }

                    if (record == "change tab")
                    {
                        SendKeys.Send("^{TAB}"); // Ctrl Tab
                    }

                    if (record == "new tab")
                    {
                        SendKeys.Send("^T"); //Ctrl T
                    }

                    if (record == "close tab")
                    {
                        SendKeys.Send("^W");
                    }

                    if (record == "close previous tab")
                    {
                        SendKeys.Send("^+{TAB}"); //Ctrl Shift Tab
                        SendKeys.Send("^W");
                    }

                    if (record == "close following tab")
                    {
                        SendKeys.Send("^{TAB}");
                        SendKeys.Send("^W");
                    }

                    if (record == "open secret tab")
                    {
                        SendKeys.Send("^+N");
                    }

                    /**
                     * Open Arduino file to see what will happen when 
                     * port.WriteLine("A") or port.WriteLine("B") is called.
                     * 
                     * Note: If you do not connect sensor to your computer,
                     * it will cause bug.
                     */
                    if (record == "light on")
                    {
                        port.Open();
                        port.WriteLine("A");
                        port.Close();
                    }

                    if (record == "light off")
                    {
                        port.Open();
                        port.WriteLine("B");
                        port.Close();
                    }
                }
                //One response only
                else
                {
                    if(resList[count].Contains("Sleep"))
                    {
                        wake = false;
                        label3.Text = "State: Deaf";
                    }

                    if (resList[count].Contains("Like"))
                    {
                        isLiked = true;
                    }

                    if (resList[count].Contains("Hate"))
                    {
                        isLiked = false;
                    }

                    if (resList[count].Contains("Nezuko"))
                    {
                        //Open wav file
                        System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\nezuko chan.wav");
                        player.Play();
                    }

                    //Change Voice Bot window status
                    if (resList[count].Contains("Minimise"))
                    {
                        this.WindowState = FormWindowState.Normal;
                    }

                    if (resList[count].Contains("Maximise"))
                    {
                        this.WindowState = FormWindowState.Maximized;
                    }

                    //Disrupt voice bot speech
                    if (resList[count].Contains("Stop"))
                    {
                        speech.SpeakAsyncCancelAll();
                    }

                    //Accessing applications
                    if (resList[count].Contains("Google"))
                    {
                        Process.Start(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe");
                    }

                    if (resList[count].Contains("Spotify"))
                    {
                        Process.Start(@"C:\Users\pc\AppData\Roaming\Spotify\Spotify.exe");
                    }

                    if (resList[count].Contains("Teams"))
                    {
                        Process.Start(@"C:\Users\pc\AppData\Local\Microsoft\Teams\current\Teams.exe");
                    }

                    //It will send the link to the search bar, and then enter
                    if (resList[count].Contains("YouTube"))
                    {
                        SendKeys.Send("https://www.youtube.com");
                        SendKeys.Send(((char)13).ToString()); //enter
                    }

                    if (resList[count].Contains("Max screen"))
                    {
                        SendKeys.Send("f");
                    }

                    //If response contains []
                    if (resList[count].Contains("["))
                    {
                        resList[count] = resList[count].Substring(resList[count].LastIndexOf(']') + 1);
                    }

                    say(resList[count]);
                }
            }

            textBox1.AppendText(record + "\n"); //add speech to "TALK" textbox

            record = ""; //reset speech 
        }   

        public void restart()
        {
            //open another "Voice Bot.exe" - voice bot application
            Process.Start(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\bin\Debug\Voice Bot.exe");

            //Exit current voice bot application
            //this 0 parameter indicates for the process's sucessful completion.
            Environment.Exit(0);
        }

        public void say(String response)
        {
            //Speak "response"
            speech.Speak(response); //Text appears after speech

            //Write "response" in the "RESPONSE" textbox.
            textBox2.AppendText(response + "\n");

            //Set status to "Deaf"
            wake = false;
            label3.Text = "Status: Deaf";
        }

        public static bool killProgram(string processName)
        {
            //if the application is not running
            if(Process.GetProcessesByName(processName).Length == 0)
            {
                return false;
            }
            else
            {
                foreach(Process process in Process.GetProcessesByName(processName))
                {
                    process.Kill();
                }

                return true;
            }
        }
    }
}