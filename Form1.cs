﻿using System;
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
 * In other to make it speak, you need to add System.Speak reference:
 * 1/ In Solution Explorer window, under Voice Bot Solution, there are several files such as Properties, 
 *    References, App.config, Form1.cs, Program.cs, etc.
 * 2/ Right click on References -> Add Reference...
 * 3/ Search (Ctrl+E) Speech
 * 4/ Click System.Speech -> OK
 * 
 */
namespace Voice_Bot
{
    public partial class Form1 : Form
    {
        string name = "Taylor";     //set my name
        static bool wake = false;   //set status to "Deaf"
        bool isLiked = true;        //set if I like the voice bot
        bool search = false;        //set if I want to search something in Google
        bool closeApp = false;      //set if I want to close application
        bool isRunning = false;     //set if the app is running
        bool exit = false;          //set if I want to exit voice bot

        //Declare Speech
        SpeechSynthesizer speech = new SpeechSynthesizer();
        Choices grammarList = new Choices();
        SpeechRecognitionEngine record = new SpeechRecognitionEngine();

        //Declare WeatherData
        static string city = "Auckland"; 
        WeatherData weather = new WeatherData(city);

        //Declare cursor position
        int positionY = Cursor.Position.Y;
        int positionX = Cursor.Position.X;
        int moveArea = 50; //the amount of unit your mouse will move

        //Declare Arduino port
        static SerialPort port = new SerialPort("COM5", 9600, Parity.None, 8, StopBits.One);
        static bool light = false; //set light status

        //Assign command and response lists
        string[] comList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Command.txt");
        string[] resList = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Response.txt");
        string[] googleSearch = File.ReadAllLines(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Text\Google search.txt");

        /**
         * Note: Command.txt and Response.txt are linking to each other by line number 
         *   (eg. line 2 in Command.txt is "how are you" and line 2 in Response.txt is "Good, and you").
         */

        public Form1()
        {
            //Define the contraints for speech recognition
            grammarList.Add(comList);
            grammarList.Add(googleSearch);
            Grammar grammar = new Grammar(new GrammarBuilder(grammarList));

            try
            {
                //Working with speech recogniser
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

            //This sentence will be spoken at the beginning.
            speech.Speak("Hohohoho, I am Santa Claus.");  

            InitializeComponent();
        }

        /**
         * Interacting method (getting command and responding)
         * When you say something assigned in grammarList, this method will be invoked.
         * Say "Wake" to make it work ("Listening")
         * Say "Sleep" to make it stop listening ("Deaf")
         * Whenever the voice bot talks something to you, its state will change
         * from "Listening" to "Deaf".
         */
        private void record_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Saving your speech and its position in ComList.
            String record = e.Result.Text;
            int count = Array.IndexOf(comList, record);

            /**
             * If you say "search for" -> search = true  
             * Then, when you say something in the grammarList, 
             * that word will be searched on the Internet.
             */
            if (search)
            {
                Process.Start("https://www.google.com/search?q=" + record);
                changeState(false, stateLabel); //Change state to "Deaf"
                search = false;
            }

            /**
             * If you say "close" -> close = true  
             * Then, when you say an application name in the grammarList 
             * (eg. Teams), this app will be killed.
             */
            if (closeApp)
            {
                if(record == "Google")
                {
                    isRunning = killProgram("chrome"); 
                }
                else
                {
                    isRunning = killProgram(record);
                }

                closeApp = false;
            }

            //Say "wake" to invoke the voice bot
            if (record == "wake")
            {
                changeState(true, stateLabel);
            }

            /**
             * If wake = true and search == false, response will be processed according to its type 
             *           (Multi responses, Date and Time, No response, One response, Exit) 
             *  and its keyword (eg. If response contains "Weather", it will show weather condition)
             */
            if (wake == true && search == false)
            {
                //Multi responses
                if (resList[count].StartsWith("+"))
                {
                    //Seperate multi responses
                    List<string> multiRes = resList[count].
                        Substring(resList[count].LastIndexOf(']') + 1).Split('/').ToList();

                    //If response contains a specific key word
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
                        if (isRunning) 
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
                //No response (the state will not automatically change to "Deaf")
                else if (resList[count].StartsWith("~"))
                {
                    if (record == "restart" || record == "update")
                    {
                        restart();
                    }

                    if (record == "close")
                    {
                        closeApp = true;
                    }

                    //Google search
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

                    //Mouse click
                    if (record == "click")
                    {
                        Voice_Bot.Peripherals.Mouse.DoMouseClick();
                    }

                    //Mouse right click
                    if (record == "right click")
                    {
                        Voice_Bot.Peripherals.Mouse.DoMouseRightClick();
                    }

                    //Send Keys
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

                    if (record == "enter")
                    {
                        SendKeys.Send("{ENTER}");
                    }

                    if (record == "change tab")
                    {
                        SendKeys.Send("^{TAB}");
                    }

                    if (record == "new tab")
                    {
                        SendKeys.Send("^T");
                    }

                    if (record == "close tab")
                    {
                        SendKeys.Send("^W");
                    }

                    if (record == "close previous tab")
                    {
                        SendKeys.Send("^+{TAB}");
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
                     * Open Arduino file (One_led) to see what will happen when 
                     * port.WriteLine("A") or port.WriteLine("B") is called.
                     * 
                     * Note: If you do not connect sensor to your computer, 
                     * it will cause bug.
                     */
                    if (record == "light on")
                    {
                        changeLightStatus("A");
                    }

                    if (record == "light off")
                    {
                        changeLightStatus("B");
                    }
                }
                //One response 
                else if (record != "exit" && record != "yes" && record != "no")
                {
                    if (resList[count].Contains("Sleep"))
                    {
                        changeState(false, stateLabel);
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
                        System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\Media\nezuko chan.wav");
                        player.Play();
                    }

                    //Change window status
                    if (resList[count].Contains("Minimise"))
                    {
                        this.WindowState = FormWindowState.Normal;
                    }

                    if (resList[count].Contains("Maximise"))
                    {
                        this.WindowState = FormWindowState.Maximized;
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

                    //It will send the link to search bar and then enter
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
                //Exit
                else
                {
                    if (record == "exit")
                    {
                        //If the light is on
                        if (light)
                        {
                            say(resList[count].Substring(resList[count].LastIndexOf(']') + 1));
                            exit = true;
                        }
                        else
                        {
                            say(resList[count + 1].Substring(resList[count + 1].LastIndexOf(']') + 1));
                            Environment.Exit(0);
                        }
                    }
                    
                    if (record == "yes" && exit == true) //turn off the light and exit
                    {
                        say(resList[count].Substring(resList[count].LastIndexOf(']') + 1));
                        changeLightStatus("B");
                        Environment.Exit(0);
                    }
                    
                    if (record == "no" && exit == true) //exit without turning off the light
                    {
                        say(resList[count].Substring(resList[count].LastIndexOf(']') + 1));
                        Environment.Exit(0);
                    }
                }
            }

            //add speech to "TALK" textbox
            textBox1.AppendText(record + "\n");

            //reset speech 
            record = ""; 
        }   

        public void restart()
        {
            //open another "Voice Bot.exe" - voice bot application
            Process.Start(@"C:\Users\pc\Desktop\PROject\C#\Voice Bot\bin\Debug\Voice Bot.exe");

            //Exit current voice bot application
            Environment.Exit(0);
        }

        public void say(String response)
        {
            //Speak "response"
            speech.Speak(response); //Text appears after speech

            //Write "response" in the "RESPONSE" textbox.
            textBox2.AppendText(response + "\n");

            //Set status to "Deaf"
            changeState(false, stateLabel);
        }

        public static void changeState(bool currentState, Label stateLabel)
        {
            //Change current state
            if (currentState)
            {
                wake = true;
                stateLabel.Text = "Status: Listening";
            }
            else
            {
                wake = false;
                stateLabel.Text = "Status: Deaf";
            }
        }

        public static bool killProgram(string processName)
        {
            //if application is not running
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

        public static void changeLightStatus(string id)
        {
            //Accessing port and update light status
            port.Open();
            port.WriteLine(id);
            port.Close();

            if (id == "A")
            {
                light = true;
            }
            else
            {
                light = false;
            }
        }
    }
}