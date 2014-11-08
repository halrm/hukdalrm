using BrightIdeasSoftware;
using IWshRuntimeLibrary;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string emailDef    = "hukdalrm@gmail.com";
        string passwordDef = "";

        string version     = "0.8";

        SoundPlayer alarm_Deal      = new SoundPlayer();
        SoundPlayer alarm_Coupon    = new SoundPlayer();
        SoundPlayer alarm_Freebie   = new SoundPlayer();

        List<Deal>  reportedDeals    = new List<Deal>();
        List<Deal>  reportedCoupons  = new List<Deal>();
        List<Deal>  reportedFreebies = new List<Deal>();

        Settings    settings  = new Settings();

        bool showReports = false;
        bool soundOn     = true;
        bool alarmsOn    = true;

        public class Settings
        {
            public string emailAddress
            { get; set; }

            public bool startMinimized
            { get; set; }

            public Subsettings dealsSettings
            { get; set; }

            public Subsettings couponsSettings
            { get; set; }
            
            public Subsettings freebiesSettings
            { get; set; }
        }
        
        public class Subsettings
        {
            public int minTempAll
            { get; set; }

            public int maxAgeAll
            { get; set; }
            
            public int minTempAlarm
            { get; set; }

            public int maxAgeAlarm
            { get; set; }
            
            public int pages
            { get; set; }

            public int timer
            { get; set; }

            public bool popup
            { get; set; }

            public bool audio
            { get; set; }

            public string audioFile
            { get; set; }

            public bool email
            { get; set; }

            public bool region
            { get; set; }

            public string regionList
            { get; set; }

            public bool keywordAlarm
            { get; set; }

            public string keywordList
            { get; set; }  
        }

        class Deal
        {
            public Deal()
            {

            }

            public Deal(string title, double age, int heat, double price, string link)
            {
                this.title = title;
                this.age = age;
                this.heat = heat;
                this.price = price;
                this.link = link;
            }

            public string Title
            {
                get { return title; }
                set { title = value; }
            }
            private string title;

            public double Age
            {
                get { return age; }
                set { age = value; }
            }
            private double age;

            public int Heat
            {
                get { return heat; }
                set { heat = value; }
            }
            private int heat;

            public double Price
            {
                get { return price; }
                set { price = value; }
            }
            private double price;

            public string Link
            {
                get { return link; }
                set { link = value; }
            }
            private string link;
        }
        
        public Form1()
        {
            InitializeComponent();            

            this.Text = "hukdalrm " + version;

            if (System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\hukdalrm.lnk"))
                checkBox14.Checked = true;
            else checkBox14.Checked = false;

            initSoundFiles();

            initSettings();                   

            timer1.Interval = (int)numericUpDown8.Value * 1000;
            timer2.Interval = (int)numericUpDown9.Value * 1000;
            timer3.Interval = (int)numericUpDown17.Value * 1000;

            toolTip1.SetToolTip(button1, "Zeige zurückliegende Meldungen an");
            toolTip1.SetToolTip(button4, "Aktualisiere Daten");
            toolTip1.SetToolTip(button9, "Deaktiviere alle Alarme (dh. kein Popup, kein Sound, keine Mail)");
            toolTip1.SetToolTip(button3, "Deaktiviere alle Audio-Alarme");
            toolTip1.SetToolTip(button5, "Sichere alle Einstellungen");

            alarm_Deal.SoundLocation = "Alarmsounds\\" + comboBox2.Text + ".wav";
            alarm_Deal.Load();

            alarm_Coupon.SoundLocation = "Alarmsounds\\" + comboBox1.Text + ".wav";
            alarm_Coupon.Load();

            alarm_Freebie.SoundLocation = "Alarmsounds\\" + comboBox3.Text + ".wav";
            alarm_Freebie.Load();           

        }

        private void initSoundFiles()
        {
            foreach (string s in System.IO.Directory.EnumerateFiles("Alarmsounds"))
                if (s.Contains(".wav"))
                {
                    string st = s.Substring(12, s.Length - 12 - 4);
                    comboBox2.Items.Add(st);
                    comboBox1.Items.Add(st);
                    comboBox3.Items.Add(st);
                }
        }

        private void initSettings()
        {
            settings.dealsSettings = new Subsettings();
            settings.couponsSettings = new Subsettings();
            settings.freebiesSettings = new Subsettings();

            settings = DeserializeFromXML("settings.xml");

            if ((settings.startMinimized) && (checkBox14.Checked))
                this.WindowState = FormWindowState.Minimized;

            //GENERAL-SETTINGS
            checkBox4.Checked = settings.startMinimized;
            textBox2.Text = settings.emailAddress;

            //DEAL-SETTINGS
            numericUpDown1.Value = settings.dealsSettings.minTempAll;
            numericUpDown2.Value = settings.dealsSettings.maxAgeAll;           

            numericUpDown6.Value = settings.dealsSettings.minTempAlarm;
            numericUpDown5.Value = settings.dealsSettings.maxAgeAlarm;

            numericUpDown7.Value = settings.dealsSettings.pages;
            numericUpDown8.Value = settings.dealsSettings.timer;

            checkBox3.Checked = settings.dealsSettings.popup;
            checkBox2.Checked = settings.dealsSettings.audio;

            comboBox2.Text = settings.dealsSettings.audioFile;

            checkBox1.Checked = settings.dealsSettings.email;

            checkBox13.Checked = settings.dealsSettings.region;
            textBox4.Text = settings.dealsSettings.regionList;

            checkBox15.Checked = settings.dealsSettings.keywordAlarm;
            textBox5.Text = settings.dealsSettings.keywordList;

            //COUPON-SETTINGS
            numericUpDown16.Value = settings.couponsSettings.minTempAll;
            numericUpDown15.Value = settings.couponsSettings.maxAgeAll;

            numericUpDown13.Value = settings.couponsSettings.minTempAlarm;
            numericUpDown12.Value = settings.couponsSettings.maxAgeAlarm;

            numericUpDown10.Value = settings.couponsSettings.pages;
            numericUpDown9.Value = settings.couponsSettings.timer;

            checkBox6.Checked = settings.couponsSettings.popup;
            checkBox7.Checked = settings.couponsSettings.audio;

            comboBox1.Text = settings.couponsSettings.audioFile;

            checkBox8.Checked = settings.couponsSettings.email;

            checkBox9.Checked = settings.couponsSettings.region;
            textBox3.Text = settings.couponsSettings.regionList;

            checkBox17.Checked = settings.couponsSettings.keywordAlarm;
            textBox7.Text = settings.couponsSettings.keywordList;

            //FREEBIE-SETTINGS
            numericUpDown24.Value = settings.freebiesSettings.minTempAll;
            numericUpDown23.Value = settings.freebiesSettings.maxAgeAll;

            numericUpDown21.Value = settings.freebiesSettings.minTempAlarm;
            numericUpDown20.Value = settings.freebiesSettings.maxAgeAlarm;

            numericUpDown18.Value = settings.freebiesSettings.pages;
            numericUpDown17.Value = settings.freebiesSettings.timer;

            checkBox10.Checked = settings.freebiesSettings.popup;
            checkBox11.Checked = settings.freebiesSettings.audio;

            comboBox3.Text = settings.freebiesSettings.audioFile;

            checkBox12.Checked = settings.freebiesSettings.email;

            checkBox5.Checked = settings.freebiesSettings.region;
            textBox1.Text = settings.freebiesSettings.regionList;

            checkBox16.Checked = settings.freebiesSettings.keywordAlarm;
            textBox6.Text = settings.freebiesSettings.keywordList;
                        
        }

        private void saveSettings()
        {
            //DEAL-SETTINGS
            settings.dealsSettings.minTempAll = (int)numericUpDown1.Value;
            settings.dealsSettings.maxAgeAll = (int)numericUpDown2.Value;
            settings.dealsSettings.minTempAlarm = (int)numericUpDown6.Value;
            settings.dealsSettings.maxAgeAlarm = (int)numericUpDown5.Value;

            settings.dealsSettings.pages = (int)numericUpDown7.Value;
            settings.dealsSettings.timer = (int)numericUpDown8.Value;

            settings.dealsSettings.popup = checkBox3.Checked;
            settings.dealsSettings.audio = checkBox2.Checked;

            settings.dealsSettings.audioFile = comboBox2.Text;

            settings.dealsSettings.email = checkBox2.Checked;

            settings.dealsSettings.region = checkBox13.Checked;
            settings.dealsSettings.regionList = textBox4.Text;

            settings.dealsSettings.keywordAlarm = checkBox15.Checked;
            settings.dealsSettings.keywordList = textBox5.Text;

            //COUPON-SETTINGS
            settings.couponsSettings.minTempAll = (int)numericUpDown16.Value;
            settings.couponsSettings.maxAgeAll = (int)numericUpDown15.Value;
            settings.couponsSettings.minTempAlarm = (int)numericUpDown13.Value;
            settings.couponsSettings.maxAgeAlarm = (int)numericUpDown12.Value;

            settings.couponsSettings.pages = (int)numericUpDown10.Value;
            settings.couponsSettings.timer = (int)numericUpDown9.Value;

            settings.couponsSettings.popup = checkBox6.Checked;
            settings.couponsSettings.audio = checkBox7.Checked;

            settings.couponsSettings.audioFile = comboBox1.Text;

            settings.couponsSettings.email = checkBox8.Checked;

            settings.couponsSettings.region = checkBox9.Checked;
            settings.couponsSettings.regionList = textBox3.Text;

            settings.couponsSettings.keywordAlarm = checkBox17.Checked;
            settings.couponsSettings.keywordList = textBox7.Text;

            //FREEBIE-SETTINGS
            settings.freebiesSettings.minTempAll = (int)numericUpDown24.Value;
            settings.freebiesSettings.maxAgeAll = (int)numericUpDown23.Value;
            settings.freebiesSettings.minTempAlarm = (int)numericUpDown21.Value;
            settings.freebiesSettings.maxAgeAlarm = (int)numericUpDown20.Value;

            settings.freebiesSettings.pages = (int)numericUpDown18.Value;
            settings.freebiesSettings.timer = (int)numericUpDown17.Value;

            settings.freebiesSettings.popup = checkBox10.Checked;
            settings.freebiesSettings.audio = checkBox11.Checked;

            settings.freebiesSettings.audioFile = comboBox3.Text;

            settings.freebiesSettings.email = checkBox12.Checked;

            settings.freebiesSettings.region = checkBox5.Checked;
            settings.freebiesSettings.regionList = textBox1.Text;

            settings.freebiesSettings.keywordAlarm = checkBox16.Checked;
            settings.freebiesSettings.keywordList = textBox6.Text;

            //GENERAL-SETTINGS
            settings.emailAddress = textBox2.Text;
            settings.startMinimized = checkBox4.Checked;

            SerializeToXML(settings);
        }

        private List<string> grabPages(string section, int pageCount)
        {
            this.Cursor = Cursors.WaitCursor;
            List<string> pages = new List<string>();

            string urlAddress = "";

            if (section.Equals("Deal"))
                urlAddress = "http://hukd.mydealz.de/all/deals/new?page=";
            else if (section.Equals("Gutschein"))
                urlAddress = "http://hukd.mydealz.de/all/gutscheine/new?page=";
            else if (section.Equals("Freebie"))
                urlAddress = "http://hukd.mydealz.de/all/freebies/new?page=";

            try
            {
                for (int i = 1; i <= pageCount; i++)
                {
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress + i);
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        Stream receiveStream = response.GetResponseStream();
                        StreamReader readStream = null;
                        if (response.CharacterSet == null)
                            readStream = new StreamReader(receiveStream);
                        else
                            readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                        string data = readStream.ReadToEnd();
                        pages.Add(data);
                        response.Close();
                        readStream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.Cursor = Cursors.Arrow;
            return pages;
        }

        private List<Deal> getDeals(List<string> pages
                                    , string section
                                    , List<Deal> reportedItems
                                    , int minHeat
                                    , int maxAge
                                    , int minHeatAlarm
                                    , int maxAgeAlarm
                                    , bool emailAlert
                                    , bool popupAlert
                                    , bool audioAlert
                                    , bool regionFilterOn
                                    , string cities
                                    , bool keywordFilterOn
                                    , string keywords
                                   )
        {
            List<Deal> deals = new List<Deal>();

            string titleStart= "<a class=\"section-title-link space--after-2";
            string titleEnd = "</strong></div><p class=\"thread-meta\">";
            if (section.Equals("Freebie"))
                titleEnd="</h2></div><p class=\"thread-meta\">";

            string heatStart = "degreeValue--spaced tGrid-cell space--right-2 text--vAlign-middle\">";
                             //   degreeValue--spaced tGrid-cell space--right-2 text--vAlign-middle">
            string heatEnd   = "<span class=\"degreeValue";

            string linkStart = "href=\"";
            string linkEnd = "\">";

            string ageStart = "<span class=\"thread-time text--muted text--overflow-nowrap\">";
            string ageEnd = "</span></p></header>";

            string title; string titleStr; int titleStartIndex, titleEndIndex;
            int age; string ageStr; int ageStartIndex, ageEndIndex;
            int heat; string heatStr; int heatStartIndex, heatEndIndex;
            double price; string priceStr; int priceStartIndex, priceEndIndex;
            string link; string linkStr; int linkStartIndex, linkEndIndex;

            bool reportedAlready = false;

            bool regionOK = false;
            bool foundKeyword = false;
            string keyword = "";
            string keywordinfo ="";

            int hours=0, minutes=0, seconds=0;

            string tmpPage, priceInfo;

            foreach (string page in pages)
            {
                tmpPage = page;
                heatStartIndex = tmpPage.IndexOf(heatStart) + heatStart.Length;
                heatEndIndex = tmpPage.IndexOf(heatEnd);

                while ((heatStartIndex <= heatEndIndex) && (heatEndIndex - heatStartIndex<20))
                {
                    
                    heatStr = tmpPage.Substring(heatStartIndex, heatEndIndex - heatStartIndex);
                    tmpPage = tmpPage.Substring(heatStartIndex + heatStart.Length, tmpPage.Length - heatStart.Length - heatStartIndex);

                    titleStartIndex = tmpPage.IndexOf(titleStart) + titleStart.Length;
                    titleEndIndex = tmpPage.IndexOf(titleEnd);
                    titleStr = tmpPage.Substring(titleStartIndex, titleEndIndex - titleStartIndex);

                    linkStartIndex = titleStr.IndexOf(linkStart) + linkStart.Length;
                    linkEndIndex = titleStr.IndexOf(linkEnd);
                    linkStr = titleStr.Substring(linkStartIndex, linkEndIndex - linkStartIndex);
                    link = linkStr;   

                    titleStr = titleStr.Substring(linkStartIndex + linkStr.Length, titleStr.Length - linkStartIndex - linkStr.Length);

                    title = titleStr.Substring(titleStr.IndexOf("\">") + 2, titleStr.IndexOf("</a>") - titleStr.IndexOf("\">") - 2);
         

                    if (section.Equals("Deal"))
                    {
                        priceStr = titleStr.Substring(titleStr.IndexOf("xxSmall") + 13, titleStr.IndexOf("&nbsp") - titleStr.IndexOf("xxSmall") - 13).Trim();

                        if (priceStr.Split(',').Length > 2)
                            priceStr = priceStr.Remove(priceStr.IndexOf(','), 1);

                        price = double.Parse(priceStr.Replace(',', '.'), CultureInfo.InvariantCulture);
                    }
                    else price = 0;

                    title = title.Replace("&amp;", "&").Replace("&quot;", "\"").Replace("&lt;", "<").Replace("&gt;", ">").Replace("&euro;", "€").Replace("&auml;","ä").Replace("&Auml;","Ä").Replace("&uuml;","ü").Replace("&Uuml;","Ü").Replace("&ouml;","ö").Replace("&Öuml;","Ö");

                    ageStartIndex = tmpPage.IndexOf(ageStart) + ageStart.Length + 2;
                    ageEndIndex = tmpPage.IndexOf(ageEnd);
                    ageStr = tmpPage.Substring(ageStartIndex, ageEndIndex - ageStartIndex);
                    ageStr = ageStr.Substring(ageStr.IndexOf("vor")+3, ageStr.Length - ageStr.IndexOf("vor") -3).Trim();

                     if (ageStr.Contains("gestern") || ageStr.Contains("Uhr"))
                     {
                            int hmin = Convert.ToInt16(ageStr.Split(' ')[1].Split(':')[0]) * 60;
                            int mmin = Convert.ToInt16(ageStr.Split(' ')[1].Split(':')[1]);

                            minutes= 24*60+DateTime.Now.Hour*60 + DateTime.Now.Minute  - (hmin+mmin);
                     }
                     else
                     {
                        foreach (string s in ageStr.Split(' '))
                        {                       
                            if (s.Contains("h"))
                                hours = Convert.ToInt16(s.Substring(0, s.IndexOf('h')));
                            else if (s.Contains("m"))
                                minutes = Convert.ToInt16(s.Substring(0, s.IndexOf('m')));
                            else if (s.Contains("s"))
                                seconds = Convert.ToInt16(s.Substring(0, s.IndexOf('s')));
                        }
                    }

                    age = (hours * 60 * 60 + minutes * 60 + seconds)/60;

                    if (!heatStr.Trim().Equals(String.Empty) && (heatEndIndex - heatStartIndex < 20))
                    {
                        heat = Convert.ToInt16(heatStr);
                    }
                    else heat = 0;

                    //Lokalfilter
                    if (regionFilterOn && (title.Contains("lokal") || title.Contains("Lokal") || title.Contains("LOKAL")))
                    {
                        regionOK = false;
                        foreach (string s in cities.Split(','))
                        {
                            if (title.Contains(s.Trim()) && (s.Trim().Length > 2))
                            { 
                              regionOK = true;
                              break;
                            }
                        }

                    }
                    else regionOK = true;

                    //Keywordfilter
                    if (keywordFilterOn)
                    {
                        foundKeyword = false;
                        foreach (string s in keywords.Split(','))
                        {
                            if (title.Contains(s.Trim()) && (s.Trim().Length > 2))
                            {
                                foundKeyword = true;
                                keyword = s;
                                break;
                            }                               
                        }
                    }
                    else foundKeyword = false;

                    //Füge Deal hinzu
                    if ((heat >= minHeat) && (age <= maxAge) && regionOK)
                    { 
                        deals.Add(new Deal(title, age, heat, price, link));
                    }

                    //Alarm auslösen?
                    if (((heat >= minHeatAlarm) && (age <= maxAgeAlarm) && regionOK) || foundKeyword)                        
                    {
                        foreach (Deal d in reportedItems)
                            if (d.Link.Equals(link))
                                reportedAlready = true;

                        if (!reportedAlready)
                        {
                            if (section.Equals("Deal"))
                                priceInfo = "Für " + price.ToString() + " €\r\n\r\n";
                            else priceInfo = "";

                            if (foundKeyword)
                                keywordinfo = "Gemeldet da \"" + keyword + "\" im Titel gefunden\r\n\r\n";
                            else keywordinfo ="";

                            //Email?
                            if (emailAlert && alarmsOn)
                                sendMail(title + "\r\n\r\n" + priceInfo + heat.ToString() + "°C in " + age.ToString() + " Minuten!" + "\r\n\r\n" + keywordinfo + link, section);

                            reportedItems.Add(new Deal(title, age, heat, price, link));

                            //Popup-Meldung anzeigen
                            if (popupAlert && alarmsOn)
                            {
                                //mit Audio-Alarm?
                                if (audioAlert && soundOn) 
                                {
                                    if (section.Equals("Deal"))
                                    {
                                        alarm_Deal.SoundLocation = "Alarmsounds\\" + comboBox2.Text + ".wav";
                                        alarm_Deal.Load();
                                        alarm_Deal.PlayLooping();
                                    }
                                    else if (section.Equals("Gutschein"))
                                    {
                                        alarm_Coupon.SoundLocation = "Alarmsounds\\" + comboBox1.Text + ".wav";
                                        alarm_Coupon.Load();
                                        alarm_Coupon.PlayLooping();
                                    }
                                    else if (section.Equals("Freebie"))
                                    {
                                        alarm_Freebie.SoundLocation = "Alarmsounds\\" + comboBox3.Text + ".wav";
                                        alarm_Freebie.Load();
                                        alarm_Freebie.PlayLooping();
                                    }
                                }

                                DialogResult dialogResult = MessageBox.Show(heat.ToString() + "°C heisser " + section + " entdeckt!\r\n\r\n" + title + "\r\n\r\n" + priceInfo + keywordinfo + "Seite öffnen?", section + "-Alarm!", MessageBoxButtons.YesNo);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    Process.Start(link);
                                    if (section.Equals("Deal"))
                                    {
                                        alarm_Deal.Stop();
                                    }
                                    else if (section.Equals("Gutschein"))
                                    {
                                        alarm_Coupon.Stop();
                                    }
                                    else if (section.Equals("Freebie"))
                                    {
                                        alarm_Freebie.Stop();
                                    }
                                }
                                else if (dialogResult == DialogResult.No)
                                {
                                    if (section.Equals("Deal"))
                                    {
                                        alarm_Deal.Stop();
                                    }
                                    else if (section.Equals("Gutschein"))
                                    {
                                        alarm_Coupon.Stop();
                                    }
                                    else if (section.Equals("Freebie"))
                                    {
                                        alarm_Freebie.Stop();
                                    }
                                }
                            }
                        }
                    }
                    heatStartIndex = tmpPage.IndexOf(heatStart) + heatStart.Length;
                    heatEndIndex = tmpPage.IndexOf(heatEnd);                    
                }

            }

            return deals;
        }

        static public void SerializeToXML(Settings settings)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Settings));
            TextWriter textWriter = new StreamWriter(@"settings.xml");
            serializer.Serialize(textWriter, settings);
            textWriter.Close();
        }

        static Settings DeserializeFromXML(string path)
        {
            XmlSerializer deserializer = new XmlSerializer(typeof(Settings));
            TextReader textReader = new StreamReader(@path);
            Settings settings;
            settings = (Settings)deserializer.Deserialize(textReader);
            textReader.Close();

            return settings;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            getDeals();
            getCoupons();
            getFreebies();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            getDeals();
        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {
            timer1.Interval = (int)numericUpDown8.Value * 1000;           
        }

        private void sendMail(string body, string section)
        {
            try
            {
                MailMessage mail = new MailMessage();

                mail.From = new MailAddress(emailDef);
                mail.To.Add(textBox2.Text);
                mail.Subject = section.ToUpper()+"-ALARM!";
                mail.Body = body;
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                SmtpServer.Port = 587;
                
                SmtpServer.Credentials = new System.Net.NetworkCredential(emailDef, passwordDef);
                SmtpServer.EnableSsl = true;
                SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;

                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
               
            }
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
       
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            getDeals();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            alarm_Deal.SoundLocation = "Alarmsounds\\" + comboBox2.Text + ".wav";
            alarm_Deal.Load();
            alarm_Deal.Play();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            this.objectListView1.SetObjects(reportedDeals);
        }

        private void zeigeGemeldeteDealsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.objectListView1.SetObjects(reportedDeals);
        }
        
        private void alarmstopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
        
        private void zeigGemeldeteDealsToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void getDeals()
        {
            List <Deal> deals = getDeals(grabPages("Deal", (int)numericUpDown7.Value)
                                                    , "Deal"
                                                    , reportedDeals
                                                    , (int)numericUpDown1.Value
                                                    , (int)numericUpDown2.Value
                                                    , (int)numericUpDown6.Value
                                                    , (int)numericUpDown5.Value
                                                    , checkBox1.Checked
                                                    , checkBox3.Checked
                                                    , checkBox2.Checked
                                                    , checkBox13.Checked
                                                    , textBox4.Text
                                                    , checkBox15.Checked
                                                    , textBox5.Text
                                                    );

            if (!showReports)
                objectListView1.SetObjects(deals);
                     
            objectListView1.RebuildColumns();
            objectListView1.Refresh();
        }

        private void getCoupons()
        {
            objectListView2.SetObjects(getDeals(grabPages("Gutschein", (int)numericUpDown10.Value)
                                                    , "Gutschein"
                                                    , reportedCoupons
                                                    , (int)numericUpDown16.Value
                                                    , (int)numericUpDown15.Value
                                                    , (int)numericUpDown13.Value
                                                    , (int)numericUpDown12.Value
                                                    , checkBox8.Checked
                                                    , checkBox6.Checked
                                                    , checkBox7.Checked
                                                    , checkBox9.Checked
                                                    , textBox3.Text
                                                    , checkBox17.Checked
                                                    , textBox7.Text
                                                    ));
            
            objectListView2.RebuildColumns();
            objectListView2.Refresh();
        }

        private void getFreebies()
        {
            objectListView3.SetObjects(getDeals(grabPages("Freebie", (int)numericUpDown10.Value)
                                                    , "Freebie"
                                                    , reportedFreebies
                                                    , (int)numericUpDown24.Value
                                                    , (int)numericUpDown23.Value
                                                    , (int)numericUpDown21.Value
                                                    , (int)numericUpDown20.Value
                                                    , checkBox12.Checked
                                                    , checkBox10.Checked
                                                    , checkBox11.Checked
                                                    , checkBox5.Checked
                                                    , textBox1.Text
                                                    , checkBox16.Checked
                                                    , textBox6.Text
                                                    ));

            objectListView3.RebuildColumns();
            objectListView3.Refresh();
        }

        private void aktualisierenf5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            getDeals();
            getCoupons();
            getFreebies();
        }

        private void objectListView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                getDeals();
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
           
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
           
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
          
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
          
        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {
         
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
           
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
           
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
          
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                checkBox2.Enabled = true;
                comboBox2.Enabled = true;               
            }
            else
            {
                checkBox2.Enabled = false;
                comboBox2.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
          
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void sichereEinstellungenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveSettings();
        }

        private void objectListView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void objectListView4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            getCoupons();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
          /*  if (button2.Text.Equals("Gemeldet"))
            {
                button2.Text = "Aktuell";
                this.objectListView2.SetObjects(reportedCoupons);
            }
            else if (button2.Text.Equals("Aktuell"))
            {
                button2.Text = "Gemeldet";
                getCoupons();
            }*/
        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            if (button1.Text.Equals("Gemeldet"))
            {
                button1.Text = "Aktuell";
                this.objectListView1.SetObjects(reportedDeals);
            }
            else if (button1.Text.Equals("Aktuell"))
            {
                button1.Text = "Gemeldet";
                getDeals();
            }
        }

        private void numericUpDown9_ValueChanged(object sender, EventArgs e)
        {
            timer2.Interval = (int)numericUpDown9.Value * 1000;       
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            getFreebies();
        }

        private void numericUpDown17_ValueChanged(object sender, EventArgs e)
        {
            timer3.Interval = (int)numericUpDown17.Value * 1000;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            getDeals();
            getCoupons();
            getFreebies();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            saveSettings();
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
          /*  if (button3.Text.Equals("Gemeldet"))
            {
                button3.Text = "Aktuell";
                this.objectListView3.SetObjects(reportedFreebies);
            }
            else if (button3.Text.Equals("Aktuell"))
            {
                button3.Text = "Gemeldet";
                getFreebies();
            }*/
        }

        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox4_KeyDown(object sender, KeyEventArgs e)
        {
            
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            alarm_Deal.SoundLocation = "Alarmsounds\\" + comboBox2.Text + ".wav";
            alarm_Deal.Load();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            alarm_Coupon.SoundLocation = "Alarmsounds\\" + comboBox1.Text + ".wav";
            alarm_Coupon.Load();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            alarm_Freebie.SoundLocation = "Alarmsounds\\" + comboBox3.Text + ".wav";
            alarm_Freebie.Load();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            alarm_Deal.Play();
        }

        private void button7_Click(object sender, EventArgs e)
        {

            alarm_Coupon.Play();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            alarm_Freebie.Play();
        }

        private void checkBox3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            { 
                checkBox2.Enabled = true;
                comboBox2.Enabled = true;
                button8.Enabled = true;            
            }
            else
            {
                checkBox2.Enabled = false;
                comboBox2.Enabled = false;
                button8.Enabled = false;   
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                checkBox7.Enabled = true;
                comboBox1.Enabled = true;
                button7.Enabled = true;
            }
            else
            {
                checkBox7.Enabled = false;
                comboBox1.Enabled = false;
                button7.Enabled = false;
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox10.Checked)
            {
                checkBox11.Enabled = true;
                comboBox3.Enabled = true;
                button6.Enabled = true;
            }
            else
            {
                checkBox11.Enabled = false;
                comboBox3.Enabled = false;
                button6.Enabled = false;
            }
        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox11.Checked)
            {                
                comboBox3.Enabled = true;
                button6.Enabled = true;
            }
            else
            {
                comboBox3.Enabled = false;
                button6.Enabled = false;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                comboBox1.Enabled = true;
                button7.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = false;
                button7.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                comboBox2.Enabled = true;
                button8.Enabled = true;
            }
            else
            {
                comboBox2.Enabled = false;
                button8.Enabled = false;
            }
        }

        private void objectListView1_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                getDeals();
            }
        }

        private void objectListView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                getCoupons();
            }
        }

        private void objectListView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                getFreebies();
            }
        }

        private void CreateShortcut()
        {
            object shDesktop = (object)"Desktop";
            WshShell shell = new WshShell();
            string shortcutAddress = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\hukdalrm.lnk";
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutAddress);
            shortcut.Description = "hukdalrm - hukd alarm tool";
            shortcut.WorkingDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            
            shortcut.TargetPath = System.Reflection.Assembly.GetEntryAssembly().Location;
            shortcut.Save();
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox14.Checked)
            {
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\hukdalrm.lnk");
                checkBox4.Enabled = false;
            }
            else
            { 
                CreateShortcut();
                checkBox4.Enabled = true;
            }
        }

        private void button1_Click_4(object sender, EventArgs e)
        {

           if (showReports)
           {
               button1.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.update_recommended));
               pictureBox1.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.upgrade_misc));
               showReports = false;
               toolTip1.SetToolTip(button1, "Zeige zurückliegende Meldungen an");
               getDeals();
               getCoupons();
               getFreebies();                             
           }
           else
           {
               button1.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.upgrade_misc));
               pictureBox1.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.update_recommended));
               toolTip1.SetToolTip(button1, "Zeige aktuelle Daten an");
               objectListView1.SetObjects(reportedDeals);
               objectListView2.SetObjects(reportedCoupons);
               objectListView3.SetObjects(reportedFreebies);
               showReports = true;
           }

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            TextMatchFilter filter = TextMatchFilter.Contains(objectListView1, textBox8.Text);

            if (tabControl1.SelectedTab.Text.Equals("Deals"))
            {
                objectListView1.ModelFilter = filter;
                objectListView1.DefaultRenderer = new HighlightTextRenderer(filter);
                objectListView1.RebuildColumns();
            }
            else if (tabControl1.SelectedTab.Text.Equals("Gutscheine"))
            {
                objectListView2.ModelFilter = filter;
                objectListView2.DefaultRenderer = new HighlightTextRenderer(filter);
                objectListView2.RebuildColumns();
            }
            else if (tabControl1.SelectedTab.Text.Equals("Freebies"))
            {
                objectListView3.ModelFilter = filter;
                objectListView3.DefaultRenderer = new HighlightTextRenderer(filter);
                objectListView3.RebuildColumns();
            }
        }

        private void textBox8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Escape)
                textBox8.Text = "";
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void textBox8_Enter(object sender, EventArgs e)
        {
            textBox8.Text = "";
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            textBox8.Text = "Finde...";
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            textBox8.Text = "Finde...";
        }

        private void button3_Click_2(object sender, EventArgs e)
        {
            if (soundOn)
            {
                button3.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.kmixdocked_error_2));

                toolTip1.SetToolTip(button3, "Aktiviere alle Audio-Alarme");

                soundOn = false;
            }
            else
            {
                button3.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.kmixdocked_2));

                toolTip1.SetToolTip(button3, "Deaktiviere alle Audio-Alarme");

                soundOn = true;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (alarmsOn)
            {
                button9.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.kalarm_disabled));

                alarmsOn = false;

                toolTip1.SetToolTip(button9, "Aktiviere alle Alarme");               

                button3.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.kmixdocked_error_2));
                button3.Enabled = false;

                toolTip1.SetToolTip(button3, "Aktiviere alle Audio-Alarme");

                soundOn = false;
            }
            else
            {
                button9.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.kalarm));

                alarmsOn = true;

                toolTip1.SetToolTip(button9, "Deaktiviere alle Alarme (dh. kein Popup, kein Sound, keine Mail)");               

                button3.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.kmixdocked_2));
                button3.Enabled = true;

                toolTip1.SetToolTip(button3, "Deaktiviere alle Audio-Alarme");

                soundOn = true;
            }
        }

        private void numericUpDown8_ValueChanged_1(object sender, EventArgs e)
        {
            timer1.Interval = (int)numericUpDown8.Value * 1000;
        }
    }
}
