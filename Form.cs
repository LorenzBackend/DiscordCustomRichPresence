// Copyright Lorenz Fresh
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DiscordRPC;
using System.Windows.Automation;
using System.Diagnostics;

namespace Discord_Rich_Presence {
    public partial class Form1 : MetroFramework.Forms.MetroForm {
        public Form1() {
            InitializeComponent();
        }

       // Copyright Lorenz Fresh
        string LastAppID;
        string LastYouTube;
        string LastSoundCloud;
        string LastTwitch;
        string LastApp;
        string AdobeApp;
        string officeApp;
        string browser = "";
        private DiscordRpcClient client;
        private void Form1_Load(object sender, EventArgs e) {
            LoadSettings();
            updateBtn.Hide();
        }

        private void metroButton1_Click(object sender, EventArgs e) {

            if (ServiceSettings.BrowserMode || ServiceSettings.AppMode) {
                DialogResult result = MessageBox.Show("there is another service running do you want to stop it to start the browser service?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes) {
                    appService.Stop();
                    browserService.Stop();
                    client.Dispose();
                    ServiceSettings.BrowserMode = false;
                    ServiceSettings.AppMode = false;
                    activebrowser.Hide();
                    activebw.Hide();
                } else {
                    return;
                }
            }

            ServiceSettings.customMode = true;
            AppState.Text = "Service is OFF";
            BrowserState.Text = "Service is OFF";
            if (metroButton1.Text == "START SERVICE") {
                CheckSettings();
                metroButton1.Text = "STOP SERVICE";
                updateBtn.Show();
                clientID.Enabled = false;
            } else {
                if (client != null) {
                    client.Dispose();
                }

                metroButton1.Text = "START SERVICE";
                updateBtn.Hide();
                clientID.Enabled = true;
            }
        }


        private void CheckSettings() {

            appService.Stop();

            if (clientID.Text.Length < 4) {
                MessageBox.Show("No Client ID!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                metroButton1.Text = "STOP SERVICE";
                return;
            }

            if (metroToggle1.Checked) {
                if (btnText.Text.Length < 1 || btnURL.Text.Length < 1) {
                    MessageBox.Show("Button Text or Link Not configured!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    metroButton1.Text = "STOP SERVICE";
                    return;
                }
            }

            InitClientID(clientID.Text);

            if (message.Text.Length > 0 && message2.Text.Length > 0 && metroToggle1.Checked == false) {
                client.SetPresence(new DiscordRPC.RichPresence() {
                    Details = message.Text,
                    State = message2.Text,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = appIcon.Text,
                    }
                    
                });
            }

            if (message.Text.Length > 0 && message2.Text.Length > 0 && metroToggle1.Checked) {

                DiscordRPC.Button btn = new DiscordRPC.Button();
                btn.Label = btnText.Text;
                btn.Url = btnURL.Text;

                client.SetPresence(new DiscordRPC.RichPresence() {
                    Details = message.Text,
                    State = message2.Text,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = appIcon.Text,
                    },
                    Buttons = new DiscordRPC.Button[] {btn}

                });
            }


            if (message.Text.Length > 0 && message2.Text.Length < 1 && metroToggle1.Checked) {
                DiscordRPC.Button btn = new DiscordRPC.Button();
                btn.Label = btnText.Text;
                btn.Url = btnURL.Text;
                client.SetPresence(new DiscordRPC.RichPresence() {
                    Details = message.Text,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = appIcon.Text,
                    },
                    Buttons = new DiscordRPC.Button[] { btn }

                });
            }



            if (message.Text.Length > 0 && message2.Text.Length < 1 && metroToggle1.Checked == false) {
                client.SetPresence(new DiscordRPC.RichPresence() {
                    Details = message.Text,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = appIcon.Text,
                    }

                });
            }


            Properties.Settings.Default.clientID = clientID.Text;
            Properties.Settings.Default.appIcon = appIcon.Text;
            Properties.Settings.Default.message = message.Text;
            Properties.Settings.Default.message2 = message2.Text;
            Properties.Settings.Default.btnText = btnText.Text;
            Properties.Settings.Default.btnUrl = btnURL.Text;
            Properties.Settings.Default.btnToogle = metroToggle1.Checked;
            Properties.Settings.Default.Save();
        }

        private void InitClientID(string id) {
            client = new DiscordRpcClient(id);
            client.Initialize();
        }

        private void metroToggle1_CheckedChanged(object sender, EventArgs e) {
            if (metroToggle1.Checked) {
                btnText.Enabled = true;
                btnURL.Enabled = true;
            } else {
                btnText.Enabled = false;
                btnURL.Enabled = false;
            }
        }

        private void LoadSettings() {
            clientID.Text = Properties.Settings.Default.clientID;
            appIcon.Text = Properties.Settings.Default.appIcon;
            message.Text = Properties.Settings.Default.message;
            message2.Text = Properties.Settings.Default.message2;

            if (Properties.Settings.Default.btnToogle) {
                btnText.Text = Properties.Settings.Default.btnText;
                btnURL.Text = Properties.Settings.Default.btnUrl;
            } else {
                btnText.Enabled = false;
                btnURL.Enabled = false;
            }
        }

        private void appService_Tick(object sender, EventArgs e) {

            //Checking for Closed Apps
            CheckClosedApps();

            //Checking for New Apps
            CheckNewRunningApp();
        }

        private void metroButton2_Click(object sender, EventArgs e) {
            if (metroButton2.Text == "STOP SERVICE") {
                activebw.Hide();
                activebrowser.Hide();
                BrowserState.Text = "Service is OFF";
                browserService.Stop();
                if (client != null) {
                    client.Dispose();
                }
                metroButton2.Text = "START SERVICE";
                return;
            }
            if (ServiceSettings.AppMode || ServiceSettings.customMode) {
                DialogResult result = MessageBox.Show("there is another service running do you want to stop it to start the browser service?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes) {
                    browserService.Stop();
                    if (client != null) {
                        client.Dispose();
                    }
                    updateBtn.Hide();
                    clientID.Enabled = true;
                    ServiceSettings.AppMode = false;
                    ServiceSettings.customMode = false;
                    ServiceSettings.BrowserMode = true;
                    browserService.Start();
                    AppState.Text = "Service is OFF";
                    BrowserState.Text = "Service is Running";
                    metroButton2.Text = "STOP SERVICE";
                    activebw.Show();
                    activebrowser.Show();
                    return;
                }
          }
            browserService.Start();
            activebw.Show();
            activebrowser.Show();
            ServiceSettings.BrowserMode = true;
            AppState.Text = "Service is OFF";
            BrowserState.Text = "Service is Running";
            metroButton2.Text = "STOP SERVICE";
        }

        private void metroButton4_Click(object sender, EventArgs e) {
            this.Hide();
            notifyIcon1.ShowBalloonTip(5);
        }

        private void metroButton3_Click(object sender, EventArgs e) {

            if (metroButton3.Text == "STOP SERVICE") {
                ServiceSettings.AppMode = false;
                activebrowser.Hide();
                activebw.Hide();
                BrowserState.Text = "Service is OFF";
                appService.Stop();
                if (client != null) {
                    client.Dispose();
                }
                metroButton3.Text = "START SERVICE";
                return;
            }

            if (ServiceSettings.customMode || ServiceSettings.BrowserMode) {
                    DialogResult result = MessageBox.Show("there is another service running do you want to stop it to start the App service?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes) {
                        browserService.Stop();
                    if (client != null) {
                        client.Dispose();
                    }
                    updateBtn.Hide();
                        metroButton1.Text = "START SERVICE";
                    clientID.Enabled = true;
                    ServiceSettings.AppMode = true;
                        ServiceSettings.customMode = false;
                        ServiceSettings.BrowserMode = false;
                    activebrowser.Hide();
                    activebw.Hide();
                    appService.Start();
                        BrowserState.Text = "Service is OFF";
                        AppState.Text = "Service is ON";
                        metroButton3.Text = "STOP SERVICE";
                    return;
                    }
                }

            metroButton3.Text = "STOP SERVICE";
            appService.Start();
            BrowserState.Text = "Service is OFF";
            AppState.Text = "Service is Running";
            ServiceSettings.AppMode = true;
        }

        private void metroToggle2_CheckedChanged(object sender, EventArgs e) {
            Properties.Settings.Default.AdobeApps = metroToggle2.Checked;
            Properties.Settings.Default.Save();
        }

        private void metroToggle3_CheckedChanged(object sender, EventArgs e) {
            Properties.Settings.Default.OfficeApps = metroToggle3.Checked;
            Properties.Settings.Default.Save();
        }

        private void metroToggle5_CheckedChanged(object sender, EventArgs e) {
            Properties.Settings.Default.unity = metroToggle5.Checked;
            Properties.Settings.Default.Save();
        }

        private void metroToggle4_CheckedChanged(object sender, EventArgs e) {
            Properties.Settings.Default.NetflixApp = metroToggle4.Checked;
            Properties.Settings.Default.Save();
        }

        private void metroToggle6_CheckedChanged(object sender, EventArgs e) {
            Properties.Settings.Default.UE = metroToggle6.Checked;
            Properties.Settings.Default.Save();
        }

        private void notifyIcon1_Click(object sender, EventArgs e) {
            this.Show();
        }


        // Check Apps for Running and return result
        private bool CheckUnity() {
            Process[] unity = Process.GetProcessesByName("Unity");
            if (unity.Length > 0) {
                return true;
            }
            return false;
        }

        private bool CheckAdobe() {
            Process[] ps = Process.GetProcessesByName("Photoshop");
            Process[] pr = Process.GetProcessesByName("Adobe Premiere Pro");
            Process[] ae = Process.GetProcessesByName("AfterFX");
            Process[] xd = Process.GetProcessesByName("XD");

            if (ps.Length > 0) {
                AdobeApp = "ps";
                return true;
            }
            if (pr.Length > 0) {
                AdobeApp = "pr";
                return true;
            }
            if (ae.Length > 0) {
                AdobeApp = "ae";
                return true;
            }
            if (xd.Length > 0) {
                AdobeApp = "xd";
                return true;
            }

            return false;
        }

        private bool CheckOfficeApps() {
            Process[] Word = Process.GetProcessesByName("WINWORD");
            Process[] PowerPoint = Process.GetProcessesByName("POWERPNT");
            Process[] Excel = Process.GetProcessesByName("EXCEL");
            if (Word.Length > 0) {
                officeApp = "w";
                return true;
            }

            if (PowerPoint.Length > 0) {
                officeApp = "p";
                return true;
            }

            if (Excel.Length > 0) {
                officeApp = "e";
                return true;
            }
            return false;
        }

        private bool CheckNetflix() {
            Process[] netflix = Process.GetProcessesByName("WWAHost");
            if (netflix.Length > 0) {
                return true;
            }
            return false;
        }

        private bool CheckUE() {
            Process[] UE = Process.GetProcessesByName("UE4Editor");
            if (UE.Length > 0) {
                return true;
            }
            return false;
        }


        // Check for Active Browsers
        private bool isBrowserActive() {
            Process[] procsChrome = Process.GetProcessesByName("chrome");
            Process[] firefox = Process.GetProcessesByName("firefox");
            Process[] opera = Process.GetProcessesByName("opera");
            Process[] edge = Process.GetProcessesByName("msedge");
            Process[] explorer = Process.GetProcessesByName("iexplore");


            if (procsChrome.Length > 0) {
                browser = "chrome";
                return true;
            }

            if (firefox.Length > 0) {
                browser = "firefox";
                return true;
            }

            if (opera.Length > 0) {
                browser = "opera";
                return true;
            }

            return false;
        }

        private void CheckClosedApps() {
            if (Properties.Settings.Default.unity && CheckUnity() == false && LastAppID == AppIDTypes.Unity) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.UE && CheckUE() == false && LastAppID == AppIDTypes.UnrealEngine) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.NetflixApp && CheckOfficeApps() == false && LastAppID == AppIDTypes.NetflixID) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.AdobeApps && CheckAdobe() == false && LastAppID == AppIDTypes.PremierePro) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.AdobeApps && CheckAdobe() == false && LastAppID == AppIDTypes.xD) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.AdobeApps && CheckAdobe() == false && LastAppID == AppIDTypes.AfterEffects) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.AdobeApps && CheckAdobe() == false && LastAppID == AppIDTypes.Photoshop) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.OfficeApps && CheckOfficeApps() == false && LastAppID == AppIDTypes.Word) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.OfficeApps && CheckOfficeApps() == false && LastAppID == AppIDTypes.PowerPoint) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }

            if (Properties.Settings.Default.OfficeApps && CheckOfficeApps() == false && LastAppID == AppIDTypes.Word) {
                if (client != null) {
                    client.Dispose();
                }
                LastAppID = "";
            }
        }

        private void CheckNewRunningApp() {

            //If Netflix Running
            if (Properties.Settings.Default.NetflixApp && CheckNetflix() && LastAppID != AppIDTypes.NetflixID) {
                if (LastAppID != "") {
                    if (client != null) {
                        client.Dispose();
                    }
                    LastAppID = AppIDTypes.NetflixID;
                }
                InitClientID(AppIDTypes.NetflixID);
                client.SetPresence(new RichPresence() {
                    Details = "Watching Netflix",
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = "app_icon"
                    }
                });
            }

                //If Unity Running
                if (Properties.Settings.Default.unity && CheckUnity() && LastAppID != AppIDTypes.Unity) {
                    if (LastAppID != "") {
                    if (client != null) {
                        client.Dispose();
                    }
                    LastAppID = AppIDTypes.Unity;
                    }
                    InitClientID(AppIDTypes.Unity);
                    client.SetPresence(new RichPresence() {
                        Details = "is developing a game",
                        Timestamps = Timestamps.Now,
                        Assets = new Assets() {
                            LargeImageKey = "app_icon"
                        }
                    });
                }


            //If Unreal Running
            if (Properties.Settings.Default.UE && CheckUE() && LastAppID != AppIDTypes.UnrealEngine) {
                if (LastAppID != "") {
                    if (client != null) {
                        client.Dispose();
                    } 
                    LastAppID = AppIDTypes.UnrealEngine;
                }
                InitClientID(AppIDTypes.UnrealEngine);
                client.SetPresence(new RichPresence() {
                    Details = "is developing a game",
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = "app_icon"
                    }
                });
            }


            //If a Office app  Running
            if (Properties.Settings.Default.AdobeApps && CheckAdobe()) {
                switch (officeApp) {
                    case "w":
                        if (LastAppID != AppIDTypes.Word) {
                            LastAppID = AppIDTypes.Word;
                            if (client != null) {
                                client.Dispose();
                            }

                            InitClientID(AppIDTypes.Word);
                            client.SetPresence(new RichPresence() {
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon"
                                }
                            });
                        }
                        break;
                    case "p":
                        if (LastAppID != AppIDTypes.PowerPoint) {
                            LastAppID = AppIDTypes.PowerPoint;
                            if (client != null) {
                                client.Dispose();
                            }

                            InitClientID(AppIDTypes.PowerPoint);
                            client.SetPresence(new RichPresence() {
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon"
                                }
                            });
                        }
                        break;
                    case "e":
                        if (LastAppID != AppIDTypes.Excel) {
                                        LastAppID = AppIDTypes.Word;
                            if (client != null) {
                                client.Dispose();
                            }

                            InitClientID(AppIDTypes.Excel);
                            client.SetPresence(new RichPresence() {
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon"
                                }
                            });
                        }
                        break;
                }

            }


            //If a Adobe app  Running
            if (Properties.Settings.Default.AdobeApps && CheckAdobe()) {
               
                switch (AdobeApp) {
                    case "ps":
                        if (LastAppID != AppIDTypes.Photoshop) {
                            LastAppID = AppIDTypes.Photoshop;
                            if (client != null) {
                                client.Dispose();
                            }
                            InitClientID(AppIDTypes.Photoshop);
                            client.SetPresence(new RichPresence() {
                                Details = "Edited a beautiful picture",
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon",
                                    LargeImageText = "DiscordPresence by LorenzFresh",
                                }
                            });

                        }
                        break;

                    case "ae":
                        if (LastAppID != AppIDTypes.AfterEffects) {
                            LastAppID = AppIDTypes.AfterEffects;
                            if (client != null) {
                                client.Dispose();
                            }
                            InitClientID(AppIDTypes.AfterEffects);
                            client.SetPresence(new RichPresence() {
                                Details = "Edits a perfect video",
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon"
                                }
                            });

                        }
                        break;

                    case "pr":
                        if (LastAppID != AppIDTypes.PremierePro) {
                            LastAppID = AppIDTypes.PremierePro;
                            if (client != null) {
                                client.Dispose();
                            }
                            InitClientID(AppIDTypes.PremierePro);
                            client.SetPresence(new RichPresence() {
                                Details = "Edits a perfect video",
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon"
                                }
                            });

                        }
                        break;

                    case "xd":
                        if (LastAppID != AppIDTypes.xD) {
                            LastAppID = AppIDTypes.xD;
                            if (client != null) {
                                client.Dispose();
                            }
                            InitClientID(AppIDTypes.xD);
                            client.SetPresence(new RichPresence() {
                                Details = "Designs something",
                                Timestamps = Timestamps.Now,
                                Assets = new Assets() {
                                    LargeImageKey = "app_icon"
                                }
                            });

                        }
                        break;
                }
         
            }

        }

        private void browserService_Tick(object sender, EventArgs e) {
            if (isBrowserActive()) {
                activebrowser.Show();
                activebw.Show();
                 activebrowser.Text = browser;

  
                Process[] procsChrome = Process.GetProcessesByName(browser);
                foreach (Process chrome in procsChrome) {
                    if (chrome.MainWindowHandle == IntPtr.Zero) {
                        continue;
                    }
                    AutomationElement sourceElement = AutomationElement.FromHandle(chrome.MainWindowHandle);
                    AutomationElement elmTitleBar = sourceElement.FindFirst(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TitleBar));
                    if (elmTitleBar != null) {
                        UpdateDiscordBrowserTab(elmTitleBar.Current.Name.Replace(browser, ""));
                    }

                }
            } else {
                activebrowser.Hide();
                activebw.Hide();
                if (client != null) {
                    client.Dispose();
                }
            }
        }

        // Update the DiscordPresence
        private void UpdateDiscordBrowserTab(string tab) {
            if (tab.Contains("YouTube") && LastYouTube != tab) {
                LastYouTube = tab;
                if (client != null) {
                    client.Dispose();
                }
                InitClientID(AppIDTypes.YouTubeID);
                string CheckLength = tab.Replace("- YouTube", "");
                client.SetPresence(new RichPresence() {
                    Details = "just looks",
                    State = CheckLength,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = "app_icon"
                    }
                });
                return;
            }

            if (tab.Contains("Twitch") && LastTwitch != tab) {
                LastTwitch = tab;
                if (client != null) {
                    client.Dispose();
                }
                InitClientID(AppIDTypes.TwitchID);
                string CheckLength = tab.Replace("- Twitch", "");
                client.SetPresence(new RichPresence() {
                    Details = "just looks",
                    State = CheckLength,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = "app_icon"
                    }
                });
                return;
            }

            if (tab.Contains("▶") && LastSoundCloud != tab) {
                LastSoundCloud = tab;
                if (client != null) {
                    client.Dispose();
                }
                InitClientID(AppIDTypes.SoundCloud);
                client.SetPresence(new RichPresence() {
                    Details = "is listening",
                    State = tab,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = "app_icon"
                    }
                });
                return;
            }

            if (tab.Contains("Netflix") && LastApp != AppIDTypes.NetflixID) {
                LastApp = AppIDTypes.NetflixID;

                if (client != null) {
                    client.Dispose();
                }
                InitClientID(AppIDTypes.NetflixID);
                client.SetPresence(new RichPresence() {
                    Details = "is listening",
                    State = tab,
                    Timestamps = Timestamps.Now,
                    Assets = new Assets() {
                        LargeImageKey = "app_icon"
                    }
                });
                return;
            }

        }

        private void UpdateDisocrdCustom(string appIcon, string message, string message2) {
            client.UpdateLargeAsset(appIcon);
            client.UpdateDetails(message);
            client.UpdateState(message2);
        }

        private void updateBtn_Click(object sender, EventArgs e) {
            UpdateDisocrdCustom(appIcon.Text, message.Text, message2.Text);
        }

        private void label22_Click(object sender, EventArgs e) {
            Process.Start("https://discord.com/developers/applications");
        }
    }
}
