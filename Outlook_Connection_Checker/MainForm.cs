using Microsoft.Office.Interop.Outlook;
using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application; // Alias for Outlook Application

namespace Outlook_Connection_Checker
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            this.Visible = false;
            this.Load += MainForm_Load;
        }

       

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);

        private void MainForm_Load(object sender, EventArgs e)
        {
            Console.WriteLine("Checking Outlook connection...");
            CheckOutlookConnection();
            

        }

        private void CheckOutlookConnection()
        {
            if (IsOutlookRunning())
            {
                OutlookApp outlookApp = null; // alias used
                try
                {
                    Console.WriteLine("Outlook is running. Attempting to retrieve Outlook application object...");
                    outlookApp = (OutlookApp)Marshal.GetActiveObject("Outlook.Application");

                    if (outlookApp.Session.Offline)
                    {
                       // Debug.WriteLine("Outlook is working offline.");
                
                        ShowDisconnectedWarning("Outlook is working offline.\n\nYou will not be receiving new emails\n\nOutbound mails will stay in your outbox and not send.\n\nPlease check!");
                    }
                    else if (IsOfflineFromCachedExchange(outlookApp))
                    {
                        // Debug.WriteLine("Outlook is working offline.");

                        ShowDisconnectedWarning("Outlook is working offline.\n\nYou will not be receiving new emails\n\nOutbound mails will stay in your outbox and not send.\n\nPlease check!");

                    }
                    else if (IsDisconnectedFromExchange(outlookApp))
                    {
                        // Debug.WriteLine("Outlook is not connected to Microsoft 365.");
                        ShowDisconnectedWarning("Outlook is not connected to Microsoft 365 Servers. \n\nYou may not be receiving new emails\n\nSent emails may stay in your outbox and not send.\n\nPlease check!");

                    }
                    else if (IsDisconnectedFromCachedExchange(outlookApp))
                    {
                        // Debug.WriteLine("Outlook is not connected to Microsoft 365.");
                        ShowDisconnectedWarning("Outlook is not connected to Microsoft 365 Servers. \n\nYou may not be receiving new emails\n\nSent emails may stay in your outbox and not send.\n\nPlease check!");

                    }
                

                    else
                    {
                        //  Debug.WriteLine("Outlook is connected to Microsoft 365.");
                        //we don't need to do anything if outlook is connected fine
                        IntPtr outlookHandle = FindWindow("rctrl_renwnd32", null);
                       // Only if outlook is a foreground process.
                        if (outlookHandle != IntPtr.Zero)
                        {
                        // MessageBox.Show("Outlook is connected to Microsoft 365.", "Outlook Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        // Check if there are 5 or more items in the Outbox
                        int outboxItemCount = GetOutboxItemCount(outlookApp);
                        if (outboxItemCount >= 5)
                            
                        {
                            //Sleep for 4 minutes
                                Thread.Sleep(240000);


                            {
                                int secondCheckItemCount = GetOutboxItemCount(outlookApp);
                                if (secondCheckItemCount >= 5)
                                {
                                    outlookHandle = FindWindow("rctrl_renwnd32", null);
                                    // Only if outlook is a foreground process.
                                    if (outlookHandle != IntPtr.Zero)
                                    {
                                        // Show warning about too many items in the Outbox after 4 minutes
                                        ShowDisconnectedWarning($"There are  {secondCheckItemCount} items stuck in your Outbox after 4+ minutes. \n\nOutlook may be having problems sending messages.");
                                    }
                                    }
                               
                            };
                           
                        }
                    }
                }
                catch (COMException ex)
                {
                   // Debug.WriteLine($"Failed to retrieve Outlook information. Error: {ex.Message}");
                   // ShowDisconnectedWarning($"Failed to retrieve Outlook information. Error: {ex.Message} \n\nPlease confirm your outlook is connected and working!\n\nContact your system administrator if you arent sure");
                }
                finally
                {
                    if (outlookApp != null)
                        Marshal.ReleaseComObject(outlookApp);
                }
            }
            else
            {
               // Debug.WriteLine("Outlook is not running.");
               //We dont need to check if outlook is not running
               // ShowDisconnectedWarning("Outlook is not running.");


            }

            // Close the application after displaying messages
            System.Environment.Exit(1);
            System.Windows.Forms.Application.Exit();
        }


        private bool IsOutlookRunning()
        {
            Process[] processes = Process.GetProcessesByName("OUTLOOK");
            bool isRunning = processes.Length > 0;
            Console.WriteLine($"Outlook is {(isRunning ? "running" : "not running")}.");
            return isRunning;
        }

   
        
        private bool IsDisconnectedFromExchange(OutlookApp outlookApp)
        {
            Accounts accounts = outlookApp.Session.Accounts;
            foreach (Account account in accounts)
            {
                if (account.ExchangeConnectionMode == OlExchangeConnectionMode.olDisconnected)
                {
                    return true;
                }
            }
            return false;
        }
        private bool IsDisconnectedFromCachedExchange(OutlookApp outlookApp)
        {
            Accounts accounts = outlookApp.Session.Accounts;
            foreach (Account account in accounts)
            {
                if (account.ExchangeConnectionMode == OlExchangeConnectionMode.olCachedDisconnected)
                {
                    return true;
                }
            }
            return false;
        }

        private bool IsOfflineFromCachedExchange(OutlookApp outlookApp)
        {
            Accounts accounts = outlookApp.Session.Accounts;
            foreach (Account account in accounts)
            {
                if (account.ExchangeConnectionMode == OlExchangeConnectionMode.olCachedOffline)
                {
                    return true;
                }
            }
            return false;
        }
        private int GetOutboxItemCount(OutlookApp outlookApp)
        {
            MAPIFolder outbox = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderOutbox);
            return outbox.Items.Count;
        }

        private void ShowDisconnectedWarning(string message)
        {
            //Debug.WriteLine(message);
            new ToastContentBuilder()

.AddText("Outlook Problem Detected")
.AddText(message)
.Show();
            MessageBox.Show(message, "Outlook Problem Detected", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            System.Environment.Exit(1);
        
        }
    }
}
