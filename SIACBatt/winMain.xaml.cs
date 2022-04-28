using SIACBatt.Languages;
using SPORTident;
using SPORTident.Common;
using SPORTident.Communication;
using SPORTident.Communication.UsbDevice;
using System;
using System.Collections.Generic;
using System.IO;
using System.Timers;
using System.Windows;
using System.Windows.Media;

namespace SIACBatt
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class winMain : Window
    {
        private readonly Reader _SIReader;
        private Dictionary<int, DeviceInfo> siDeviceInfoList;
        private bool connected;


        System.Timers.Timer ClearTimer = new System.Timers.Timer();


        public winMain()
        {
            InitializeComponent();

            _SIReader = new Reader
            {
                WriteBackupFile = false,
                BackupFileName = Path.Combine(Environment.CurrentDirectory, $@"backup\{DateTime.Now:yyyy-MM-dd}_stamps.bak"),
                WriteLogFile = false
            };

            // Bind event functions
            _SIReader.SiCardReadCompleted += siReader_CardRead;
            _SIReader.InputDeviceChanged += siReader_InputDeviceChanged;
            _SIReader.InputDeviceStateChanged += siReader_InputDeviceStateChanged;
            _SIReader.LogEvent += siReader_LogEvent;
            _SIReader.ErrorOccured += siReader_ErrorOccured;
            _SIReader.DeviceConfigurationRead += siReader_DeviceConfigurationRead;

            ListPorts();
        }

        /// <summary>
        /// Clear all data fields
        /// </summary>
        private void ClearDataAfterTime(object source, ElapsedEventArgs e)
        {
            ClearTimer.Enabled = false;
            Application.Current.Dispatcher.Invoke(new Action(() => {
                txbNumber.Text = "";

                txbVoltage.Text = "";
                lbVoltageStatus.Content = "";
                recVoltageBackground.Fill = Brushes.Transparent;
                txbVoltageDescription.Text = "";

                txbDate.Text = "";
                lbDateStatus.Content = "";
                recDateBackground.Fill = Brushes.Transparent;
                txbDateDescription.Text = "";
            }));
        }

        #region Button events
        /// <summary>
        /// Start/end connection with device
        /// </summary>
        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {
            if (!connected) SIConnect();
            else SIDisconnect();
        }

        /// <summary>
        /// Open window in fullscreen
        /// </summary>
        private void btnFullscreen_Click(object sender, RoutedEventArgs e)
        {
            this.WindowStyle = WindowStyle.None;
            this.WindowState = WindowState.Maximized;
            grpButtons.Visibility = Visibility.Collapsed;
            grpFormInput.Visibility = Visibility.Collapsed;
            btnFullscreenExit.Visibility = Visibility.Visible;
            grpFormOutput.Header = Resource.Title;
            grpFormOutput.FontSize = 50;
        }

        /// <summary>
        /// Close fullscreen mode
        /// </summary>
        private void btnFullscreenExit_Click(object sender, RoutedEventArgs e)
        {
            this.WindowStyle = WindowStyle.ToolWindow;
            this.WindowState = WindowState.Normal;
            grpButtons.Visibility = Visibility.Visible;
            grpFormInput.Visibility = Visibility.Visible;
            btnFullscreenExit.Visibility = Visibility.Collapsed;
            grpFormOutput.Header = Resource.OutputHeader;
            grpFormOutput.FontSize = 12;
        }

        /// <summary>
        /// Exit application
        /// </summary>
        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        /// <summary>
        /// Refresh device list
        /// </summary>
        private void btnCOMPortRefresh_Click(object sender, RoutedEventArgs e)
        {
            ListPorts();
        }
        #endregion

        #region Device connection
        /// <summary>
        /// List all available ports
        /// </summary>
        private void ListPorts()
        {
            siDeviceInfoList = new Dictionary<int, DeviceInfo>();
            var devList = DeviceInfo.GetAvailableDeviceList(true, (int)DeviceType.Serial | (int)DeviceType.UsbHid);
            var selectList = new Dictionary<int, string>();
            var n = 0;

            cmbCOMPorts.ItemsSource = selectList;

            foreach (var item in devList)
            {
                selectList.Add(n, DeviceInfo.GetPrettyDeviceName(item));
                siDeviceInfoList.Add(n, item);
                n++;
            }

            if (cmbCOMPorts.Items.Count > 0) cmbCOMPorts.SelectedIndex = 0;
        }

        /// <summary>
        /// Connect with device
        /// </summary>
        private void SIConnect()
        {
            ReaderDeviceInfo inDevice = null;
            ReaderDeviceInfo outDevice = null;

            if (!siDeviceInfoList.ContainsKey(cmbCOMPorts.SelectedIndex) ||
                !ReaderDeviceInfo.IsDeviceValid(siDeviceInfoList[cmbCOMPorts.SelectedIndex].DeviceName))
            {
                MessageBox.Show(Resource.MsgBoxErrorDeviceDetermine, Resource.MsgBoxErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                inDevice = new ReaderDeviceInfo(siDeviceInfoList[cmbCOMPorts.SelectedIndex], ReaderDeviceType.SiDevice);
                outDevice = new ReaderDeviceInfo(ReaderDeviceType.None);

                _SIReader.InputDevice = inDevice;
                _SIReader.OutputDevice = outDevice;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(Resource.MsgBoxErrorDeviceException, inDevice, ex.Message), Resource.MsgBoxErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                _SIReader.OpenInputDevice();
            }
            catch (Exception)
            {
                MessageBox.Show(Resource.MsgBoxErrorDeviceOpen, Resource.MsgBoxErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }


            Console.WriteLine("Connected to " + inDevice);

            grpFormInput.IsEnabled = false;
            btnConnect.Content = Resource.BtnDisconnect;
            connected = true;
        }

        /// <summary>
        /// Disconnect from device
        /// </summary>
        private void SIDisconnect()
        {
            if (_SIReader.InputDeviceIsOpen) _SIReader.CloseInputDevice();
            grpFormInput.IsEnabled = true;
            txtInfo.Text = "";
            btnConnect.Content = Resource.BtnConnect;
            connected = false;
        }
        #endregion

        #region Reader event handlers

        /// <summary>
        /// Handles the event that is thrown when the reader class read a card completely
        /// </summary>
        private void siReader_CardRead(object sender, SportidentDataEventArgs e)
        {
            foreach (var card in e.Cards)
            {
                Console.WriteLine(string.Format("Punch from SI:{0} at {1}", card.Siid, card.ReadoutDateTime));
                Application.Current.Dispatcher.Invoke(new Action(() => {
                    // SSID
                    txbNumber.Text = card.Siid;

                    // Battery Voltage
                    txbVoltage.Text = card.BatteryVoltage.ToString() + "V";
                    if (card.BatteryVoltage<=card.BatteryLowThreshold)
                    {
                        lbVoltageStatus.Content = "✘";
                        recVoltageBackground.Fill = Brushes.IndianRed;
                        txbVoltageDescription.Text = string.Format(Resource.OutputResultVoltageError, card.BatteryLowThreshold.ToString());
                    }
                    else if (card.BatteryVoltage <= 2.72)
                    {
                        lbVoltageStatus.Content = "❗";
                        recVoltageBackground.Fill = Brushes.LightSalmon;
                        txbVoltageDescription.Text = string.Format(Resource.OutputResultVoltageWarning, 2.72);
                    }
                    else
                    {
                        lbVoltageStatus.Content = "✔";
                        recVoltageBackground.Fill = Brushes.LightGreen;
                        txbVoltageDescription.Text = string.Format(Resource.OutputResultVoltageOK, card.BatteryLowThreshold.ToString());
                    }

                    // Battery date
                    DateTime thisDate = DateTime.Today;
                    DateTime BatteryChangeDate = card.BatteryDate.AddYears(3);
                    txbDate.Text = card.BatteryDate.ToShortDateString();
                    if ((BatteryChangeDate - thisDate).Days < 0)
                    {
                        lbDateStatus.Content = "✘";
                        recDateBackground.Fill = Brushes.IndianRed;
                        txbDateDescription.Text = string.Format(Resource.OutputResultDateError, 3);
                    }
                    else if ((BatteryChangeDate - thisDate).Days < 60)
                    {
                        lbDateStatus.Content = "❗";
                        recDateBackground.Fill = Brushes.LightSalmon;
                        txbDateDescription.Text = string.Format(Resource.OutputResultDateWaring, 3, BatteryChangeDate.ToString("MMM yyy"));
                    }
                    else
                    {
                        lbDateStatus.Content = "✔";
                        recDateBackground.Fill = Brushes.LightGreen;
                        txbDateDescription.Text = string.Format(Resource.OutputResultDateOK, BatteryChangeDate.ToString("MMM yyy"));
                    }
                }));
            }

            // Start timer to clear data
            ClearTimer.Elapsed += new ElapsedEventHandler(ClearDataAfterTime);
            ClearTimer.Interval = 5000;
            ClearTimer.Enabled = true;
        }

        /// <summary>
        /// Handles the event that is thrown when the reader class logs a message (info, warning, error...)
        /// </summary>
        private static void siReader_LogEvent(object sender, FileLoggerEventArgs e)
        {
            Console.WriteLine("Log: " + e.Message);
        }

        /// <summary>
        /// Handles the event that is thrown when the reader class indicates a state change for the input device
        /// </summary>
        private void siReader_InputDeviceStateChanged(object sender, ReaderDeviceStateChangedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                switch (e.CurrentState)
                {
                    case DeviceState.D0Online:
                        sbarSIStatus.Background = Brushes.Green;
                        break;
                    case DeviceState.D5Busy:
                        sbarSIStatus.Background = Brushes.Yellow;
                        break;
                    default:
                        sbarSIStatus.Background = Brushes.Red;
                        break;
                }

            }));
        }

        /// <summary>
        /// Handles the event that is thrown when the reader class indicates that the input device has changed
        /// </summary>
        private void siReader_InputDeviceChanged(object sender, ReaderDeviceChangedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                var inputSource = string.Empty;

                switch (e.CurrentDevice.ReaderDeviceType)
                {
                    case ReaderDeviceType.SiDevice:
                        inputSource = DeviceInfo.GetPrettyDeviceName(e.CurrentDevice);
                        break;
                    default:
                        inputSource = e.CurrentDevice.ReaderDeviceType.ToString();
                        break;
                }

                txtInfo.Text = string.Format(Resource.InfoInputSource, inputSource);
            }));
        }

        /// <summary>
        /// Handles the event when an error occures
        /// </summary>
        private void siReader_ErrorOccured(object sender, FileLoggerEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                MessageBox.Show(string.Format("{0}\n\n{1}", e.Message, e.ThrownException.Message), Resource.MsgBoxErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);

                if (e.ThrownException.Message != "Disconnected" || !connected) return;

                SIDisconnect();
            }));
        }

        /// <summary>
        /// Handles the event that is thrown when the reader class indicates that the stations config has been read successfully
        /// </summary>
        private void siReader_DeviceConfigurationRead(object sender, StationConfigurationEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
            {
                txtInfo.Text = Resource.InfoUnknownDevice;

                var msg = Resource.MsgBoxErrorConfigDescriptionNo;
                var failed = false;
                switch (e.Result)
                {
                    case StationConfigurationResult.OperatingModeNotSupported:
                        msg = Resource.MsgBoxErrorConfigDescriptionOpMode;
                        failed = true;
                        break;
                    case StationConfigurationResult.DeviceDoesNotHaveBackup:
                        msg = Resource.MsgBoxErrorConfigDescriptionNoBackup;
                        failed = true;
                        break;
                    case StationConfigurationResult.ReadoutMasterBackupNotSupported:
                        msg = Resource.MsgBoxErrorConfigDescriptionBackupReadout;
                        failed = true;
                        break;
                }

                if (failed)
                {
                    MessageBox.Show(
                        string.Format(
                            Resource.MsgBoxErrorConfig,
                            e.Result, msg), Resource.MsgBoxErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);

                    return;
                }

                switch (e.Device.Product.ProductFamily)
                {
                    case ProductFamily.SimSrr:
                        txtInfo.Text = string.Format(Resource.InfoConnectedSimSrr,
                            e.Device.SerialNumber, e.Device.FirmwareVersion, e.Device.Product.ProductType,
                            e.Device.SimSrrUseModD3Protocol, e.Device.SimSrrChannel);
                        break;
                    case ProductFamily.Bs8SiMaster:
                    case ProductFamily.Bsx7:
                    case ProductFamily.Bsx8:
                        txtInfo.Text =
                            string.Format(Resource.InfoConnectedMaster,
                                e.Device.CodeNumber, e.Device.OperatingMode, e.Device.AutoSendMode, e.Device.LegacyProtocolMode);

                        // Check device configuration
                        if (e.Device.OperatingMode != OperatingMode.Readout)
                        {
                            MessageBox.Show(Resource.MsgBoxErrorConfigRedout, Resource.MsgBoxErrorConfigTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);

                        }
                        else if (e.Device.AutoSendMode)
                        {
                            MessageBox.Show(Resource.MsgBoxErrorConfigAutosend, Resource.MsgBoxErrorConfigTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);

                        }
                        else if (e.Device.LegacyProtocolMode)
                        {
                            MessageBox.Show(Resource.MsgBoxErrorConfigLegacy, Resource.MsgBoxErrorConfigTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        }

                        break;
                    default:
                        MessageBox.Show(Resource.MsgBoxErrorConfigNoDevice, Resource.MsgBoxErrorConfigTitle, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        break;
                }
            }));
        }
        #endregion

    }

}
