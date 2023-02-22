using System;
using System.IO;
using System.Windows.Forms;
using Vintasoft.WinTwain;

namespace TwainFileTransferDemo
{
    public partial class MainForm : Form
    {

        #region Fields

        /// <summary>
        /// TWAIN device manager.
        /// </summary>
        DeviceManager _deviceManager;

        /// <summary>
        /// Current device.
        /// </summary>
        Device _currentDevice;

        /// <summary>
        /// Path to directory where acquired images will be saved.
        /// </summary>
        string _directoryForImages;
        /// <summary>
        /// Index of acquired image.
        /// </summary>
        int _imageIndex;

        #endregion



        #region Constructors

        public MainForm()
        {
            InitializeComponent();

            this.Text = string.Format("VintaSoft TWAIN File Transfer Demo v{0}", TwainGlobalSettings.ProductVersion);

            // create instance of the DeviceManager class
            _deviceManager = new DeviceManager(this, this.Handle);
        }

        #endregion



        #region Methods

        /// <summary>
        /// Application form is shown.
        /// </summary>
        private void MainForm_Shown(object sender, EventArgs e)
        {
            // get path to directory where acquired images will be saved

            _directoryForImages = Path.GetDirectoryName(Application.ExecutablePath);
            _directoryForImages = Path.Combine(_directoryForImages, "Images");
            if (!Directory.Exists(_directoryForImages))
                Directory.CreateDirectory(_directoryForImages);

            directoryForImagesTextBox.Text = _directoryForImages;

            // open TWAIN device manager
            if (OpenDeviceManager())
            {
                // fill the device list

                devicesComboBox.Items.Clear();
                DeviceInfo deviceInfo;
                DeviceCollection devices = _deviceManager.Devices;
                for (int i = 0; i < devices.Count; i++)
                {
                    deviceInfo = devices[i].Info;
                    devicesComboBox.Items.Add(deviceInfo.ProductName);

                    if (devices[i] == _deviceManager.DefaultDevice)
                        devicesComboBox.SelectedIndex = i;
                }
            }
        }


        /// <summary>
        /// Sets form's UI state.
        /// </summary>
        private void SetFormUiState(bool enabled)
        {
            devicesComboBox.Enabled = enabled;
            deviceSettingsGroupBox.Enabled = enabled;
            acquireImageWithUIButton.Enabled = enabled;
            acquireImageWithUIButton.Enabled = enabled;
            acquireImageWithoutUIButton.Enabled = enabled;
        }


        /// <summary>
        /// Opens TWAIN device manager.
        /// </summary>
        private bool OpenDeviceManager()
        {
            SetFormUiState(false);

            try
            {
                // try to find the device manager 2.x
                _deviceManager.IsTwain2Compatible = true;
                // if TWAIN device manager 2.x is NOT available
                if (!_deviceManager.IsTwainAvailable)
                {
                    // try to find the device manager 1.x
                    _deviceManager.IsTwain2Compatible = false;
                    // if TWAIN device manager 1.x is NOT available
                    if (!_deviceManager.IsTwainAvailable)
                    {
                        MessageBox.Show("TWAIN device manager is not found.");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                // show dialog with error message
                MessageBox.Show(GetFullExceptionMessage(ex), "TWAIN device manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // if 64-bit TWAIN2 device manager is used
            if (IntPtr.Size == 8 && _deviceManager.IsTwain2Compatible)
            {
                if (!InitTwain2DeviceManagerMode())
                    return false;
            }

            try
            {
                // open the device manager
                _deviceManager.Open();
            }
            catch (Exception ex)
            {
                // show dialog with error message
                MessageBox.Show(GetFullExceptionMessage(ex), "TWAIN device manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // if no devices are found in the system
            if (_deviceManager.Devices.Count == 0)
            {
                MessageBox.Show("No devices found.");
                return false;
            }

            SetFormUiState(true);
            return true;
        }

        /// <summary>
        /// Initializes the device manager mode.
        /// </summary>
        private bool InitTwain2DeviceManagerMode()
        {
            // create a form that allows to view and edit mode of 64-bit TWAIN2 device manager
            using (SelectDeviceManagerModeForm form = new SelectDeviceManagerModeForm())
            {
                // initialize form
                form.StartPosition = FormStartPosition.CenterParent;
                form.Owner = this;
                form.Use32BitDevices = _deviceManager.Are32BitDevicesUsed;

                // show dialog
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // if device manager mode is changed
                    if (form.Use32BitDevices != _deviceManager.Are32BitDevicesUsed)
                    {
                        try
                        {
                            // if 32-bit devices must be used
                            if (form.Use32BitDevices)
                                _deviceManager.Use32BitDevices();
                            else
                                _deviceManager.Use64BitDevices();
                        }
                        catch (TwainDeviceManagerException ex)
                        {
                            // show dialog with error message
                            MessageBox.Show(GetFullExceptionMessage(ex), "TWAIN device manager", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            return false;
                        }
                    }
                }
                else
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Selects directory for acquired images.
        /// </summary>
        private void selectDirectoryForImagesButton_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = _directoryForImages;
            if (folderBrowserDialog1.ShowDialog() != DialogResult.OK)
                return;

            _directoryForImages = folderBrowserDialog1.SelectedPath;
            directoryForImagesTextBox.Text = _directoryForImages;
        }

        /// <summary>
        /// Current device is changed.
        /// </summary>
        private void devicesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetFormUiState(false);

            // get device
            Device device = _deviceManager.Devices.Find((string)devicesComboBox.SelectedItem);

            // get file formats and compressions supported by device
            GetSupportedFileFormatsAndCompressions(device);

            SetFormUiState(true);
        }

        /// <summary>
        /// Gets file formats and compressions supported by device in File Transfer mode.
        /// </summary>
        private void GetSupportedFileFormatsAndCompressions(Device device)
        {
            try
            {
                // open the device
                device.Open();


                // get supported file formats

                TwainImageFileFormat currentFileFormat = device.FileFormat;
                TwainImageFileFormat[] fileFormats = device.GetSupportedImageFileFormats();

                supportedFileFormatsComboBox.Items.Clear();
                for (int i = 0; i < fileFormats.Length; i++)
                {
                    supportedFileFormatsComboBox.Items.Add(fileFormats[i]);

                    if (currentFileFormat == fileFormats[i])
                        supportedFileFormatsComboBox.SelectedIndex = i;
                }

                // get supported compressions

                TwainImageCompression currentCompression = device.ImageCompression;
                TwainImageCompression[] compressions = device.GetSupportedImageCompressions();

                supportedCompressionsComboBox.Items.Clear();
                for (int i = 0; i < compressions.Length; i++)
                {
                    supportedCompressionsComboBox.Items.Add(compressions[i]);

                    if (currentCompression == compressions[i])
                        supportedCompressionsComboBox.SelectedIndex = i;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(GetFullExceptionMessage(ex));
            }
            finally
            {
                // close the device
                device.Close();
            }
        }


        /// <summary>
        /// Acquires image with UI.
        /// </summary>
        private void acquireImageWithUIButton_Click(object sender, EventArgs e)
        {
            AcquireImage(true);
        }

        /// <summary>
        /// Acquires image without UI.
        /// </summary>
        private void acquireImageWithoutUIButton_Click(object sender, EventArgs e)
        {
            AcquireImage(false);
        }

        /// <summary>
        /// Acquires image.
        /// </summary>
        private void AcquireImage(bool showUI)
        {
            SetFormUiState(false);

            if (_currentDevice != null)
                UnsubscribeFromDeviceEvents();

            Device device = _deviceManager.Devices.Find((string)devicesComboBox.SelectedItem);

            _currentDevice = device;
            // subscribe to the device events
            SubscribeToDeviceEvents();

            try
            {
                // set settings of scan session
                device.TransferMode = TransferMode.File;
                device.ShowUI = showUI;
                device.DisableAfterAcquire = !showUI;

                // open the device
                device.Open();

                try
                {
                    // set the file format in which acquired images must be saved
                    TwainImageFileFormat newFileFormat = TwainImageFileFormat.Bmp;
                    if (supportedFileFormatsComboBox.Items.Count > 0)
                        newFileFormat = (TwainImageFileFormat)supportedFileFormatsComboBox.SelectedItem;
                    if (device.FileFormat != newFileFormat)
                        device.FileFormat = newFileFormat;

                    TwainImageCompression newImageCompression = TwainImageCompression.None;
                    if (supportedCompressionsComboBox.Items.Count > 0)
                        newImageCompression = (TwainImageCompression)supportedCompressionsComboBox.SelectedItem;
                    if (device.ImageCompression != newImageCompression)
                        device.ImageCompression = newImageCompression;
                }
                catch
                {
                }

                // start the asynchronous image acquisition process
                device.Acquire();
            }
            catch (TwainDeviceException ex)
            {
                // close the device
                _currentDevice.Close();
                MessageBox.Show(GetFullExceptionMessage(ex));
                SetFormUiState(true);
                return;
            }
        }

        /// <summary>
        /// Subscribes to the device events.
        /// </summary>
        private void SubscribeToDeviceEvents()
        {
            _currentDevice.ImageAcquiring += new EventHandler<ImageAcquiringEventArgs>(device_ImageAcquiring);
            _currentDevice.ImageAcquired += new EventHandler<ImageAcquiredEventArgs>(device_ImageAcquired);
            _currentDevice.ScanCompleted += new EventHandler(device_ScanCompleted);
            _currentDevice.ScanCanceled += new EventHandler(device_ScanCanceled);
            _currentDevice.ScanFailed += new EventHandler<ScanFailedEventArgs>(device_ScanFailed);
            _currentDevice.UserInterfaceClosed += new EventHandler(device_UserInterfaceClosed);
            _currentDevice.ScanFinished += new EventHandler(device_ScanFinished);
        }

        /// <summary>
        /// Unsubscribes from the device events.
        /// </summary>
        private void UnsubscribeFromDeviceEvents()
        {
            _currentDevice.ImageAcquiring -= new EventHandler<ImageAcquiringEventArgs>(device_ImageAcquiring);
            _currentDevice.ImageAcquired -= new EventHandler<ImageAcquiredEventArgs>(device_ImageAcquired);
            _currentDevice.ScanCompleted -= new EventHandler(device_ScanCompleted);
            _currentDevice.ScanCanceled -= new EventHandler(device_ScanCanceled);
            _currentDevice.ScanFailed -= new EventHandler<ScanFailedEventArgs>(device_ScanFailed);
            _currentDevice.UserInterfaceClosed -= new EventHandler(device_UserInterfaceClosed);
            _currentDevice.ScanFinished -= new EventHandler(device_ScanFinished);
        }

        /// <summary>
        /// Image is acquiring.
        /// </summary>
        void device_ImageAcquiring(object sender, ImageAcquiringEventArgs e)
        {
            string fileExtension = "bmp";
            switch (e.FileFormat)
            {
                case TwainImageFileFormat.Tiff:
                    fileExtension = "tif";
                    break;

                case TwainImageFileFormat.Jpeg:
                    fileExtension = "jpg";
                    break;
            }

            e.Filename = Path.Combine(_directoryForImages, string.Format("page{0}.{1}", _imageIndex, fileExtension));
        }

        /// <summary>
        /// Image is acquired.
        /// </summary>
        void device_ImageAcquired(object sender, ImageAcquiredEventArgs e)
        {
            statusTextBox.Text += string.Format("Image is saved to file '{0}'{1}", Path.GetFileName(e.Filename), Environment.NewLine);
            _imageIndex++;
        }

        /// <summary>
        /// Scan is completed.
        /// </summary>
        void device_ScanCompleted(object sender, EventArgs e)
        {
            statusTextBox.Text += string.Format("Scan is completed{0}", Environment.NewLine);
        }

        /// <summary>
        /// Scan is canceled.
        /// </summary>
        void device_ScanCanceled(object sender, EventArgs e)
        {
            statusTextBox.Text += string.Format("Scan is canceled{0}", Environment.NewLine);
        }

        /// <summary>
        /// User interface of device is closed.
        /// </summary>
        void device_UserInterfaceClosed(object sender, EventArgs e)
        {
            statusTextBox.Text += string.Format("User Interface is closed{0}", Environment.NewLine);
        }

        /// <summary>
        /// Scan is failed.
        /// </summary>
        void device_ScanFailed(object sender, ScanFailedEventArgs e)
        {
            statusTextBox.Text += string.Format("Scan is failed: {0}{1}", e.ErrorString, Environment.NewLine);
        }

        /// <summary>
        /// Scan is finished.
        /// </summary>
        void device_ScanFinished(object sender, EventArgs e)
        {
            // close the device
            _currentDevice.Close();

            SetFormUiState(true);
        }


        /// <summary>
        /// Application form is closing.
        /// </summary>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_currentDevice != null)
            {
                UnsubscribeFromDeviceEvents();
                _currentDevice = null;
            }

            try
            {
                // close the device manager
                _deviceManager.Close();
                // dispose the device manager
                _deviceManager.Dispose();
            }
            catch
            {
            }
        }

        /// <summary>
        /// Returns the message of exception and inner exceptions.
        /// </summary>
        private string GetFullExceptionMessage(Exception ex)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            sb.AppendLine(ex.Message);

            Exception innerException = ex.InnerException;
            while (innerException != null)
            {
                if (ex.Message != innerException.Message)
                    sb.AppendLine(string.Format("Inner exception: {0}", innerException.Message));
                innerException = innerException.InnerException;
            }

            return sb.ToString();
        }

        #endregion

    }
}
