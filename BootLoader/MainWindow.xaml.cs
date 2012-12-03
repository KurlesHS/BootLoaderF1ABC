using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.ComponentModel;
using System.IO.Ports;
using System.Xml;
using System.Runtime.InteropServices;


namespace BootLoader
{


    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    public partial class MainWindow : Window
    {
        private string hexFilename = "";
        private string serialPort = "";
        private byte[] buffer = new byte[0x10000];
        private BackgroundWorker bgWorker = new BackgroundWorker();
        string[] ports;
        string settingFile;

        public MainWindow()
        {
            InitializeComponent();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            ports = SerialPort.GetPortNames();
            bgWorker.DoWork += bgWorker_DoWork;
            bgWorker.RunWorkerCompleted += bgWorker_RunWorkerCompleted;
            bgWorker.ProgressChanged += bgWorker_ProgressChanged;
            comboboxForPortsNames.ItemsSource = ports;
            if (ports.Length > 0)
            {
                comboboxForPortsNames.SelectedItem = ports[0];
                serialPort = ports[0];
            }
            settingFile = AppDomain.CurrentDomain.BaseDirectory;
            if (settingFile.Length > 0)
                if (settingFile.Substring(settingFile.Length - 1, 1) != "\\")
                    settingFile += "\\";
            settingFile += "settings.xml";
            progressBar.Text = "Выберите файл";
            bool settingFileIsPresents = true;
            try
            {
                XmlDocument xd = new XmlDocument();
                xd.Load(settingFile);
                foreach (XmlNode node in xd.DocumentElement.ChildNodes)
                {
                    if (node.Name == "FileName")
                        hexFilename = node.InnerText;
                    if (node.Name == "SerialPort")
                        serialPort = node.InnerText;
                }
                parseHexFile();
                comboboxForPortsNames.SelectedItem = serialPort;
            }
            catch (Exception)
            {
                settingFileIsPresents = false; 
            }
            if (!settingFileIsPresents)
            {
                updateSettings();
            }
        }

        void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            return;
        }

        void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                progressBar.Text = "Операция отменена";
            }
            else if (e.Error != null)
            {
                progressBar.Text = String.Format("Ошибка: {0}", e.Error.Message);
            }
            else
            {
                progressBar.Text = e.Result.ToString();
            }

            ButtonSelectFile.IsEnabled = true;
            ButtonStartFlashing.IsEnabled = true;
            comboboxForPortsNames.IsEnabled = true;
        }

        void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (worker == null)
                return;
            string portName = e.Argument.ToString();
            string portStatus = String.Format("Порт {0} открыт", portName);
            using (SerialPort sp = new SerialPort(portName))
            {
                try
                {
                    sp.Open();
                    sp.Parity = System.IO.Ports.Parity.None;
                    sp.DataBits = 8;
                    sp.BaudRate = 9600;
                    sp.Handshake = Handshake.None;
                    sp.StopBits = StopBits.One;
                    sp.DataReceived += onSerialDataReceived;
                }
                catch (Exception)
                {
                    e.Result = String.Format("Не получилось открыть {0} порт", portName);
                    return;
                }
                setTextForProgressBar(portStatus);
               
                try
                {
                    for (int i = 0; i <= 100; ++i)
                    {
                        System.Threading.Thread.Sleep(10);
                        worker.ReportProgress(i);

                        if (i == 30)
                            setTextForProgressBar(String.Format("Прошиваем через {0} порт", portName));
                    }
                }
                catch (Exception exc)
                {

                    e.Result = exc.Message;
                    return;
                }
                e.Result = "Устройство прошито успешно";
                return;
            }
        }

        void onSerialDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = sender as SerialPort;
            if (sp == null)
                return;
            //sp.Read()
            return;
        }

        private void updateSettings()
        {
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings();

                // включаем отступ для элементов XML документа
                // (позволяет наглядно изобразить иерархию XML документа)
                settings.Indent = true;
                settings.IndentChars = "    "; // задаем отступ, здесь у меня 4 пробела

                // задаем переход на новую строку
                settings.NewLineChars = "\n";

                // Нужно ли опустить строку декларации формата XML документа
                // речь идет о строке вида "<?xml version="1.0" encoding="utf-8"?>"
                settings.OmitXmlDeclaration = false;
                using (XmlWriter xw = XmlWriter.Create(settingFile, settings))
                {
                    xw.WriteStartElement("Flasher");

                    xw.WriteElementString("FileName", hexFilename);
                    xw.WriteElementString("SerialPort", serialPort);
                    xw.WriteEndElement();
                    xw.Flush();
                    xw.Close();
                }
            }
            catch (Exception)
            {

            }
        }

        delegate void delegateWithoutParam();
        delegate void delegateWithIntParam(int value);
        delegate void delegateWhthStringParam(string value);

        private void setMaxValueForProgressBar(int value)
        {
            progressBar.Dispatcher.BeginInvoke(new delegateWithIntParam(delegate(int x){progressBar.Maximum = x;}), value);
        }

        private void setValueForProgressBar(int value)
        {
            progressBar.Dispatcher.BeginInvoke(new delegateWithIntParam(delegate(int x) { progressBar.Value = x; }), value);
        }

        private void setTextForProgressBar(string text)
        {
            progressBar.Dispatcher.BeginInvoke(new delegateWhthStringParam(delegate(string x) { progressBar.Text = x; }), text);
        }

        protected static ushort GetChecksum(byte[] bytes, int startAddress = 0)
        {
            ulong sum = 0;
            // Sum all the words together, adding the final byte if size is odd
            int i = startAddress;
            for (; i < bytes.Length - 1; i += 2)
            {
                sum += BitConverter.ToUInt16(bytes, i);
            }
            if (i != bytes.Length)
                sum += bytes[i];
            // Do a little shuffling
            sum = (sum >> 16) + (sum & 0xFFFF);
            sum += (sum >> 16);
            return (ushort)(~sum);
        }

        private void encryptDecrypt(int startAddres = 0)
        {
            byte c = 55;
            for (int i = startAddres; i < buffer.Length; ++i)
            {
                buffer[i] ^= c;
                c += 34;
            }
        }
        delegate void ChangeProgressBarValue(int value);
        public void changeProgerssBarValue(int value)
        {
            if (progressBar.Dispatcher.CheckAccess())
                progressBar.Value = value;
            else
            {
                progressBar.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal,
                    new ChangeProgressBarValue(changeProgerssBarValue),
                    value);
            }
        }

        private void ButtonSelectFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".hex";
            dlg.Filter = "HEX files (*.hex)|*.hex";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                hexFilename = dlg.FileName;
                parseHexFile();
                updateSettings();
            }
        }

        private void parseHexFile()
        {
            labelForFileName.Text = hexFilename;
            for (int i = 0; i < 0x10000; ++i)
                buffer[i] = 0xff;
            StreamReader sr = new StreamReader(hexFilename);
            int minAddress = 0x10000;
            int maxAddress = 0;
            bool converted = true;
            while (!sr.EndOfStream)
            {
                string line = sr.ReadLine();

                try
                {
                    if (line.Substring(0, 1) != ":")
                        throw new Exception();
                    byte crc = 0;
                    // считаем контрольную сумму
                    for (int x = 1; x < line.Length; x += 2)
                    {
                        byte val = Convert.ToByte(line.Substring(x, 2), 16);
                        crc += val;
                    }
                    if (crc != 0)
                        throw new Exception();
                    // контрольная сумма совпала
                    int lenData = Convert.ToInt32(line.Substring(1, 2), 16);
                    int startAddress = Convert.ToInt32(line.Substring(3, 4), 16);
                    int type = Convert.ToInt32(line.Substring(7, 2), 16);
                    switch (type)
                    {
                        case 5:
                            {
                                // pc counter. Игнорируем
                            }
                            break;
                        case 1:
                            {
                                // конец файла
                                if (lenData != 0)
                                    throw new Exception();
                            }
                            break;
                        case 0:
                            {
                                // данные
                                // заполняем буфер
                                for (int i = 0; i < lenData; ++i)
                                    buffer[i + startAddress] = Convert.ToByte(line.Substring(i * 2 + 9, 2), 16);
                                //корректируем занчения
                                if (minAddress > startAddress)
                                    minAddress = startAddress;
                                if (maxAddress < (startAddress + lenData))
                                    maxAddress = startAddress + lenData;
                            }
                            break;
                        default:
                            throw new Exception();
                    }
                }
                catch (Exception)
                {
                    converted = false;
                    break;
                }
            }
            if (!converted)
            {
                progressBar.Text = "Не верный формат hex файла";
                ButtonStartFlashing.IsEnabled = false;
            }
            else
            {
                ButtonStartFlashing.IsEnabled = true;
                ushort crc = GetChecksum(buffer, minAddress);
                // добавляем крк в конец файла
                buffer[0xffff] = (byte)(crc >> 8);
                buffer[0xfffe] = (byte)crc;
                encryptDecrypt(0x8000);
                Debug.WriteLine(String.Format("minAddres = {0}, maxAddress = {1}", minAddress, maxAddress));
                progressBar.Text = "Всё готово к прошивке";
            }
        }

        private void comboboxForPortsNames_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            serialPort = comboboxForPortsNames.SelectedItem.ToString();
            updateSettings();
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetCursorPos(ref Win32Point pt);

        [StructLayout(LayoutKind.Sequential)]
        internal struct Win32Point
        {
            public Int32 X;
            public Int32 Y;
        };
        public static Point GetMousePosition()
        {
            Win32Point w32Mouse = new Win32Point();
            GetCursorPos(ref w32Mouse);
            return new Point(w32Mouse.X, w32Mouse.Y);
        }
        private void ButtonCloseWindow_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private bool mouseIsCaptured;
        private Point lastMousePos;

        private void Window_MouseMove_1(object sender, MouseEventArgs e)
        {
            if (mouseIsCaptured)
            {
                Point curMousePos = GetMousePosition();
                double deltax = lastMousePos.X - curMousePos.X;
                double deltay = lastMousePos.Y - curMousePos.Y;
                lastMousePos.X = curMousePos.X;
                lastMousePos.Y = curMousePos.Y;
                this.Top -= deltay;
                this.Left -= deltax;
            }
        }

        private void TextBlock_MouseLeftButtonDown_1(object sender, MouseButtonEventArgs e)
        {
            //this.DragMove();
            UIElement el = sender as UIElement;
            if (el == null)
                return;
            if (e.LeftButton == MouseButtonState.Pressed)
                mouseIsCaptured =  el.CaptureMouse();
            lastMousePos = GetMousePosition();
        }

        private void TextBlock_MouseLeftButtonUp_1(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Released)
            {
                mouseIsCaptured = false;
                UIElement el = sender as UIElement;
                if (el == null)
                    return;
                el.ReleaseMouseCapture();
            }
        }

        private void ButtonStartFlashing_Click(object sender, RoutedEventArgs e)
        {
            ButtonSelectFile.IsEnabled = false;
            ButtonStartFlashing.IsEnabled = false;
            comboboxForPortsNames.IsEnabled = false;
            progressBar.Text = "Идет прошивка, подождите...";
            bgWorker.RunWorkerAsync(serialPort);
        }
    }
}
