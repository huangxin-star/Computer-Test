﻿using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Computer_Test
{
    public partial class MainForm : Form
    {
        private OpenFileDialog ofd = new OpenFileDialog();
        private FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
        string mappingTextBoxName = "";
        string mappingFilePath = "";
        string modelTableTextBoxName = "";
        string savePathTextBoxName = "";
        string jizhongming = "";
        string startupPath = Application.StartupPath; //System.Diagnostics.Process.Start(Application.StartupPath + "//" + 文件名);
        private Timer timer;
        XSSFWorkbook WB;
        

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            initFiles();
            timer = new Timer();
            timer.Interval = 1000; // 设置定时器间隔，单位为毫秒
            timer.Tick += Timer_Tick;
            timer.Start(); // 启动定时器
        }
        private void initFiles()
        {
            // 获取当前目录下的指定文件路径
            string currentDirectory = Directory.GetCurrentDirectory();
            string specificFile = Path.Combine(currentDirectory, "FT映射表.xlsx");
            mapTextBox.Text = specificFile;
            mappingFilePath = specificFile;
            // 获取桌面路径
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            savaPathTextBox.Text = desktopPath;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            // 每次定时器触发，更新当前系统时间到 Label 控件
            Time.Text = "当前时间：" + DateTime.Now.ToString();
        }

        public class Reflector
        {
            #region variables

            string m_ns;
            Assembly m_asmb;

            #endregion

            #region Constructors

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="ns">The namespace containing types to be used</param>
            public Reflector(string ns)
                : this(ns, ns)
            { }

            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="an">A specific assembly name (used if the assembly name does not tie exactly with the namespace)</param>
            /// <param name="ns">The namespace containing types to be used</param>
            public Reflector(string an, string ns)
            {
                m_ns = ns;
                m_asmb = null;
                foreach (AssemblyName aN in Assembly.GetExecutingAssembly().GetReferencedAssemblies())
                {
                    if (aN.FullName.StartsWith(an))
                    {
                        m_asmb = Assembly.Load(aN);
                        break;
                    }
                }
            }

            #endregion

            #region Methods

            /// <summary>
            /// Return a Type instance for a type 'typeName'
            /// </summary>
            /// <param name="typeName">The name of the type</param>
            /// <returns>A type instance</returns>
            public Type GetType(string typeName)
            {
                Type type = null;
                string[] names = typeName.Split('.');

                if (names.Length > 0)
                    type = m_asmb.GetType(m_ns + "." + names[0]);

                for (int i = 1; i < names.Length; ++i)
                {
                    type = type.GetNestedType(names[i], BindingFlags.NonPublic);
                }
                return type;
            }

            /// <summary>
            /// Create a new object of a named type passing along any params
            /// </summary>
            /// <param name="name">The name of the type to create</param>
            /// <param name="parameters"></param>
            /// <returns>An instantiated type</returns>
            public object New(string name, params object[] parameters)
            {
                Type type = GetType(name);

                ConstructorInfo[] ctorInfos = type.GetConstructors();
                foreach (ConstructorInfo ci in ctorInfos)
                {
                    try
                    {
                        return ci.Invoke(parameters);
                    }
                    catch { }
                }

                return null;
            }

            /// <summary>
            /// Calls method 'func' on object 'obj' passing parameters 'parameters'
            /// </summary>
            /// <param name="obj">The object on which to excute function 'func'</param>
            /// <param name="func">The function to execute</param>
            /// <param name="parameters">The parameters to pass to function 'func'</param>
            /// <returns>The result of the function invocation</returns>
            public object Call(object obj, string func, params object[] parameters)
            {
                return Call2(obj, func, parameters);
            }

            /// <summary>
            /// Calls method 'func' on object 'obj' passing parameters 'parameters'
            /// </summary>
            /// <param name="obj">The object on which to excute function 'func'</param>
            /// <param name="func">The function to execute</param>
            /// <param name="parameters">The parameters to pass to function 'func'</param>
            /// <returns>The result of the function invocation</returns>
            public object Call2(object obj, string func, object[] parameters)
            {
                return CallAs2(obj.GetType(), obj, func, parameters);
            }

            /// <summary>
            /// Calls method 'func' on object 'obj' which is of type 'type' passing parameters 'parameters'
            /// </summary>
            /// <param name="type">The type of 'obj'</param>
            /// <param name="obj">The object on which to excute function 'func'</param>
            /// <param name="func">The function to execute</param>
            /// <param name="parameters">The parameters to pass to function 'func'</param>
            /// <returns>The result of the function invocation</returns>
            public object CallAs(Type type, object obj, string func, params object[] parameters)
            {
                return CallAs2(type, obj, func, parameters);
            }

            /// <summary>
            /// Calls method 'func' on object 'obj' which is of type 'type' passing parameters 'parameters'
            /// </summary>
            /// <param name="type">The type of 'obj'</param>
            /// <param name="obj">The object on which to excute function 'func'</param>
            /// <param name="func">The function to execute</param>
            /// <param name="parameters">The parameters to pass to function 'func'</param>
            /// <returns>The result of the function invocation</returns>
            public object CallAs2(Type type, object obj, string func, object[] parameters)
            {
                MethodInfo methInfo = type.GetMethod(func, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                return methInfo.Invoke(obj, parameters);
            }

            /// <summary>
            /// Returns the value of property 'prop' of object 'obj'
            /// </summary>
            /// <param name="obj">The object containing 'prop'</param>
            /// <param name="prop">The property name</param>
            /// <returns>The property value</returns>
            public object Get(object obj, string prop)
            {
                return GetAs(obj.GetType(), obj, prop);
            }

            /// <summary>
            /// Returns the value of property 'prop' of object 'obj' which has type 'type'
            /// </summary>
            /// <param name="type">The type of 'obj'</param>
            /// <param name="obj">The object containing 'prop'</param>
            /// <param name="prop">The property name</param>
            /// <returns>The property value</returns>
            public object GetAs(Type type, object obj, string prop)
            {
                PropertyInfo propInfo = type.GetProperty(prop, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                return propInfo.GetValue(obj, null);
            }

            /// <summary>
            /// Returns an enum value
            /// </summary>
            /// <param name="typeName">The name of enum type</param>
            /// <param name="name">The name of the value</param>
            /// <returns>The enum value</returns>
            public object GetEnum(string typeName, string name)
            {
                Type type = GetType(typeName);
                FieldInfo fieldInfo = type.GetField(name);
                return fieldInfo.GetValue(null);
            }

            #endregion

        }
        public class FolderSelectDialog
        {
            // Wrapped dialog
            System.Windows.Forms.OpenFileDialog ofd = null;

            /// <summary>
            /// Default constructor
            /// </summary>
            public FolderSelectDialog()
            {
                ofd = new System.Windows.Forms.OpenFileDialog();

                ofd.Filter = "Folders|\n";
                ofd.AddExtension = false;
                ofd.CheckFileExists = false;
                ofd.DereferenceLinks = true;
                ofd.Multiselect = false;
            }

            #region Properties

            /// <summary>
            /// Gets/Sets the initial folder to be selected. A null value selects the current directory.
            /// </summary>
            public string InitialDirectory
            {
                get { return ofd.InitialDirectory; }
                set { ofd.InitialDirectory = value == null || value.Length == 0 ? Environment.CurrentDirectory : value; }
            }

            /// <summary>
            /// Gets/Sets the title to show in the dialog
            /// </summary>
            public string Title
            {
                get { return ofd.Title; }
                set { ofd.Title = value == null ? "Select a folder" : value; }
            }

            /// <summary>
            /// Gets the selected folder
            /// </summary>
            public string FileName
            {
                get { return ofd.FileName; }
            }

            #endregion

            #region Methods

            /// <summary>
            /// Shows the dialog
            /// </summary>
            /// <returns>True if the user presses OK else false</returns>
            public bool ShowDialog()
            {
                return ShowDialog(IntPtr.Zero);
            }

            public class WindowWrapper : System.Windows.Forms.IWin32Window
            {
                /// <summary>
                /// Constructor
                /// </summary>
                /// <param name="handle">Handle to wrap</param>
                public WindowWrapper(IntPtr handle)
                {
                    _hwnd = handle;
                }

                /// <summary>
                /// Original ptr
                /// </summary>
                public IntPtr Handle
                {
                    get { return _hwnd; }
                }

                private IntPtr _hwnd;
            }

            /// <summary>
            /// Shows the dialog
            /// </summary>
            /// <param name="hWndOwner">Handle of the control to be parent</param>
            /// <returns>True if the user presses OK else false</returns>
            public bool ShowDialog(IntPtr hWndOwner)
            {
                bool flag = false;

                if (Environment.OSVersion.Version.Major >= 6)
                {
                    var r = new Reflector("System.Windows.Forms");

                    uint num = 0;
                    Type typeIFileDialog = r.GetType("FileDialogNative.IFileDialog");
                    object dialog = r.Call(ofd, "CreateVistaDialog");
                    r.Call(ofd, "OnBeforeVistaDialog", dialog);

                    uint options = (uint)r.CallAs(typeof(System.Windows.Forms.FileDialog), ofd, "GetOptions");
                    options |= (uint)r.GetEnum("FileDialogNative.FOS", "FOS_PICKFOLDERS");
                    r.CallAs(typeIFileDialog, dialog, "SetOptions", options);

                    object pfde = r.New("FileDialog.VistaDialogEvents", ofd);
                    object[] parameters = new object[] { pfde, num };
                    r.CallAs2(typeIFileDialog, dialog, "Advise", parameters);
                    num = (uint)parameters[1];
                    try
                    {
                        int num2 = (int)r.CallAs(typeIFileDialog, dialog, "Show", hWndOwner);
                        flag = 0 == num2;
                    }
                    finally
                    {
                        r.CallAs(typeIFileDialog, dialog, "Unadvise", num);
                        GC.KeepAlive(pfde);
                    }
                }
                else
                {
                    var fbd = new FolderBrowserDialog();
                    fbd.Description = this.Title;
                    fbd.SelectedPath = this.InitialDirectory;
                    fbd.ShowNewFolderButton = false;
                    if (fbd.ShowDialog(new WindowWrapper(hWndOwner)) != DialogResult.OK) return false;
                    ofd.FileName = fbd.SelectedPath;
                    flag = true;
                }

                return flag;
            }

            #endregion
        }

        private void selectMappingTable_Click(object sender, EventArgs e)
        {
            // 获取当前目录下的指定文件路径
            //string currentDirectory = Directory.GetCurrentDirectory();
            //string specificFile = Path.Combine(currentDirectory, "FT映射表.xlsx");
            //mapTextBox.Text = specificFile;
            //mappingFilePath = specificFile;
            //bool flag = this.ofd.ShowDialog() == DialogResult.OK;
            //string fullPath = ofd.FileName;

            //if (flag)
            //{


            //    mappingTextBoxName = specificFile;
            //    mapTextBox.Text = Path.GetFileName(specificFile);

            //    if (Path.GetExtension(mappingTextBoxName).Contains("xlsx") || Path.GetExtension(mappingTextBoxName).Contains("xls"))
            //    {
            //        mappingFilePath = fullPath;
            //    }
            //    else
            //    {
            //        MessageBox.Show("請選擇正確的 xlsx 表格格式！");
            //        mapTextBox.Text = "";
            //    }
            //}
        }

        private void selectModelTable_Click(object sender, EventArgs e)
        {
            machineTextBox.Text ="";
            bool flag = this.ofd.ShowDialog() == DialogResult.OK;
            if (flag)
            {
                modelTableTextBoxName = ofd.FileName;
                machineTextBox.Text = Path.GetFileName(ofd.FileName);
                if (Path.GetExtension(modelTableTextBoxName).Contains("xlsx") || Path.GetExtension(modelTableTextBoxName).Contains("xls"))
                {
                  

                }
                else
                {
                    MessageBox.Show("請選擇正確的xlsx表格格式！");
                    machineTextBox.Text = "";
                }
            }


        }
        //private void savePath_Click(object sender, EventArgs e)
        //{
        //    // 获取桌面路径
        //    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //    savaPathTextBox.Text = desktopPath; 
        //}
        private void transform_click(object sender, EventArgs e)
        {
            if (savaPathTextBox.Text == "" || machineTextBox.Text==""|| mapTextBox.Text=="")
            {
                MessageBox.Show("請選擇对应的文件以及存放路径！");
                return;
            }
            Document document = new Document(this);
            document.createDocuments(mappingFilePath, modelTableTextBoxName, startupPath, savaPathTextBox.Text);
            MessageBox.Show("轉檔成功！！");
            savaPathTextBox.Text = "";
            machineTextBox.Text = "";
            mapTextBox.Text = "";

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}

