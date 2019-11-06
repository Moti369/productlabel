using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using Seagull.BarTender.Print;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Configuration;

namespace DetProductionPrint
{
    public partial class Form1 : Form
    {
        string sqlcon = ConfigurationManager.ConnectionStrings["ApplicationServices"].ConnectionString;
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        public string inipath = System.IO.Directory.GetCurrentDirectory() + "\\LabelName.ini";


        /// <summary> 
        /// 构造方法 
        /// </summary> 
        /// <param name="INIPath">文件路径</param> 
        public void IniFiles(string INIPath)
        {
            inipath = INIPath;
        }

        public void IniFiles() { }

        /// <summary> 
        /// 写入INI文件 
        /// </summary> 
        /// <param name="Section">项目名称(如 [TypeName] )</param> 
        /// <param name="Key">键</param> 
        /// <param name="Value">值</param> 
        public void IniWriteValue(string Section, string Key, string Value)
        {
            WritePrivateProfileString(Section, Key, Value, this.inipath);
        }
        /// <summary> 
        /// 读出INI文件 
        /// </summary> 
        /// <param name="Section">项目名称(如 [TypeName] )</param> 
        /// <param name="Key">键</param> 
        public string IniReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(500);
            int i = GetPrivateProfileString(Section, Key, "", temp, 500, this.inipath);
            return temp.ToString();
        }
        /// <summary> 
        /// 验证文件是否存在 
        /// </summary> 
        /// <returns>布尔值</returns> 
        public bool ExistINIFile()
        {
            return File.Exists(inipath);
        }


        public string path = "";
        public Form1()
        {
            InitializeComponent();
            BindDataSource();

            comboBoxPrint.Items.Clear();
            entryPrinter.Items.Clear();

            List<String> listPrinters = LocalPrinter.GetLocalPrinters();
            foreach (string str in listPrinters)
            {
                //int index;
                comboBoxPrint.Items.Add(str);
                entryPrinter.Items.Add(str);
            }
            if (comboBoxPrint.Items.Count > 0)
            {
                comboBoxPrint.SelectedIndex = 0;
                entryPrinter.SelectedIndex = 0;
            }

            try
            {    //OR  (aspnet_Roles.RoleName = N'系统管理')不查询系统管理员，系统管理手动输入
                string sql = "SELECT   aspnet_Roles.RoleName, aspnet_Users.UserName  AS name FROM  aspnet_Roles INNER JOIN aspnet_UsersInRoles ON aspnet_Roles.RoleId = aspnet_UsersInRoles.RoleId INNER JOIN aspnet_Users ON aspnet_UsersInRoles.UserId = aspnet_Users.UserId WHERE   (aspnet_Roles.RoleName = N'芯片仓库') ";
                SqlConnection conn = new SqlConnection(sqlcon);
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader sdr = cmd.ExecuteReader();
                while (sdr.Read())
                {
                    productUser.Items.Add(sdr["name"].ToString().Trim());
                }
                sdr.Close();
                conn.Close(); //关闭数据库连接
                productUser.SelectedIndex = 0;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error:" + ex.Message, "Error");
            }
        }


        string command = null;
        SqlConnection connect = null;
        SqlDataAdapter da = null;
        DataSet ds = null;
        private void BindDataSource()
        {
            command = "select LabelProductDataId as '序号', ProductName as '品名', ProductLotNo as '批号', ProductProcess as '流程', ProductConformance as '良品数', ProductDefective as '不良品数',UserName as '操作者', WriteDate as '录入时间' from[LabelProductData]";
            connect = new SqlConnection(sqlcon);
            connect.Open();
            da = new SqlDataAdapter(command, connect);
            ds = new DataSet();
            ds.Clear();
            da.Fill(ds, "LabelProductData");
            dataGridView1.DataSource = ds.Tables["LabelProductData"];
          
            //labelProductData.DataSource = ds;
            //labelProductData.DataMember = "LabelProductData";
            connect.Close();

        }

        public class LocalPrinter
        {
            private static PrintDocument fPrintDocument = new PrintDocument();
            /// <summary>
            /// 获取本机默认打印机名称
            /// </summary>
            public static String DefaultPrinter
            {
                get { return fPrintDocument.PrinterSettings.PrinterName; }
            }
            /// <summary>
            /// 获取本机的打印机列表。列表中的第一项就是默认打印机。
            /// </summary>
            public static List<String> GetLocalPrinters()
            {
                List<String> fPrinters = new List<string>();
                fPrinters.Add(DefaultPrinter); // 默认打印机始终出现在列表的第一项
                foreach (String fPrinterName in PrinterSettings.InstalledPrinters)
                {
                    if (!fPrinters.Contains(fPrinterName))
                        fPrinters.Add(fPrinterName);
                }
                return fPrinters;
            }
        }

        private void btPrint_Click(object sender, EventArgs e)
        {
          
            if (tbName.Text.Trim().Length<1)
            {
                MessageBox.Show("请输入打印品名！");
                return;
            }
            if (tbLotNo.Text.Trim().Length < 1)
            {
                MessageBox.Show("请输入打印批号！");
                return;
            }
            if (tbCount.Text.Trim().Length < 1)
            {
                MessageBox.Show("请输入打印数量！");
                return;
            }
            if (tbDate.Text.Trim().Length < 1)
            {
                MessageBox.Show("请输入打印日期！");
                return;
            }
            if ( tbTechnology.Text.Trim().Length < 1)
            {
                MessageBox.Show("请输入打印工艺！");
                return;
            }
          
            if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked && !checkBox4.Checked && !checkBox5.Checked && !checkBox6.Checked)
            {
                MessageBox.Show("请选择流程！");
                return;
            }
            PrintLabel();
            //打印完之后清除文本框内容
        }

        private void PrintLabel()
        {
            try
            {
                path = "\\btw\\DET随件单.btw";
                List<string> ProcessList = new List<string>();

                if (checkBox1.Checked)
                {
                    string strProcess = "分割";
                    ProcessList.Add(strProcess);
                }
                if (checkBox2.Checked)
                {
                    string strProcess = "初始化";
                    ProcessList.Add(strProcess);
                }
                if (checkBox3.Checked)
                {
                    string strProcess = "外检";
                    ProcessList.Add(strProcess);
                }
                if (checkBox4.Checked)
                {
                    string strProcess = "注塑";
                    ProcessList.Add(strProcess);
                }
                if (checkBox5.Checked)
                {
                    string strProcess = "成测";
                    ProcessList.Add(strProcess);
                }
                if (checkBox6.Checked)
                {
                    string strProcess = "完工检";
                    ProcessList.Add(strProcess);
                }

                Engine btEngine = new Engine();
                btEngine.Start();
                LabelFormatDocument btFormat = btEngine.Documents.Open(System.IO.Directory.GetCurrentDirectory() + path);//这里是Bartender软件生成的模板文件，你需要先把模板文件做好。
                                                                                                                         //LabelFormatDocument btFormat = System.IO.Directory.GetCurrentDirectory() + "\\btw\\LabelName.ini";
                btFormat.PrintSetup.PrinterName = comboBoxPrint.Text;
                btFormat.PrintSetup.IdenticalCopiesOfLabel = 1; //打印份数
                                                                //int count = Int32.Parse(printCount.Text);
                SubStrings substring = btFormat.SubStrings;
                for (int j = 0; j < ProcessList.Count; j++)
                {
                    for (int i = 0; i < substring.Count; i++)
                    {
                        if (substring[i].Name.ToUpper() == "P1")
                            substring[i].Value = ProcessList[j];
                        if (substring[i].Name.ToUpper() == "P2")
                            substring[i].Value = tbTechnology.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P3")
                            substring[i].Value = tbName.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P4")
                            substring[i].Value = tbLotNo.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P5")
                            substring[i].Value = tbCount.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P6")
                            substring[i].Value = tbDate.Text.Trim();

                    }
                    //Thread.Sleep(10);
                    Result nResult = btFormat.Print();
                    btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                    btFormat.Close(SaveOptions.DoNotSaveChanges);//不保存对打开模板的修改
                }
               
            }
            catch (Exception ex)
            {

                throw new NotImplementedException(); ;
            }
           
            //throw new NotImplementedException();
        }

        private void EntryPrint_Click(object sender, EventArgs e)
        {
            //当点击打印后将数量写入本地文件
            int num = 0;
            int numBox = 0;
            string strMiddleCount = middleCount.Text.Trim();
            string strBoxCount = boxCount.Text.Trim();

            inipath = System.IO.Directory.GetCurrentDirectory() + "\\LabelName.ini";
            List<string> countList = new List<string>();
            List<string> countBoxList = new List<string>();
            if (ExistINIFile())
            {
                string strMiddle = IniReadValue("Test", "label_numMiddle");
                string strBox = IniReadValue("Test", "label_numBox");

                num = strMiddle.Length > 0 ? Int32.Parse(strMiddle) : 0;
                numBox = strBox.Length > 0 ? Int32.Parse(strBox) : 0;

                countList.Clear();
                countBoxList.Clear();
                for (int i = 0; i < num; i++)
                {
                    countList.Add(IniReadValue("Test", "label_countMiddle" + i.ToString()));
                }
                for (int i = 0; i < numBox; i++)
                {
                   
                    countBoxList.Add(IniReadValue("Test", "label_countBox" + i.ToString()));
                    
                }
            }
            if (strMiddleCount.Length<1)
            {
                MessageBox.Show("请输入中包装数量！");
                return;
            }
            else
            {
                if (!countList.Contains(strMiddleCount))
                {
                    IniWriteValue("Test", "label_countMiddle" + num, strMiddleCount);
                    num++;
                    IniWriteValue("Test", "label_numMiddle", num.ToString());
                }
              
            }
            if (strBoxCount.Length < 1)
            {
                //此处若选择中包装的话一箱中包数可不填写
                MessageBox.Show("请输入一箱中包数！");
                return;
            }
            else
            {
                if (!countBoxList.Contains(strBoxCount))
                {
                    IniWriteValue("Test", "label_countBox" + numBox, strBoxCount);
                    numBox++;
                    IniWriteValue("Test", "label_numBox", numBox.ToString());
                }
               
            }

            //打印完之后清楚文本框内容 点击打印后将数量保存到本地，下次读取
            if (entryCount.Text.Trim().Length < 1)
            {
                MessageBox.Show("请输入入库标签打印数量！");
                return;
            }
            if (middleCount.Text.Trim().Length<1)
            {
                MessageBox.Show("请输入中包装数量！");
                return;
            }
            if (boxCount.Text.Trim().Length < 1)
            {
                MessageBox.Show("请输入一箱中包数！");
                return;
            }
            PrintEntryLabel();
           
        }

        private void PrintEntryLabelPublic(string path, string pathOut, string temp) {

            bool signTrue = true;
            try
            {
                Engine btEngine = new Engine();
                btEngine.Start();
               
                int pageMiddle = 1;
                int pageOut = 1;
                // 中包装
                if (checkBox10.Checked)
                {
                    //中外标签不一样的话
                    LabelFormatDocument btFormat = btEngine.Documents.Open(System.IO.Directory.GetCurrentDirectory() + path);//这里是Bartender软件生成的模板文件，你需要先把模板文件做好。
                    btFormat.PrintSetup.PrinterName = entryPrinter.Text;
                    SubStrings substring = btFormat.SubStrings;
                    boxCount.Enabled = false;
                    int remainder = int.Parse(entryCount.Text.Trim()) % int.Parse(middleCount.Text.Trim());
                    pageMiddle = int.Parse(entryCount.Text.Trim()) / int.Parse(middleCount.Text.Trim());
                    if (remainder != 0)
                    {
                        pageMiddle = pageMiddle + 1;
                    }
                    btFormat.PrintSetup.IdenticalCopiesOfLabel = pageMiddle; //打印份数
                    for (int i = 0; i < substring.Count; i++)
                    {
                        if (substring[i].Name.ToUpper() == "P1")
                            substring[i].Value = entryName.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P2")
                            substring[i].Value = entryLotNo.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P3")
                            substring[i].Value = middleCount.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P4")
                            substring[i].Value = entryDate.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P5")
                            substring[i].Value = temp;

                    }
                    Result nResult = btFormat.Print();
                    btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                    btFormat.Close(SaveOptions.DoNotSaveChanges);
                }
                if (checkBox11.Checked)
                {
                    LabelFormatDocument btFormat = btEngine.Documents.Open(System.IO.Directory.GetCurrentDirectory() + pathOut);//这里是Bartender软件生成的模板文件，你需要先把模板文件做好。
                    btFormat.PrintSetup.PrinterName = entryPrinter.Text;
                    SubStrings substring = btFormat.SubStrings;
                    boxCount.Enabled = true;
                    int remainder = int.Parse(entryCount.Text.Trim()) % (int.Parse(middleCount.Text.Trim()) * int.Parse(boxCount.Text.Trim()));
                    pageOut = int.Parse(entryCount.Text.Trim()) / int.Parse(middleCount.Text.Trim()) / int.Parse(boxCount.Text.Trim());
                    if (remainder != 0)
                    {
                        pageOut = pageOut + 1;
                    }
                    btFormat.PrintSetup.IdenticalCopiesOfLabel = pageOut; //打印份数
                    for (int i = 0; i < substring.Count; i++)
                    {
                        if (substring[i].Name.ToUpper() == "P1")
                            substring[i].Value = entryName.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P2")
                            substring[i].Value = entryLotNo.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P3")
                            substring[i].Value = entryCount.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P4")
                            substring[i].Value = entryDate.Text.Trim();
                        if (substring[i].Name.ToUpper() == "P5")
                            substring[i].Value = temp;

                    }
                    Result nResult = btFormat.Print();
                    btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                    btFormat.Close(SaveOptions.DoNotSaveChanges);
                }

            }
            catch (Exception ex)
            {
                signTrue = false;
                MessageBox.Show("出错："+ex.Message);
            }
        }

        private void PrintEntryLabel()
        {
            try
            {
                //int count = Int32.Parse(middleCount.Text.Trim()) * Int32.Parse(boxCount.Text.Trim());
                
               // List<string> ProcessList = new List<string>();
                Engine btEngine;
                LabelFormatDocument btFormat;
                int pageMiddle = 1;
                int pageOut = 1;
                SubStrings substring;
                if (rdStandard.Checked)
                {
                    // 中包装
                    if (checkBox10.Checked)
                    {
                        path = "\\btw\\DET入库标签.btw";
                        btEngine = new Engine();
                        btEngine.Start();
                        btFormat = btEngine.Documents.Open(System.IO.Directory.GetCurrentDirectory() + path);//这里是Bartender软件生成的模板文件，你需要先把模板文件做好。
                        btFormat.PrintSetup.PrinterName = entryPrinter.Text;
                        substring = btFormat.SubStrings;
                        //中外标签不一样的话
                        boxCount.Enabled = false;
                        int remainder = int.Parse(entryCount.Text.Trim()) % int.Parse(middleCount.Text.Trim());
                        pageMiddle = int.Parse(entryCount.Text.Trim()) / int.Parse(middleCount.Text.Trim());
                        if (remainder != 0)
                        {
                            pageMiddle = pageMiddle+1;
                        }
                        
                        btFormat.PrintSetup.IdenticalCopiesOfLabel = pageMiddle; //打印份数
                        for (int i = 0; i < substring.Count; i++)
                        {
                            if (substring[i].Name.ToUpper() == "P1")
                                substring[i].Value = entryName.Text.Trim();
                            if (substring[i].Name.ToUpper() == "P2")
                                substring[i].Value = entryLotNo.Text.Trim();
                            if (substring[i].Name.ToUpper() == "P3")
                                substring[i].Value = middleCount.Text.Trim();
                            if (substring[i].Name.ToUpper() == "P4")
                                substring[i].Value = entryDate.Text.Trim();

                        }
                        Result nResult = btFormat.Print();
                        btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                    }
                    if (checkBox11.Checked)
                    {
                        path = "\\btw\\DET入库标签－外箱.btw";
                        btEngine = new Engine();
                        btEngine.Start();
                        btFormat = btEngine.Documents.Open(System.IO.Directory.GetCurrentDirectory() + path);//这里是Bartender软件生成的模板文件，你需要先把模板文件做好。
                        btFormat.PrintSetup.PrinterName = entryPrinter.Text;
                        substring = btFormat.SubStrings;
                        boxCount.Enabled = true;
                        int remainder = int.Parse(entryCount.Text.Trim()) % (int.Parse(middleCount.Text.Trim()) * int.Parse(boxCount.Text.Trim()));
                        pageOut = int.Parse(entryCount.Text.Trim()) / int.Parse(middleCount.Text.Trim()) / int.Parse(boxCount.Text.Trim());
                        if (remainder != 0)
                        {
                            pageOut = pageOut + 1;
                        }
                        btFormat.PrintSetup.IdenticalCopiesOfLabel = pageOut; //打印份数
                        for (int i = 0; i < substring.Count; i++)
                        {
                            if (substring[i].Name.ToUpper() == "P1")
                                substring[i].Value = entryName.Text.Trim();
                            if (substring[i].Name.ToUpper() == "P2")
                                substring[i].Value = entryLotNo.Text.Trim();
                            if (substring[i].Name.ToUpper() == "P3")
                                substring[i].Value = entryCount.Text.Trim();
                            if (substring[i].Name.ToUpper() == "P4")
                                substring[i].Value = entryDate.Text.Trim();

                        }
                        Result nResult = btFormat.Print();
                        btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                    }
                        
                   
                }
                
                
                if (rdSteel.Checked)
                {
                    string filePath = "\\btw\\DET入库标签－钢带.btw";
                    string pathOut = "\\btw\\DET入库标签－钢带－外箱.btw";
                    try
                    {
                        PrintEntryLabelPublic(filePath, pathOut, "钢带");
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                   
           
                }
                if (rdSteelInjection.Checked)
                {
                    string filePath = "\\btw\\DET入库标签－钢带.btw";
                    string pathOut = "\\btw\\DET入库标签－钢带－外箱.btw";
                    try
                    {
                        PrintEntryLabelPublic(filePath, pathOut, "钢带注塑");
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
               
                //btEngine = new Engine();
                //btEngine.Start();
                //btFormat = btEngine.Documents.Open(System.IO.Directory.GetCurrentDirectory() + filePath);//这里是Bartender软件生成的模板文件，你需要先把模板文件做好。
                //                                                                                                            //LabelFormatDocument btFormat = System.IO.Directory.GetCurrentDirectory() + "\\btw\\LabelName.ini";
                //page = int.Parse(entryCount.Text.Trim());
                //btFormat.PrintSetup.PrinterName = comboBoxPrint.Text;
                //btFormat.PrintSetup.IdenticalCopiesOfLabel = page; //打印份数
                //                                                    //int count = Int32.Parse(printCount.Text);int.Parse(entryCount.Text.Trim())
                //substring = btFormat.SubStrings;
                //for (int j = 0; j < (ProcessList.Count); j++)
                //{
                //    for (int i = 0; i < substring.Count; i++)
                //    {
                //        if (substring[i].Name.ToUpper() == "P1")
                //            substring[i].Value = entryName.Text.Trim();
                //        if (substring[i].Name.ToUpper() == "P2")
                //            substring[i].Value = entryLotNo.Text.Trim();
                //        if (substring[i].Name.ToUpper() == "P3")
                //            substring[i].Value = count.ToString();
                //        if (substring[i].Name.ToUpper() == "P4")
                //            substring[i].Value = entryDate.Text.Trim();
                //        if (substring[i].Name.ToUpper() == "P5")
                //            substring[i].Value = ProcessList[j];

                //    }
                //    //Thread.Sleep(10);
                //    Result nResult = btFormat.Print();
                //    btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;
                //}
                
               
            }
            catch (Exception ex)
            {

                throw;
            }
         
            //throw new NotImplementedException();
        }

        private void BtnClean_Click(object sender, EventArgs e)
        {
            textBoxOUT.Text = "";
            textBoxOUT.Focus();
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex==1)
            {
                //textBoxOUT.Text = "";
                textBoxOUT.Focus();
                if (ExistINIFile())
                {
                    string strMiddle = IniReadValue("Test", "label_numMiddle");
                    string strBox = IniReadValue("Test", "label_numBox");

                    int num = strMiddle.Length > 0 ? Int32.Parse(strMiddle) : 0;
                    int numBox = strBox.Length > 0 ? Int32.Parse(strBox) : 0;


                    middleCount.Items.Clear();
                    boxCount.Items.Clear();
                    for (int i = 0; i < num; i++)
                    {
                        middleCount.Items.Add(IniReadValue("Test", "label_countMiddle" + i.ToString()));
                    }
                    for (int i = 0; i < numBox; i++)
                    {
                        boxCount.Items.Add(IniReadValue("Test", "label_countBox" + i.ToString()));
                    }
                    if (middleCount.Items.Count > 0)
                    {
                        middleCount.SelectedIndex = 0;
                    }
                    if (boxCount.Items.Count > 0)
                    {
                        boxCount.SelectedIndex = 0;
                    }
                }
            }
        }

        private void TextBoxOUT_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string strcode = textBoxOUT.Text.Trim().ToUpper();
                string[] sArray = strcode.Split(' ');
                try
                {
                    if (textBoxOUT.Text.Length < 1)
                    {
                        MessageBox.Show("请先扫描二维码");
                    }
                    else
                    {
                        entryName.Text = sArray[2].ToString();
                        entryLotNo.Text = sArray[3].ToString();
                        string[] date = sArray[4].ToString().Split('\\');
                        //DateTime datetime = DateTime.Parse(date[0] + "/" + date[2] + "/" + date[4]);
                        entryDate.Text = date[0] + "/" + date[2] + "/" + date[4];
                    }
                }
                catch (Exception)
                {
                    //////////////当前出错：未有合适模板//////
                    throw;
                }
            }
        }

        private void ProductCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string strcode = textBoxOUT.Text.Trim().ToUpper();
                string[] sArray = strcode.Split(' ');
                try
                {
                    if (textBoxOUT.Text.Length < 1)
                    {
                        MessageBox.Show("请先扫描二维码");
                    }
                    else
                    {
                        productName.Text = sArray[2].ToString();
                        productLotNo.Text = sArray[3].ToString();
                        productDate.Text = sArray[4].ToString();
                        productProcess.Text = sArray[0].ToString();
                    }
                }
                catch (Exception)
                {
                    //////////////当前出错：未有合适模板//////
                    throw;
                }
            }
        }

        private void ProductEntry_Click(object sender, EventArgs e)
        {
            if (productConformance.Text.Length < 1)
            {
                MessageBox.Show("请输入良品数");
                return;
            }
            if (productDefective.Text.Length < 1)
            {
                MessageBox.Show("请输入不良品数");
                return;
            }
            if (productUser.Text.Length<1)
            {
                MessageBox.Show("请输入操作者");
                return;
            }
            string[] date = productDate.Text.Split('\\');
            DateTime datetime = DateTime.Parse(date[0]+"-"+date[2]+"-"+date[4]);
            string strSql = "select * from [LabelProductData] where ProductName='"+ productName.Text + "' and ProductLotNo='"+ productLotNo.Text + "' and ProductProcess='"+ productProcess.Text + "' and ProductDate='"+ datetime.ToShortDateString() + "' and ProductConformance= '" + productConformance.Text + "' and ProductDefective='"+ productDefective.Text + "' and UserName='"+ productUser.Text + "'";
            SqlConnection con = new SqlConnection(sqlcon);
            con.Open();
            SqlCommand cmd = new SqlCommand(strSql, con);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                con.Close();
                MessageBox.Show("该条数据已存在");
                return;
            }

            string strsql = "insert into [LabelProductData] (ProductName,ProductLotNo,ProductDate,ProductProcess,ProductConformance,ProductDefective,UserName,WriteDate) Values('" + productName.Text + "','" + productLotNo.Text + "','" + datetime.ToShortDateString() + "','" + productProcess.Text + "','" + productConformance.Text + "','" + productDefective.Text + "','" + productUser.Text + "','" + DateTime.Now.ToString() + "')";

            //SqlDataReader dr = null;
            //dr = cmd.ExecuteReader();
            dr.Close();
            try
            {
                cmd.CommandText = strsql;
                cmd.ExecuteNonQuery();
                BindDataSource();
                MessageBox.Show("插入成功");
            }
            catch (Exception)
            {
                MessageBox.Show("插入失败");
                throw;
            }
            productConformance.Text = "";
            productDefective.Text = "";

        }

        private void Insert()
        {
            throw new NotImplementedException();
        }

        private void ProductClear_Click(object sender, EventArgs e)
        {
            productCode.Clear();
            productCode.Focus();
        }

        private void Edit_Click_1(object sender, EventArgs e)
        {
            try
            {
                SqlConnection con = new SqlConnection(sqlcon);
                con.Open();
                string sql = "update LabelProductData set ProductConformance='" + productConformance.Text + "',ProductDefective='" + productDefective.Text + "' where LabelProductDataId=" + Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString()); ;
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandText = sql;
                int x = cmd.ExecuteNonQuery();
                if (x == 1)
                {
                    //如果添加成功，那么给用户提示一下 
                    MessageBox.Show("修改成功");
                    productConformance.Text = "";
                    productDefective.Text = "";
                    BindDataSource();
                }

            }
            catch (Exception)
            {

                throw;
            }
            //Form2 form2 = new Form2();
            ////this.Hide();     //隐藏当前窗体    
            
            //form2.ShowDialog();
        }

        private void Delete_Click_1(object sender, EventArgs e)
        {
            string sql = string.Format("delete from LabelProductData where LabelProductDataId = '{0}'", dataGridView1.SelectedRows[0].Cells[0].Value);
            SqlConnection con = new SqlConnection(sqlcon);
            con.Open();
            SqlCommand cmd = new SqlCommand(sql, con);

            cmd.ExecuteNonQuery(); //更新数据库  
            BindDataSource(); 
            con.Close();
            //关闭数据训            
            MessageBox.Show("记录已从数据库删除，请按刷新按钮刷新显示列表", "友情提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
       
        }

        private void ProductClear_Click_1(object sender, EventArgs e)
        {
            productCode.Clear();
            productCode.Focus();
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                productConformance.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                productDefective.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message+"请勿选择表头");
            }
            
        }
    }
}
