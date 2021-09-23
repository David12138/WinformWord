using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace WinFormsApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //调用ReadLog向richTextBox1内写入日志
            ReadLog("日志窗口初始化成功");
        }

        /// <summary>
        /// 执行按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text.Trim();
            ReadLog(name);
            //MessageBox.Show(text);
        }

        public void ReadLog(string log)
        {
            string Time = Convert.ToString(DateTime.Now);
            richTextBox1.AppendText(Time + "  " + log + "\n");
        }

        /// <summary>
        /// 导出按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("请选择文件名！");
            }

            try
            {
                string path = System.Windows.Forms.Application.StartupPath + @"\Templates\" + name;
                if (File.Exists(path))
                {
                    #region 文件路由选择框
                    FolderBrowserDialog dirDialog = new FolderBrowserDialog();
                    dirDialog.ShowDialog(); 
                    #endregion

                    if (dirDialog.SelectedPath != string.Empty)
                    {
                        string newFileName = dirDialog.SelectedPath + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx";

                        Dictionary<string, string> wordLableList = new Dictionary<string, string>();
                        wordLableList.Add("年", "2021");
                        wordLableList.Add("月", "9");
                        wordLableList.Add("日", "18");
                        wordLableList.Add("星期", "六");
                        wordLableList.Add("标题", "Word导出数据");
                        wordLableList.Add("内容", "我是内容——Kiba518");
                        wordLableList.Add("姓名", "周大伟");

                        Export(path, newFileName, wordLableList);
                        MessageBox.Show("导出成功!");
                    }
                    else
                    {
                        MessageBox.Show("请选择导出位置");
                    }
                }
                else
                {
                    MessageBox.Show("Word模板文件不存在!");
                    ReadLog("Word模板文件不存在!");
                }
            }
            catch (Exception ex)
            {
                string res = "报错：" + ex.ToString();
                ReadLog(res);
                throw ex;
            }
        }

        /// <summary>
        /// 导出方法
        /// </summary>
        /// <param name="wordTemplatePath"></param>
        /// <param name="newFileName"></param>
        /// <param name="wordLableList"></param>
        public static void Export(string wordTemplatePath, string newFileName, Dictionary<string, string> wordLableList)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string TemplateFile = wordTemplatePath;
            File.Copy(TemplateFile, newFileName);
            _Document doc = new Document();
            object obj_NewFileName = newFileName;
            object obj_Visible = false;
            object obj_ReadOnly = false;
            object obj_missing = System.Reflection.Missing.Value;

            doc = app.Documents.Open(ref obj_NewFileName, ref obj_missing, ref obj_ReadOnly, ref obj_missing,
                ref obj_missing, ref obj_missing, ref obj_missing, ref obj_missing,
                ref obj_missing, ref obj_missing, ref obj_missing, ref obj_Visible,
                ref obj_missing, ref obj_missing, ref obj_missing,
                ref obj_missing);
            doc.Activate();

            if (wordLableList.Count > 0)
            {
                object what = WdGoToItem.wdGoToBookmark;
                foreach (var item in wordLableList)
                {
                    object lableName = item.Key;
                    if (doc.Bookmarks.Exists(item.Key))
                    {
                        doc.ActiveWindow.Selection.GoTo(ref what, ref obj_missing, ref obj_missing, ref lableName);//光标移动书签的位置
                        doc.ActiveWindow.Selection.TypeText(item.Value);//在书签处插入的内容 
                        doc.ActiveWindow.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//设置插入内容的Alignment
                    }
                }
            }

            object obj_IsSave = true;
            doc.Close(ref obj_IsSave, ref obj_missing, ref obj_missing);
        }
    }
}