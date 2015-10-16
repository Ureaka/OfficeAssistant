using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;

using System.Runtime.InteropServices;

namespace walcl
{
    public partial class Form1 : Form
    {
        int intRows = 0;
        int intCols = 0;
        String[,] strData;
        String WorkDirPath;
        const String DefaultDirPath = @"D:\O";
        const String TEMPLETNAME = @"Templet.dot";
        const String DATAFILENAME = @"MyExcel.xls";

        public Form1()
        {
            InitializeComponent();
        }

        private void btExec_Click(object sender, EventArgs e)
        {
            if (LoadFromExcel())
            {
                WriteDocument();
            }
            strData = null;
        }

        public void WriteDocument()
        {
            MSWord.ApplicationClass wordApp = new MSWord.ApplicationClass();
            MSWord.Document wordDoc = null;
            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object isVisible = true;
            object TempletPath = WorkDirPath + @"\" + TEMPLETNAME;
            try
            {
                //Generate all documents, start from the 2nd line
                for (int i = 2; i <= intRows; i++)
                {
                    Log("Create document for " + strData[i - 1, 0] + "...");
                    wordDoc = wordApp.Documents.Open(ref TempletPath, ref missing, ref readOnly,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                    for (int j = 1; j <= intCols; j++)
                    {
                        string[] sArray = strData[0, j - 1].Split('|');
                        foreach (string sBmk in sArray)
                        {
                            object bkObj = sBmk;
                            if (sArray.Length != 1) Log("Combined: " + strData[0, j - 1] + " contains " + sArray.Length + " bootmark , Finding  -- " + sBmk);
                            if (wordApp.ActiveDocument.Bookmarks.Exists(sBmk))
                            {
                                wordApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
                                String s = strData[i - 1, j - 1];
                                int len = Encoding.Default.GetBytes(s).Length;
                                Log(s + " length: " + s.Length + ", bytes:" + len);
                                wordApp.Selection.Delete(Type.Missing, len);
                                wordApp.Selection.TypeText(strData[i - 1, j - 1]);
                            }
                        }
                    }
                    object sDocPath = WorkDirPath + @"\" + strData[i - 1, 0] + ".doc";
                    //object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;
                    object format = MSWord.WdSaveFormat.wdFormatDocument97;
                    wordDoc.SaveAs(ref sDocPath, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    wordDoc.Close(ref missing, ref missing, ref missing);
                }
                wordApp.Quit(ref missing, ref missing, ref missing);
                Log("Create OK");
            }
            catch (Exception e)
            {
                Log("Excel App error:" + e.ToString());
                MessageBox.Show("WordApp App error:" + e.ToString());
            }
            finally
            {
                if (wordApp != null)
                {
                    Kill(wordApp);
                    Log("Kill WordApp" + Environment.NewLine);
                }
            }

        }

        [DllImport("user32.dll")]
        public static extern System.IntPtr FindWindowEx(System.IntPtr parent, System.IntPtr childe, string strclass, string strname);
        public void Kill(MSWord.Application word)
        {

            IntPtr p = FindWindowEx(System.IntPtr.Zero, System.IntPtr.Zero, null, word.Caption);
            int k = 0;

            GetWindowThreadProcessId(p, out k);   //得到本进程唯一标志k  
            if (k != 0)
            {
                System.Diagnostics.Process fp = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用  
                fp.Kill();
            }
        }

        public Boolean LoadFromExcel()
        {
            MSExcel.ApplicationClass excelApp = new MSExcel.ApplicationClass();
            MSExcel._Workbook workBook = null;
            MSExcel._Worksheet workSheet = null;
            MSExcel.Range rng = null; //Ref https://msdn.microsoft.com/zh-cn/library/microsoft.office.interop.excel.range_members.aspx
            object missing = System.Reflection.Missing.Value;
            object readOnly = true;
            try
            {
                // 打开Excel文件 
                workBook = excelApp.Workbooks.Open(WorkDirPath + @"\" + DATAFILENAME, missing, readOnly, missing,
                    missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing);
                //excelApp.Visible = true; // Excel应用程序可见 
                // 针对Excel文档中的第一个Sheet就行操作。Sheet的下标从1开始 
                workSheet = (MSExcel._Worksheet)workBook.Sheets[1];
                workSheet.Activate();

                //取得总记录行数 (包括标题列)
                int N = 100;
                int i, j;
                for (i = 1; i <= N; i++)
                {
                    if (((MSExcel.Range)workSheet.Cells[1, i]).Value2 == null)
                    {
                        break;
                    }
                }
                intCols = i - 1;
                for (i = 1; i <= N; i++)
                {
                    if (((MSExcel.Range)workSheet.Cells[i, 1]).Value2 == null)
                    {
                        break;
                    }
                }
                intRows = i - 1;
                Log("Rows " + intRows.ToString() + ", Cols " + intCols.ToString());
                //Log("System: " + workSheet.UsedRange.Cells.Rows.Count + ", " + workSheet.UsedRange.Cells.Columns.Count);

                //Read whole table, Ref http://blog.xudan123.com/121.html
                strData = new String[intRows, intCols];
                for (i = 1; i <= intRows; i++)
                {
                    for (j = 1; j <= intCols; j++)
                    {
                        rng = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[i, j];
                        if (rng.Value2 == null)
                        {
                            Log("Row " + i + ", Col " + j + " Value is " + String.Empty);
                            if (1 == j)
                            {
                                MessageBox.Show("Row " + i + ", Col " + j + ", " + strData[0, j - 1] + " must not be Null!");
                                return false;
                            }
                            DialogResult dr = DialogResult.Yes;//MessageBox.Show(strData[i - 1, 0] + "'s " + strData[0, j - 1] + " (Row " + i + ", Col " + j + ") is Null! Continue?", "Warning", MessageBoxButtons.YesNo);
                            if (dr == DialogResult.Yes)
                            {
                                strData[i - 1, j - 1] = String.Empty;
                                continue;
                            }
                            else if (dr == DialogResult.No)
                            {
                                return false;
                            }
                        }
                        //Use rng.Text other than rng.Value2 to Get the special TEXT
                        strData[i - 1, j - 1] = rng.Text.ToString().Trim();
                        Log("Row " + i + ", Col " + j + " Text is " + strData[i - 1, j - 1]);
                    }
                }

                // 保存并关闭Excel工作簿 
                workBook.Close(false, missing, missing);
                // 退出Excel程序。如果不退出，可能出现，没有Excel文件打开，但是任务管理器中还有Excel进程 
                excelApp.Quit();
            }
            catch (Exception e)
            {
                Log("Excel App error:" + e.ToString());
                MessageBox.Show("Excel App error:" + e.ToString());
            }
            finally
            {
                if (excelApp != null)
                {
                    Kill(excelApp);
                    Log("Kill excel app" + Environment.NewLine);
                }
            }
            return true;
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(MSExcel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口 
            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
            p.Kill();     //关闭进程k
        }

        public void Log(String s)
        {
            tbInfo.AppendText(s + Environment.NewLine);
            tbInfo.ScrollToCaret();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tbFolder.Text = DefaultDirPath;
            WorkDirPath = tbFolder.Text;
        }

        private void btFolder_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderBrowserDialog.SelectedPath = tbFolder.Text;
            System.Windows.Forms.DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                tbFolder.Text = folderBrowserDialog.SelectedPath;
                WorkDirPath = tbFolder.Text;
            }
        }

        private void tbFolder_DragDrop(object sender, DragEventArgs e)
        {
            String path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            if (Directory.Exists(path))
            {
                tbFolder.Text = path;
                WorkDirPath = tbFolder.Text;
            }
            else if (File.Exists(path))
            {
                path = System.IO.Path.GetDirectoryName(path);
                tbFolder.Text = path;
                WorkDirPath = tbFolder.Text;
            }
        }

        private void tbFolder_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
    }
}
