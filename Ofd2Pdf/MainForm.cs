using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Spire.Pdf.Conversion;

namespace Ofd2Pdf
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
        }

        List<OFDFile> fileList = new List<OFDFile>();

        Converter converter = new Converter();

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "OFD 文档(*.ofd)|*.ofd";
            ofd.Multiselect = true;
            ofd.CheckFileExists = true;
            var result = ofd.ShowDialog();
            if (result == DialogResult.OK)
            {
                foreach (var file in ofd.FileNames)
                {
                    OFDFile oFDFile = new OFDFile(file);
                    fileList.Add(oFDFile);
                }
                LoadFilesToListView();
            }
        }

        private Color ConvertColor(Status status)
        {
            switch (status)
            {
                case Status.等待转换:
                    return Color.Black;
                case Status.正在转换:
                    return Color.CadetBlue;
                case Status.转换完成:
                    return Color.LimeGreen;
                case Status.转换失败:
                    return Color.IndianRed;
                default:
                    return Color.Black;
            }
        }

        private void LoadFilesToListView()
        {
            listView1.BeginUpdate();
            listView1.Items.Clear();
            foreach (OFDFile oFDFile in fileList)
            {
                var item = new ListViewItem(oFDFile.FileName);
                item.UseItemStyleForSubItems = false;
                var subItem = new ListViewItem.ListViewSubItem();
                subItem.Text = oFDFile.Status.ToString();
                subItem.ForeColor = ConvertColor(oFDFile.Status);
                item.SubItems.Add(subItem);
                listView1.Items.Add(item);
            }
            listView1.EndUpdate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fileList.Clear();
            LoadFilesToListView();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (fileList.Count == 0)
            {
                MessageBox.Show("请先添加或拖拽文件！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Task.Run(() =>
            {
                bool didSth = false;
                for (int i = 0; i < fileList.Count; i++)
                {
                    if (fileList[i].Status == Status.转换完成)
                    {
                        continue;
                    }

                    // get the list item
                    var item = listView1.Items[i].SubItems[1];

                    // set converting
                    item.ForeColor = ConvertColor(Status.正在转换);
                    item.Text = Status.正在转换.ToString();
                    fileList[i].Status = Status.正在转换;

                    // convert to pdf
                    string PdfName = fileList[i].FileName.Substring(0, fileList[i].FileName.Length - 3) + "pdf";
                    ConvertResult result = converter.ConvertToPdf(fileList[i].FileName, PdfName);

                    if (result == ConvertResult.Failed)
                    {
                        item.ForeColor = ConvertColor(Status.转换失败);
                        item.Text = Status.转换失败.ToString();
                        fileList[i].Status = Status.转换失败;
                    }
                    else
                    {
                        item.ForeColor = ConvertColor(Status.转换完成);
                        item.Text = Status.转换完成.ToString();
                        fileList[i].Status = Status.转换完成;
                        didSth = true;
                    }
                }
                if (!didSth)
                {
                    MessageBox.Show("当前所有任务已经完成，请添加新的文件或清空重来。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            });
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in files)
            {
                OFDFile oFDFile = new OFDFile(file);
                fileList.Add(oFDFile);
            }
            LoadFilesToListView();
        }
    }
}
