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

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "OFD 文档(*.ofd)|*.ofd";
            ofd.Multiselect = true;
            ofd.CheckFileExists = true;
            var result = ofd.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileList.Clear();
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
            Task.Run(() =>
            {
                for (int i = 0; i < fileList.Count; i++)
                {
                    var item = listView1.Items[i].SubItems[1];
                    OfdConverter converter = new OfdConverter(fileList[i].FileName);
                    try
                    {
                        item.ForeColor = ConvertColor(Status.正在转换);
                        item.Text = Status.正在转换.ToString();
                        string PdfName = fileList[i].FileName.Substring(0, fileList[i].FileName.Length - 3) + "pdf";
                        converter.ToPdf(PdfName);
                    }
                    catch (Exception)
                    {
                        item.ForeColor = ConvertColor(Status.转换失败);
                        item.Text = Status.转换失败.ToString();
                    }
                    item.ForeColor = ConvertColor(Status.转换完成);
                    item.Text = Status.转换完成.ToString();
                }
            });
        }
    }
}
