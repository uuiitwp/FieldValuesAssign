using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geoprocessor;
using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace FieldValuesAssign
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();
            eh = new ExcelHelper();
            int listBoxCount = eh.getListBoxCount();
            KeyValuePair<int, string>[] fieldNames = eh.getFieldNames().ToArray<KeyValuePair<int, string>>();
            IEnumerable<string> q = (from row in fieldNames
                                     select row.Value).Distinct<string>();
            if (q.Count<string>() != fieldNames.Count<KeyValuePair<int, string>>())
            {
                MessageBox.Show("存在指定字段重复");
                button1_Click(null, null);
                return;
            }
            for (int i = 0; i < listBoxCount; i++)
            {
                ListBox lb = new ListBox
                {
                    Location = new Point(10 + i * 126, 60)
                };
                lb.SelectedIndexChanged += listBox_SelectedIndexChanged;
                lb.HorizontalScrollbar = true;
                Label j = new Label();
                lb.Name = (j.Text = fieldNames[i].Value);
                j.Location = new Point(10 + i * 126, 35);
                listBoxs.Add(lb);
                Labels.Add(j);
                base.Controls.Add(lb);
                base.Controls.Add(j);
            }
            resize(null, null);
            reload();
        }
        private readonly List<ListBox> listBoxs = new List<ListBox>();
        private readonly List<Label> Labels = new List<Label>();
        private readonly ExcelHelper eh;
        private void reload()
        {
            using (List<ListBox>.Enumerator enumerator = listBoxs.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ListBox lb = enumerator.Current;
                    lb.BeginUpdate();
                    string[] q = (from d in eh.data
                                  select d[lb.Name].ToString().Trim()).Distinct<string>().ToArray<string>();
                    lb.Items.Clear();
                    lb.ClearSelected();
                    lb.Items.AddRange(q);
                    lb.EndUpdate();
                }
            }
            updateTip();
        }
        private List<ILayer> getSelectedLayers()
        {
            List<ILayer> r = new List<ILayer>();
            for (int i = 0; i < ArcMap.Document.FocusMap.LayerCount; i++)
            {
                ILayer lyr = ArcMap.Document.FocusMap.get_Layer(i);
                if ((lyr as IFeatureSelection).SelectionSet.Count != 0)
                {
                    r.Add(lyr);
                }
            }
            return r;
        }

        private void resize(object sender, EventArgs e)
        {
            textBox1.Size = new Size(base.Size.Width - 150, textBox1.Size.Height);
            int listBoxCount = listBoxs.Count;
            for (int i = 0; i < listBoxCount; i++)
            {
                Label l = Labels[i];
                l.Location = new Point(i * ((Width - 10) / listBoxCount) + 10, l.Location.Y);
                ListBox lb = listBoxs[i];
                lb.Location = new Point(i * ((Width - 10) / listBoxCount) + 10, lb.Location.Y);
                lb.Width = (Width - 20 - listBoxCount * 20) / listBoxCount;
                lb.Height = Height - 200;
            }
        }

        public List<egi> getFieldsandValues()
        {
            if (listBoxs.All((ListBox x) => x.Items.Count == 1))
            {
                Dictionary<int, string[]> f2f = eh.getFieldtoField();
                Dictionary<int, string> fns = eh.getFieldNames();
                IEnumerable<egi> q = from t1 in fns
                                     join t2 in f2f on t1.Key equals t2.Key
                                     join t3 in listBoxs on t1.Value equals t3.Name
                                     select new egi
                                     {
                                         excelfield = t1.Value,
                                         gisfield = t2.Value,
                                         input = t3.Items[0].ToString()
                                     };
                return q.ToList<egi>();
            }
            return new List<egi>();
        }

        private void updateTip()
        {
            if (listBoxs.All((ListBox x) => x.Items.Count == 1))
            {
                Dictionary<int, string[]> f2f = eh.getFieldtoField();
                Dictionary<int, string> fns = eh.getFieldNames();
                var q = from t1 in fns
                        join t2 in f2f on t1.Key equals t2.Key
                        select new
                        {
                            excelfield = t1.Value,
                            gisfield = t2.Value
                        };
                List<string> ss = new List<string>();
                foreach (var e in q)
                {
                    foreach (string e2 in e.gisfield)
                    {
                        string s = e.excelfield.Trim() + " -> " + e2.Trim();
                        ss.Add(s);
                    }
                }
                string str = string.Format("字段映射为:{0}", string.Join(",\t", ss.ToArray()));
                textBox2.Text = str;
                return;
            }
            textBox2.Text = string.Empty;
        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox currentListBox = sender as ListBox;
            object selectItem = currentListBox.SelectedItem;
            if (selectItem == null)
            {
                return;
            }
            currentListBox.BeginUpdate();
            currentListBox.SelectedIndexChanged -= listBox_SelectedIndexChanged;
            currentListBox.Items.Clear();
            currentListBox.Items.Add(selectItem);
            currentListBox.SelectedIndex = 0;
            Dictionary<string, string> d = new Dictionary<string, string>();
            foreach (ListBox lb2 in listBoxs)
            {
                if (lb2.SelectedItem != null)
                {
                    d.Add(lb2.Name, lb2.SelectedItem.ToString());
                }
            }
            List<LinqToExcel.Row> q = eh.data;
            using (Dictionary<string, string>.Enumerator enumerator2 = d.GetEnumerator())
            {
                while (enumerator2.MoveNext())
                {
                    KeyValuePair<string, string> i = enumerator2.Current;
                    q = q.Where(delegate (LinqToExcel.Row row)
                    {
                        string a2 = row[i.Key].ToString().Trim();
                        return a2 == i.Value.ToString().Trim();
                    }).ToList<LinqToExcel.Row>();
                }
            }
            using (List<ListBox>.Enumerator enumerator3 = listBoxs.GetEnumerator())
            {
                while (enumerator3.MoveNext())
                {
                    ListBox lb = enumerator3.Current;
                    if (lb.SelectedItem == null)
                    {
                        lb.Items.Clear();
                        string[] a = (from row in q
                                      select row[lb.Name].ToString().Trim()).Distinct<string>().ToArray<string>();
                        lb.Items.AddRange(a);
                    }
                }
            }
            updateTip();
            currentListBox.SelectedIndexChanged += listBox_SelectedIndexChanged;
            currentListBox.EndUpdate();
        }
        private List<IField> getField(ILayer lyr)
        {
            List<IField> r = new List<IField>();
            IFields fs = (lyr as IFeatureLayer).FeatureClass.Fields;
            for (int i = 0; i < fs.FieldCount; i++)
            {
                IField fd = fs.get_Field(i);
                r.Add(fd);
            }
            return r;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            base.Close();
            eh.open();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "uuiitwp")
            {
                MessageBox.Show("作者:uuiitwp\r\nTEL:15651785611\r\nEMAIL:uuiitwp@163.com");
            }
            IEnumerable<LinqToExcel.Row> q = from d in eh.data
                                             where d.Any((Cell x) => x.ToString().Contains(textBox1.Text))
                                             select d;
            using (List<ListBox>.Enumerator enumerator = listBoxs.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ListBox lb = enumerator.Current;
                    lb.BeginUpdate();
                    lb.Items.Clear();
                    lb.ClearSelected();
                    IEnumerable<string> q2 = from r in q
                                             select r[lb.Name].ToString().Trim();
                    lb.Items.AddRange(q2.Distinct().ToArray());
                    lb.EndUpdate();
                }
            }
            updateTip();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            reload();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (ArcMap.Document.FocusMap.SelectionCount == 0)
            {
                MessageBox.Show("未选中任何要素");
                return;
            }
            List<ILayer> lyrs = getSelectedLayers();
            Dictionary<int, string[]> f2f = eh.getFieldtoField();
            List<string> fds = new List<string>();
            foreach (string[] v in f2f.Values)
            {
                fds.AddRange(v);
            }
            fds = fds.Distinct<string>().ToList<string>();
            foreach (ILayer lyr in lyrs)
            {
                List<IField> fds2 = getField(lyr);
                IEnumerable<string> q = from row in fds2
                                        select row.Name;
                List<bool> flags = new List<bool>();
                foreach (string fdn2 in fds)
                {
                    flags.Add(q.Contains(fdn2));
                }
                if (!flags.All((bool x) => x))
                {
                    MessageBox.Show(string.Format("指定了图层{0}中以外的字段", lyr.Name));
                    return;
                }
                new List<bool>();
                List<egi> egi = getFieldsandValues();
                List<string[]> tt = new List<string[]>();
                foreach (egi e2 in egi)
                {
                    foreach (string e3 in e2.gisfield)
                    {
                        string[] ss = new string[]
                        {
                            e2.excelfield,
                            e3,
                            e2.input
                        };
                        tt.Add(ss);
                    }
                }
                using (List<string[]>.Enumerator enumerator5 = tt.GetEnumerator())
                {
                    while (enumerator5.MoveNext())
                    {
                        string[] fdn = enumerator5.Current;
                        var q2 = (from row in fds2
                                  where row.Name == fdn[1].Trim()
                                  select new
                                  {
                                      Type = row.Type,
                                      Name = row.Name,
                                      input = fdn[2].Trim()
                                  }).ToList();
                        foreach (var row2 in q2)
                        {
                            if (row2.Type == esriFieldType.esriFieldTypeDouble || row2.Type == esriFieldType.esriFieldTypeSingle)
                            {
                                if (!double.TryParse(row2.input, out double temp))
                                {
                                    MessageBox.Show(string.Format("不能对图层{0}的字段{1}赋值\"{2}\",因为该字段类型为{3}", new object[]
                                    {
                                        lyr.Name,
                                        row2.Name,
                                        row2.input,
                                        row2.Type
                                    }));
                                    return;
                                }
                            }
                            else if (row2.Type == esriFieldType.esriFieldTypeInteger || row2.Type == esriFieldType.esriFieldTypeSmallInteger)
                            {
                                if (!long.TryParse(row2.input, out long temp2))
                                {
                                    MessageBox.Show(string.Format("不能对图层{0}的字段{1}赋值\"{2}\",因为该字段类型为{3}", new object[]
                                    {
                                        lyr.Name,
                                        row2.Name,
                                        row2.input,
                                        row2.Type
                                    }));
                                    return;
                                }
                            }
                            else if (row2.Type == esriFieldType.esriFieldTypeOID)
                            {
                                MessageBox.Show(string.Format("不能对图层{0}的字段{1}赋值\"{2}\",因为该字段类型为{3}", new object[]
                                {
                                    lyr.Name,
                                    row2.Name,
                                    row2.input,
                                    row2.Type
                                }));
                                return;
                            }
                        }
                    }
                }
            }
            Geoprocessor g = new Geoprocessor();
            CalculateField cf = new CalculateField();
            string tp = eh.getFieldsType();
            bool isstr = false;
            if (tp == "STRING")
            {
                tp = "VB";
                isstr = true;
            }
            cf.expression_type = tp;
            textBox2.Text = string.Empty;
            List<egi> egi2 = getFieldsandValues();
            if (egi2.Count == 0)
            {
                MessageBox.Show("请选择");
                return;
            }
            foreach (ILayer lyr2 in lyrs)
            {
                cf.in_table = lyr2;
                foreach (egi e4 in egi2)
                {
                    foreach (string e5 in e4.gisfield)
                    {
                        cf.field = e5;
                        if (isstr)
                        {
                            cf.expression = "\"" + e4.input.Replace("\"", "\"\"") + "\"";
                            try
                            {
                                g.Execute(cf, null);
                                textBox2.AppendText(string.Format("完成对图层{0}中选中要素字段{1}进行赋值,表达式类型为{2},表达式为:{3}\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression));
                                textBox2.ScrollToCaret();
                                textBox2.Refresh();
                            }
                            catch
                            {
                                string em = string.Empty;
                                for (int i = 0; i < g.MessageCount; i++)
                                {
                                    em += g.GetMessage(i) + "\r\n";
                                }
                                textBox2.AppendText(string.Format("对图层{0}中选中要素字段{1}进行赋值时发生错误,表达式类型为{2},表达式为:{3}\r\n错误信息为:{4}\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression,
                                em));
                                textBox2.ScrollToCaret();
                                textBox2.Refresh();
                                if (!checkBox1.Checked)
                                {
                                    MessageBox.Show("赋值发生错误,查看文本框中记录");
                                    return;
                                }
                            }
                            textBox2.Refresh();
                        }
                        else
                        {
                            cf.expression = e4.input;
                            try
                            {
                                g.Execute(cf, null);
                                textBox2.AppendText(string.Format("完成对图层{0}中选中要素字段{1}进行赋值,表达式类型为{2},表达式为:{3}\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression
                            ));
                                textBox2.ScrollToCaret();
                                textBox2.Refresh();
                            }
                            catch
                            {
                                string em = string.Empty;
                                for (int i = 0; i < g.MessageCount; i++)
                                {
                                    em += g.GetMessage(i) + "\r\n";
                                }
                                textBox2.AppendText(string.Format("对图层{0}中选中要素字段{1}进行赋值时发生错误,表达式类型为{2},表达式为:{3}\r\n错误信息为:{4}\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression,
                                em
                            ));
                                textBox2.ScrollToCaret();
                                textBox2.Refresh();
                                if (!checkBox1.Checked)
                                {
                                    MessageBox.Show("赋值发生错误,查看文本框中记录");
                                    return;
                                }
                            }
                            textBox2.Refresh();
                        }
                    }
                }
            }
            textBox2.AppendText("全部完成");
            textBox2.ScrollToCaret();
            textBox2.Refresh();
        }

        private void kd(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button2_Click(null, null);
            }
        }
    }

}
