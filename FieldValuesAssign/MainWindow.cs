using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.DataManagementTools;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geoprocessor;
using LinqToExcel;

namespace FieldValuesAssign
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();

            this.textBox2.Location = new Point(20, 160);
            this.textBox2.Multiline = true;
            this.textBox2.ReadOnly = true;
            this.textBox2.ScrollBars = ScrollBars.Both;
            this.eh = new ExcelHelper();
            int listBoxCount = this.eh.getListBoxCount();
            KeyValuePair<int, string>[] fieldNames = this.eh.getFieldNames().ToArray<KeyValuePair<int, string>>();
            IEnumerable<string> q = (from row in fieldNames
                                     select row.Value).Distinct<string>();
            if (q.Count<string>() != fieldNames.Count<KeyValuePair<int, string>>())
            {
                MessageBox.Show("存在指定字段重复");
                this.button1_Click(null, null);
                return;
            }
            for (int i = 0; i < listBoxCount; i++)
            {
                ListBox lb = new ListBox();
                lb.Location = new Point(10 + i * 126, 60);
                lb.SelectedIndexChanged += this.listBox_SelectedIndexChanged;
                lb.HorizontalScrollbar = true;
                Label j = new Label();
                lb.Name = (j.Text = fieldNames[i].Value);
                j.Location = new Point(10 + i * 126, 35);
                this.listBoxs.Add(lb);
                this.Labels.Add(j);
                base.Controls.Add(lb);
                base.Controls.Add(j);
            }
            resize(null, null);
            this.reload();
        }
        private List<ListBox> listBoxs = new List<ListBox>();
        private List<Label> Labels = new List<Label>();
        private ExcelHelper eh;
        private void reload()
        {
            using (List<ListBox>.Enumerator enumerator = this.listBoxs.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ListBox lb = enumerator.Current;
                    lb.BeginUpdate();
                    string[] q = (from d in this.eh.data
                                  select d[lb.Name].ToString().Trim()).Distinct<string>().ToArray<string>();
                    lb.Items.Clear();
                    lb.ClearSelected();
                    lb.Items.AddRange(q);
                    lb.EndUpdate();
                }
            }
            this.updateTip();
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
            this.textBox2.Size = new Size(base.Size.Width - 50, base.Size.Height - 240);
            this.textBox1.Size = new Size(base.Size.Width - 150, this.textBox1.Size.Height);
            var listBoxCount = listBoxs.Count;
            for (int i = 0; i < listBoxCount; i++)
            {
                var l = Labels[i];
                l.Location = new Point(i * ((this.Width - 10) / listBoxCount) + 10, l.Location.Y);
                var lb = listBoxs[i];
                lb.Location = new Point(i * ((this.Width - 10) / listBoxCount) + 10, lb.Location.Y);
                lb.Width = (this.Width - 40 - listBoxCount * 20) / listBoxCount;
            }
        }

        public List<egi> getFieldsandValues()
        {
            if (this.listBoxs.All((ListBox x) => x.Items.Count == 1))
            {
                Dictionary<int, string[]> f2f = this.eh.getFieldtoField();
                Dictionary<int, string> fns = this.eh.getFieldNames();
                IEnumerable<egi> q = from t1 in fns
                                     join t2 in f2f on t1.Key equals t2.Key
                                     join t3 in this.listBoxs on t1.Value equals t3.Name
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
            if (this.listBoxs.All((ListBox x) => x.Items.Count == 1))
            {
                Dictionary<int, string[]> f2f = this.eh.getFieldtoField();
                Dictionary<int, string> fns = this.eh.getFieldNames();
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
                this.textBox2.Text = str;
                return;
            }
            this.textBox2.Text = string.Empty;
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
            currentListBox.SelectedIndexChanged -= this.listBox_SelectedIndexChanged;
            currentListBox.Items.Clear();
            currentListBox.Items.Add(selectItem);
            currentListBox.SelectedIndex = 0;
            Dictionary<string, string> d = new Dictionary<string, string>();
            foreach (ListBox lb2 in this.listBoxs)
            {
                if (lb2.SelectedItem != null)
                {
                    d.Add(lb2.Name, lb2.SelectedItem.ToString());
                }
            }
            var q = this.eh.data;
            using (Dictionary<string, string>.Enumerator enumerator2 = d.GetEnumerator())
            {
                while (enumerator2.MoveNext())
                {
                    KeyValuePair<string, string> i = enumerator2.Current;
                    q = q.Where(delegate(LinqToExcel.Row row)
                    {
                        string a2 = row[i.Key].ToString().Trim();
                        return a2 == i.Value.ToString().Trim();
                    }).ToList<LinqToExcel.Row>();
                }
            }
            using (List<ListBox>.Enumerator enumerator3 = this.listBoxs.GetEnumerator())
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
            this.updateTip();
            currentListBox.SelectedIndexChanged += this.listBox_SelectedIndexChanged;
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
            this.eh.open();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IEnumerable<LinqToExcel.Row> q = from d in this.eh.data
                                             where d.Any((Cell x) => x.ToString().Contains(this.textBox1.Text))
                                             select d;
            using (List<ListBox>.Enumerator enumerator = this.listBoxs.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    ListBox lb = enumerator.Current;
                    lb.BeginUpdate();
                    lb.Items.Clear();
                    lb.ClearSelected();
                    var q2 = from r in q
                             select r[lb.Name].ToString().Trim();
                    lb.Items.AddRange(q2.Distinct().ToArray());
                    lb.EndUpdate();
                }
            }
            this.updateTip();
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
            var lyrs = this.getSelectedLayers();
            var f2f = this.eh.getFieldtoField();
            var fds = new List<string>();
            foreach (string[] v in f2f.Values)
            {
                fds.AddRange(v);
            }
            fds = fds.Distinct<string>().ToList<string>();
            foreach (var lyr in lyrs)
            {
                var fds2 = this.getField(lyr);
                var q = from row in fds2
                        select row.Name;
                var flags = new List<bool>();
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
                var egi = this.getFieldsandValues();
                var tt = new List<string[]>();
                foreach (egi e2 in egi)
                {
                    foreach (var e3 in e2.gisfield)
                    {
                        var ss = new string[]
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
                                double temp;
                                if (!double.TryParse(row2.input, out temp))
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
                                long temp2;
                                if (!long.TryParse(row2.input, out temp2))
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
            string tp = this.eh.getFieldsType();
            bool isstr = false;
            if (tp == "STRING")
            {
                tp = "VB";
                isstr = true;
            }
            cf.expression_type = tp;
            this.textBox2.Text = string.Empty;
            List<egi> egi2 = this.getFieldsandValues();
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
                                TextBox textBox = this.textBox2;
                                textBox.Text += string.Format("完成对图层{0}中选中要素字段{1}进行赋值,表达式类型为{2},表达式为:{3}\r\n\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression
                            );
                            }
                            catch
                            {
                                TextBox textBox = this.textBox2;
                                var em = string.Empty;
                                for (int i = 0; i < g.MessageCount; i++)
                                {
                                    em += g.GetMessage(i) + "\r\n";
                                }
                                textBox.Text += string.Format("对图层{0}中选中要素字段{1}进行赋值时发生错误,表达式类型为{2},表达式为:{3}\r\n错误信息为:{4}\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression,
                                em
                            );
                                if (!checkBox1.Checked)
                                {
                                    MessageBox.Show("赋值发生错误,查看文本框中记录");
                                    return;
                                }
                            }
                            this.textBox2.Refresh();
                        }
                        else
                        {
                            cf.expression = e4.input;
                            try
                            {
                                g.Execute(cf, null);
                                this.textBox2.Text += string.Format("完成对图层{0}中选中要素字段{1}进行赋值,表达式类型为{2},表达式为:{3}\r\n\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression
                            );
                            }
                            catch
                            {
                                var em = string.Empty;
                                for (int i = 0; i < g.MessageCount; i++)
                                {
                                    em += g.GetMessage(i) + "\r\n";
                                }
                                this.textBox2.Text += string.Format("对图层{0}中选中要素字段{1}进行赋值时发生错误,表达式类型为{2},表达式为:{3}\r\n错误信息为:{4}\r\n",
                                lyr2.Name,
                                e5,
                                cf.expression_type,
                                cf.expression,
                                em
                            );
                                if (!checkBox1.Checked)
                                {
                                    MessageBox.Show("赋值发生错误,查看文本框中记录");
                                    return;
                                }
                            }
                            this.textBox2.Refresh();
                        }
                    }
                }
            }
            textBox2.Text += "全部完成";
            this.textBox2.Refresh();
        }
    }

}
