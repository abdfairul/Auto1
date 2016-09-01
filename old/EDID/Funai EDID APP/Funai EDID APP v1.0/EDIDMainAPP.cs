/****************************** Module Header ******************************\
Module Name:  SWTuning.cs
Project:      Software Automation for Software Tuning
Copyright (c) Funai SGP and Funai MY R&D.

This source is subject to the Funai Software Licensing Agreement.
Kindly contact Funai Co. Ltd. for further details.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
 * 
 * [Revision History]
 * v1.0.0.0, 25th May 2015 - [Chandru] First official release
 * v1.0.0.1, 17th Jun 2015 - [Kenny] Fix Excel module open after close
 *                         - changed text color of NG cases to White
 *                         - fix application crash due to not available data from QD882
 *                         - change HDMI_D to HDMI_H, output 1080i (HDMI) and 768 (VGA) based on EDID test spec
 * 
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using SAAL;
using MinimalisticTelnet;
using System.Reflection;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace Funai_EDID_APP_v1._0
{
    public partial class EDIDMainAPP : Form
    {
        string menutree_fname;
        SAAL_Interface mysaal = new SAAL_Interface();
        Int32 element_count = 0;

        public EDIDMainAPP()
        {
            InitializeComponent();
            foreach (string port in System.IO.Ports.SerialPort.GetPortNames())
            {
                cb_Comport.Items.Add(port);
            }

            label1.Visible = false;
        }

        public enum StatusDevice
        {
            idle,
            connect,
            disconnect
        }
        StatusDevice EDID_APP_STATUS = StatusDevice.idle;

        private void Form1_Load(object sender, EventArgs e)
        {
            btn_start.Enabled = false;
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel files|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                /* Returns the filename of the Menu Tree to be used */
                menutree_fname = openFileDialog1.FileName;
                //this.textBox_path.Text = openFileDialog1.FileName;
                this.textBox_path.Text = System.IO.Path.GetFileName(openFileDialog1.FileName);
            }
        }

        #region (Combobox panel, source)
        private void combobox_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource != null && dataGridView2.DataSource != null)
            {
                dataGridView1.DataSource = null;
                dataGridView2.DataSource = null;
            }

            else
            {
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
            }

            label1.Visible = true;
            label1.Text = "Ready";
            label1.ForeColor = System.Drawing.Color.Black;
        }

        private void cb_Panel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource != null && dataGridView2.DataSource != null)
            {
                dataGridView1.DataSource = null;
                dataGridView2.DataSource = null;
            }

            else
            {
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
            }
        }
        #endregion


        private void btn_start_Click(object sender, EventArgs e)
        {
            #region Clear Data gridView & Check load excel file is filled
            if (string.IsNullOrEmpty(textBox_path.Text) || cb_Panel.SelectedIndex == -1 || combobox_sheet.SelectedIndex == -1)
            {
                MessageBox.Show("Please make sure you dont have any missing field");
            }
            else
            {
                    if (dataGridView2.DataSource != null)
                    {
                        dataGridView2.DataSource = null;
                    }

                    else
                    {
                        dataGridView2.Rows.Clear();
                    }
                    //disable datagridview column sorting by user
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    foreach (DataGridViewColumn column in dataGridView2.Columns)
                    {
                        column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }

            #endregion

            #region load excel file
                    label1.Visible = true;
                    label1.Text = "Loading";
                    label1.ForeColor = System.Drawing.Color.Blue;

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet = null;
                    dataGridView1.Rows.Clear();

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(menutree_fname, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                    string sheet = cb_Panel.Text;

                    switch (sheet)
                    {
                        case "WXGA":
                            if (combobox_sheet.Text == "HDMI1")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);
                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("B5", "B260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {

                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI2")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("C5", "C260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }

                                }
                            }

                            if (combobox_sheet.Text == "HDMI3")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("D5", "D260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                        if (combobox_sheet.Text == "HDMI4")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("E5", "E260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "HDMI5")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("F5", "F260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "VGA")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.VGA, SAAL_Interface.EN_QD882_FORMAT.Q1024_768_60Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["VGA"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("B4", "B131");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }
                            break;

                        case "HD+":
                            if (combobox_sheet.Text == "HDMI1")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("H5", "H260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI2")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("I5", "I260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI3")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("J5", "J260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                        if (combobox_sheet.Text == "HDMI4")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("K5", "K260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "HDMI5")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("L5", "L260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "VGA")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.VGA, SAAL_Interface.EN_QD882_FORMAT.Q1024_768_60Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["VGA"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("D4", "D131");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }
                            break;

                        case "FHD":
                            if (combobox_sheet.Text == "HDMI1")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("N5", "N260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI2")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("O5", "O260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI3")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("P5", "P260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                        if (combobox_sheet.Text == "HDMI4")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("Q5", "Q260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "HDMI5")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("R5", "R260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "VGA")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.VGA, SAAL_Interface.EN_QD882_FORMAT.Q1024_768_60Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["VGA"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("F4", "F131");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }
                            break;
                        case "UHD":
                            if (combobox_sheet.Text == "HDMI1")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("T5", "T260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI2")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("U5", "U260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                            if (combobox_sheet.Text == "HDMI3")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("V5", "V260");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }

                        if (combobox_sheet.Text == "HDMI4")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("W5", "W260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "HDMI5")
                        {
                            mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.HDMI_H, SAAL_Interface.EN_QD882_FORMAT.Q1920_1080_30Hz);

                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["HDMI"];
                            Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("X5", "X260");
                            int numberofRows = excelCell.Rows.Count;
                            element_count = numberofRows;
                            //numberofRows /= 16;
                            for (int i = 0; i < numberofRows; i += 16)
                            {
                                dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                            excelCell.Cells[i + 2, 1].Value,
                                                            excelCell.Cells[i + 3, 1].Value,
                                                            excelCell.Cells[i + 4, 1].Value,
                                                            excelCell.Cells[i + 5, 1].Value,
                                                            excelCell.Cells[i + 6, 1].Value,
                                                            excelCell.Cells[i + 7, 1].Value,
                                                            excelCell.Cells[i + 8, 1].Value,
                                                            excelCell.Cells[i + 9, 1].Value,
                                                            excelCell.Cells[i + 10, 1].Value,
                                                            excelCell.Cells[i + 11, 1].Value,
                                                            excelCell.Cells[i + 12, 1].Value,
                                                            excelCell.Cells[i + 13, 1].Value,
                                                            excelCell.Cells[i + 14, 1].Value,
                                                            excelCell.Cells[i + 15, 1].Value,
                                                            excelCell.Cells[i + 16, 1].Value);
                                foreach (DataGridViewRow row in dataGridView1.Rows)
                                {
                                    row.HeaderCell.Value = (row.Index).ToString("X0");
                                }
                            }
                        }

                        if (combobox_sheet.Text == "VGA")
                            {
                                mysaal.QD882_SetSource(SAAL_Interface.EN_QD882_SOURCE.VGA, SAAL_Interface.EN_QD882_FORMAT.Q1024_768_60Hz);

                                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["VGA"];
                                Excel.Range excelCell = (Excel.Range)xlWorkSheet.get_Range("H4", "H131");
                                int numberofRows = excelCell.Rows.Count;
                                element_count = numberofRows;
                                //numberofRows /= 16;
                                for (int i = 0; i < numberofRows; i += 16)
                                {
                                    dataGridView1.Rows.Add("", excelCell.Cells[i + 1, 1].Value,
                                                                excelCell.Cells[i + 2, 1].Value,
                                                                excelCell.Cells[i + 3, 1].Value,
                                                                excelCell.Cells[i + 4, 1].Value,
                                                                excelCell.Cells[i + 5, 1].Value,
                                                                excelCell.Cells[i + 6, 1].Value,
                                                                excelCell.Cells[i + 7, 1].Value,
                                                                excelCell.Cells[i + 8, 1].Value,
                                                                excelCell.Cells[i + 9, 1].Value,
                                                                excelCell.Cells[i + 10, 1].Value,
                                                                excelCell.Cells[i + 11, 1].Value,
                                                                excelCell.Cells[i + 12, 1].Value,
                                                                excelCell.Cells[i + 13, 1].Value,
                                                                excelCell.Cells[i + 14, 1].Value,
                                                                excelCell.Cells[i + 15, 1].Value,
                                                                excelCell.Cells[i + 16, 1].Value);
                                    foreach (DataGridViewRow row in dataGridView1.Rows)
                                    {
                                        row.HeaderCell.Value = (row.Index).ToString("X0");
                                    }
                                }
                            }
                            break;
                    }

                    if (xlWorkBook != null && xlWorkSheet != null && xlApp != null)
                    {
                        xlWorkBook.Close(false, null, false);
                        xlApp.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkBook);
                        releaseObject(xlApp);
                    }

                    #endregion

            #region get EDID from QD882
                    string[] edid_data;

                    mysaal.QD882_GetEDID(out edid_data);

                    var f = from s in edid_data select s;
                    int c = f.Count();

                    if (c < element_count)
                    {
                        Console.WriteLine("ERROR Reading Quantum Data!! Please check your device.");
                        return; // boundary check - if fail, do not do checking
                    }

                    for (int i = 0; i < c; i += 16)
                    {

                        dataGridView2.Rows.Add("", edid_data[0 + i],
                                                    edid_data[1 + i],
                                                    edid_data[2 + i],
                                                    edid_data[3 + i],
                                                    edid_data[4 + i],
                                                    edid_data[5 + i],
                                                    edid_data[6 + i],
                                                    edid_data[7 + i],
                                                    edid_data[8 + i],
                                                    edid_data[9 + i],
                                                    edid_data[10 + i],
                                                    edid_data[11 + i],
                                                    edid_data[12 + i],
                                                    edid_data[13 + i],
                                                    edid_data[14 + i],
                                                    edid_data[15 + i]);

                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            row.HeaderCell.Value = (row.Index).ToString("X0");
                        }
                    }

                    #endregion
            Thread.Sleep(3000);
            #region Compare datagridview1 and datagridview2

            int no_of_col = dataGridView1.Columns.Count;
            int no_of_row = dataGridView1.Rows.Count;
            int j;
            int k;
            var B = "";
            var A = "";
            for (j = 0; j < no_of_row; j++)
            {
                for (k = 0; k < no_of_col; k++)
                {
                    ////if statement value is null replace with ZERO
                    //if (dataGridView1.Rows[k].Cells[j].Value != null && !string.IsNullOrWhiteSpace(dataGridView1.Rows[k].Cells[j].Value.ToString()))
                    //{
                    //    B = dataGridView1.Rows[k].Cells[j].Value.ToString();
                    //}
                    ////if db value is null replace with zero
                    //if (dataGridView2.Rows[k].Cells[j].Value != null && !string.IsNullOrWhiteSpace(dataGridView2.Rows[k].Cells[j].Value.ToString()))
                    //{
                    //    A = dataGridView2.Rows[k].Cells[j].Value.ToString();
                       
                    //}

                    //if statement value is null replace with ZERO
                    if (dataGridView1.Rows[j].Cells[k].Value != null && !string.IsNullOrWhiteSpace(dataGridView1.Rows[j].Cells[k].Value.ToString()))
                    {
                        B = dataGridView1.Rows[j].Cells[k].Value.ToString();
                    }
                    //if db value is null replace with zero
                    if (dataGridView2.Rows[j].Cells[k].Value != null && !string.IsNullOrWhiteSpace(dataGridView2.Rows[j].Cells[k].Value.ToString()))
                    {
                        A = dataGridView2.Rows[j].Cells[k].Value.ToString();

                    }

                    if (B != A)
                    {
                        dataGridView1.Rows[j].Cells[k].Style.BackColor = Color.Red;
                        dataGridView1.Rows[j].Cells[k].Style.ForeColor = Color.White;
                        dataGridView2.Rows[j].Cells[k].Style.BackColor = Color.Red;
                        dataGridView2.Rows[j].Cells[k].Style.ForeColor = Color.White;
                        label1.Visible = true;
                        label1.Text = "Result: NG";
                        label1.ForeColor = System.Drawing.Color.Red;
                    }
                }

                if (B == A)
                {
                    label1.Text = "Result: OK";
                    label1.ForeColor = System.Drawing.Color.Green;
                }
            }

                    
            #endregion
            }  

        }

        #region Button(connect, Clear, About)
        private void btn_Connect_Click(object sender, EventArgs e)
        {
            if (EDID_APP_STATUS == StatusDevice.connect)
            {
                EDID_APP_STATUS = StatusDevice.disconnect;
            }
            else
            {
                EDID_APP_STATUS = StatusDevice.connect;
            }

            if (EDID_APP_STATUS == StatusDevice.connect)
            {
                if (string.IsNullOrEmpty(cb_Comport.Text))
                {
                    MessageBox.Show("No port selected");
                }
                else
                {
                    string curItem = cb_Comport.SelectedItem.ToString();
                    int item = cb_Comport.FindString(curItem);
                    mysaal.QD882_Setup(curItem, 9600);
                    MessageBox.Show("QuantumData882 Connected!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btn_Connect.Text = "Disconnect";
                    lblControl.BackColor = System.Drawing.Color.Green;
                    combobox_sheet.Enabled = true;
                    btn_start.Enabled = true;

                }

            }

            else if (EDID_APP_STATUS == StatusDevice.disconnect)
            {
                mysaal.QD882_ClosePort();
                btn_Connect.Text = "Connect";
                lblControl.BackColor = System.Drawing.Color.Red;
                btn_start.Enabled = false;
                textBox_path.Text = String.Empty;
                combobox_sheet.SelectedIndex = -1;
                cb_Panel.Text = String.Empty;
                label1.Visible = false;
                if (dataGridView1.DataSource != null && dataGridView2.DataSource != null)
                {
                    dataGridView1.DataSource = null;
                    dataGridView2.DataSource = null;
                }

                else
                {
                    dataGridView1.Rows.Clear();
                    dataGridView2.Rows.Clear();
                }

            }          
        }

        private void btn_Clear_Click(object sender, EventArgs e)
        {
            textBox_path.Text = String.Empty;
            combobox_sheet.SelectedIndex = -1;
            cb_Panel.Text = String.Empty;
            label1.Visible = false;

            if (dataGridView1.DataSource != null && dataGridView2.DataSource != null)
            {
                dataGridView1.DataSource = null;
                dataGridView2.DataSource = null;
            }

            else
            {
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
            }
        }

        private void About_EDID_Click(object sender, EventArgs e)
        {
            About_EDID aboutbox = new About_EDID();
            aboutbox.Show();
        }
        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen)
                serialPort1.Close();
        }

        /* Function:    releaseObject
         * Parameters:  Input -> obj - supplies the valid COM object
         * Description: frees up the COM resource
         */
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
