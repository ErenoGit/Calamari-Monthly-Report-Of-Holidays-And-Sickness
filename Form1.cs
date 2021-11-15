using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;

namespace CalamariMonthlyReportOfHolidaysAndSickness
{
    public partial class Form1 : Form
    {
        string baseCalamariURL, tokenAPI, month, year = "";

        private void buttonGenerateReport_Click(object sender, EventArgs e)
        {
            progressBar.Value = 0;

            baseCalamariURL = textBoxURL.Text;
            tokenAPI = textBoxAPI.Text;
            month = comboBoxMonth.SelectedItem?.ToString();
            year = comboBoxYear.SelectedItem?.ToString();

            if (String.IsNullOrEmpty(baseCalamariURL))
                MessageBox.Show("Pole 'Bazowy URL Calamari' musi być uzupełnione!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (String.IsNullOrEmpty(tokenAPI))
                MessageBox.Show("Pole 'Token API' musi być uzupełnione!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (String.IsNullOrEmpty(month))
                MessageBox.Show("Miesiąc musi być wybrany!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (String.IsNullOrEmpty(year))
                MessageBox.Show("Rok musi być wybrany!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            baseCalamariURL = baseCalamariURL.Trim();
            if (baseCalamariURL.EndsWith("api/"))
                baseCalamariURL = baseCalamariURL.Remove(baseCalamariURL.Length - 4);
            if (baseCalamariURL.EndsWith("api"))
                baseCalamariURL = baseCalamariURL.Remove(baseCalamariURL.Length - 3);
            if (baseCalamariURL.EndsWith("/"))
                baseCalamariURL = baseCalamariURL.Remove(baseCalamariURL.Length - 1);

            tokenAPI = tokenAPI.Trim();

            GenerateAndSaveReport();
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists("config.txt"))
                {
                    string[] content = File.ReadAllLines(@"config.txt");

                    if (content.Length >= 4)
                    {
                        baseCalamariURL = content[0];
                        tokenAPI = content[1];
                        month = content[2];
                        year = content[3];

                        textBoxURL.Text = baseCalamariURL;
                        textBoxAPI.Text = tokenAPI;
                        comboBoxMonth.SelectedIndex = comboBoxMonth.FindStringExact(month);
                        comboBoxYear.SelectedIndex = comboBoxYear.FindStringExact(year);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception on reading config file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            baseCalamariURL = textBoxURL.Text;
            tokenAPI = textBoxAPI.Text;
            month = comboBoxMonth.SelectedItem?.ToString();
            year = comboBoxYear.SelectedItem?.ToString();

            File.WriteAllText("config.txt", baseCalamariURL + Environment.NewLine + tokenAPI + Environment.NewLine + month + Environment.NewLine + year);
        }


        private void GenerateAndSaveReport()
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                using (WebClient client = new WebClient())
                {
                    client.Encoding = Encoding.UTF8;
                    client.UseDefaultCredentials = true;
                    client.Headers[HttpRequestHeader.Authorization] = "Basic " + Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes("calamari:" + tokenAPI));
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Accept] = "application/json";
                    client.BaseAddress = baseCalamariURL;

                    List<SingleEmployeeReport> listOfEmployeesData = GetListOfEmployees(client);

                    progressBar.Maximum = listOfEmployeesData.Count();

                    foreach (SingleEmployeeReport singleEmployeeData in listOfEmployeesData)
                    {
                        AbsencesObject absences = GetListOfAbsences(client, singleEmployeeData.email);
                        singleEmployeeData.absences = absences;
                        progressBar.Value += 1;
                    }

                    CreateAndSaveExcelFile(listOfEmployeesData);
                }

            }
            catch (WebException ex)
            {
                progressBar.Value = 0;
                MessageBox.Show(ex.Message, "WebException", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                progressBar.Value = 0;
                MessageBox.Show(ex.Message, "Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private List<SingleEmployeeReport> GetListOfEmployees(WebClient client)
        {
            List<SingleEmployeeReport> listOfEmployeesData = new List<SingleEmployeeReport>();

            dynamic obj = new JObject();
            obj.page = 0;
            string data = JsonConvert.SerializeObject(obj);

            string response = client.UploadString($"/api/employees/v1/list", data);
            dynamic employeesList = JObject.Parse(response);

            foreach (var item in employeesList.employees)
            {
                SingleEmployeeReport singleEmployeeReport = new SingleEmployeeReport();
                singleEmployeeReport.id = item.id.Value;
                singleEmployeeReport.firstName = item.firstName.Value;
                singleEmployeeReport.lastName = item.lastName.Value;
                singleEmployeeReport.email = item.email.Value;

                listOfEmployeesData.Add(singleEmployeeReport);
            }

            return listOfEmployeesData;
        }

        private AbsencesObject GetListOfAbsences(WebClient client, string employeeEmail)
        {
            AbsencesObject singleAbsencesObject = new AbsencesObject();

            int monthInt = 0;
            string startDate, endDate;

            if (month.Contains("Styczeń"))
                monthInt = 1;
            else if (month.Contains("Luty"))
                monthInt = 2;
            else if (month.Contains("Marzec"))
                monthInt = 3;
            else if (month.Contains("Kwiecień"))
                monthInt = 4;
            else if (month.Contains("Maj"))
                monthInt = 5;
            else if (month.Contains("Czerwiec"))
                monthInt = 6;
            else if (month.Contains("Lipiec"))
                monthInt = 7;
            else if (month.Contains("Sierpień"))
                monthInt = 8;
            else if (month.Contains("Wrzesień"))
                monthInt = 9;
            else if (month.Contains("Październik"))
                monthInt = 10;
            else if (month.Contains("Listopad"))
                monthInt = 11;
            else if (month.Contains("Grudzień"))
                monthInt = 12;

            startDate = year + "-" + monthInt + "-" + "1";

            int daysInMonth = DateTime.DaysInMonth(Int32.Parse(year), monthInt);
            endDate = year + "-" + monthInt + "-" + daysInMonth;


            dynamic obj = new JObject();
            obj.employee = employeeEmail;
            obj.from = startDate;
            obj.to = endDate;
            string data = JsonConvert.SerializeObject(obj);

            client.Encoding = Encoding.UTF8;
            client.Headers[HttpRequestHeader.ContentType] = "application/json";
            client.Headers[HttpRequestHeader.Accept] = "application/json";

            string response = client.UploadString($"/api/leave/request/v1/find", data);
            Thread.Sleep(100);

            dynamic absences = JArray.Parse(response);

            foreach (var item in absences)
            {
                string from = item.from.Value;
                string to = item.to.Value;
                string absenceTypeName = item.absenceTypeName.Value;
                string status = item.status.Value;
                string reason = item.reason.Value;

                DateTime fromDateTime = DateTime.Parse(from);
                from = fromDateTime.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

                DateTime toDateTime = DateTime.Parse(to);
                to = toDateTime.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);

                if (status == "ACCEPTED")
                {
                    if (absenceTypeName == "Choroba")
                    {
                        if(from == to)
                            singleAbsencesObject.sickness += from;
                        else
                            singleAbsencesObject.sickness += from + " - " + to;

                        if(!String.IsNullOrEmpty(reason))
                            singleAbsencesObject.sickness += " (" + reason + ")";

                        singleAbsencesObject.sickness += ", ";
                        continue;
                    }
                    else
                    {
                        if (absenceTypeName != "Delegacja" && absenceTypeName != "Praca zdalna")
                        {
                            if (from == to)
                                singleAbsencesObject.holidays += from + " (" + absenceTypeName;
                            else
                                singleAbsencesObject.holidays += from + " - " + to + " (" + absenceTypeName;

                            if (!String.IsNullOrEmpty(reason))
                                singleAbsencesObject.holidays += " - " + reason;

                            singleAbsencesObject.holidays += "), ";
                            continue;
                        }
                    }
                }

            }

            if (singleAbsencesObject.sickness.EndsWith(", "))
                singleAbsencesObject.sickness = singleAbsencesObject.sickness.Remove(singleAbsencesObject.sickness.Length - 2);

            if (singleAbsencesObject.holidays.EndsWith(", "))
                singleAbsencesObject.holidays = singleAbsencesObject.holidays.Remove(singleAbsencesObject.holidays.Length - 2);

            return singleAbsencesObject;
        }

        private void CreateAndSaveExcelFile(List<SingleEmployeeReport> listOfEmployeesData)
        {
            string p_strPath = "Raport_Calamari.xlsx";
            ExcelPackage excel = new ExcelPackage();

            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            workSheet.Cells[1, 1].Value = "Imię i nazwisko";
            workSheet.Cells[1, 2].Value = "Urlop";
            workSheet.Cells[1, 3].Value = "Chorobowe";

            int recordIndex = 2;

            foreach (SingleEmployeeReport employeeData in listOfEmployeesData)
            {
                workSheet.Cells[recordIndex, 1].Value = employeeData.firstName + " " + employeeData.lastName;
                workSheet.Cells[recordIndex, 2].Value = employeeData.absences.holidays;
                workSheet.Cells[recordIndex, 3].Value = employeeData.absences.sickness;
                recordIndex++;
            }

            workSheet.Column(1).AutoFit();
            //workSheet.Column(2).AutoFit();
            //workSheet.Column(3).AutoFit();

            workSheet.Column(2).Width = 100;
            workSheet.Column(3).Width = 100;

            workSheet.Column(2).Style.WrapText = true;
            workSheet.Column(3).Style.WrapText = true;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Zapisz plik z raportem";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "*.xlsx|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                p_strPath = saveFileDialog.FileName;     
            }

            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            excel.Dispose();
        }


    }
}
