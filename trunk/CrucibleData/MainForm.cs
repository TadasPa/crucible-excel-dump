using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Net;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;
using CrucibleData.JiraSR;
using System.Diagnostics;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace CrucibleData
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
        string GetCvValue(JiraSR.RemoteIssue item, string name)
        {
            string ret = "";
            if (null != item)
            {
                foreach (RemoteCustomFieldValue cv in item.customFieldValues)
                {
                    if (cv.customfieldId == name)
                    {
                        ret = String.Join(",", cv.values);
                        break;
                    }
                }
            }
            return ret;
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reload();
            propertyGrid1.SelectedObject = Properties.Settings.Default;
        }
        private void log(string msg)
        {
            Console.WriteLine(msg);
        }

        private void rp(string msg)
        {
            backgroundWorker.ReportProgress(0, msg);
        }

        private string decorate(string xml, string tag)
        {
            return String.Format("{0}<{1}>{2}</{1}>",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>",
                tag, xml);
        }
        private void DeleteSheet(Excel.Workbook workBook, int index)
        {
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[index];
            workSheet.Delete();
        }
        private void PrepareExcelWorkbook(Excel.Workbook workBook)
        {
            DeleteSheet(workBook, 2);
            DeleteSheet(workBook, 2);//Same index, as the third will become two when third is deleted.
            Excel.Worksheet oSheet = (Excel.Worksheet)workBook.Sheets[1];
            oSheet.Name = "Crucible Data - " + Properties.Settings.Default.CrucibleProject;
            int index = 1;
            int column = 1;
            oSheet.Cells[index, column++] = "Author Username";
            oSheet.Cells[index, column++] = "Author Display Name";
            oSheet.Cells[index, column++] = "Closed Date";
            oSheet.Cells[index, column++] = "Created Date";
            oSheet.Cells[index, column++] = "Creator Username";
            oSheet.Cells[index, column++] = "Creator Display Name";
            oSheet.Cells[index, column++] = "Description";
            Excel.Range desc = (Excel.Range)(oSheet.Cells[index, column - 1]);
            desc.EntireColumn.ColumnWidth = 255;
            oSheet.Cells[index, column++] = "Due Date";
            oSheet.Cells[index, column++] = "Linked Jira Item";
            oSheet.Cells[index, column++] = "Jira - Affects Version/s";
            oSheet.Cells[index, column++] = "Jira - RAN Number";
            oSheet.Cells[index, column++] = "Moderator Username";
            oSheet.Cells[index, column++] = "Moderator Display Name";
            oSheet.Cells[index, column++] = "Name";
            oSheet.Cells[index, column++] = "Review Id";
            oSheet.Cells[index, column++] = "Project Key";
            oSheet.Cells[index, column++] = "Review State";
            oSheet.Cells[index, column++] = "Summary";
            oSheet.Cells[index, column++] = "Comment - From Line";
            oSheet.Cells[index, column++] = "Comment - To Line";
            oSheet.Cells[index, column++] = "Comment - Revision - Range";
            oSheet.Cells[index, column++] = "Comment - Created Date";
            oSheet.Cells[index, column++] = "Comment - Is Defect?";
            oSheet.Cells[index, column++] = "Comment - Is Deleted?";
            oSheet.Cells[index, column++] = "Comment - Is In Draft?";
            oSheet.Cells[index, column++] = "Comment - Comment";
            oSheet.Cells[index, column++] = "Comment - Author Username";
            oSheet.Cells[index, column++] = "Comment - Author Display Name";
            oSheet.Cells[index, column++] = "Comment - Replies";
            oSheet.Cells[index, column++] = "Comment - Is Accepted?";
            Excel.Range rng = (Excel.Range)(oSheet.Cells[index, column++]);
            rng.EntireRow.Font.Size = 9;
            rng.EntireRow.Font.Bold = true;
            rng.EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToWin32(System.Drawing.Color.LightGray);
            rng.EntireRow.WrapText = true;
            oSheet.EnableAutoFilter = true;
        }
        private string ProcessReplies(System.Xml.XmlElement[] replies, int level)
        {
            string reply = "";
            foreach (System.Xml.XmlElement elem in replies)
            {
                XmlSerializer vlcdr = new XmlSerializer(typeof(generalCommentData));
                generalCommentData gcd = (generalCommentData)vlcdr.Deserialize(new StringReader(
                    decorate(elem.InnerXml, "generalCommentData")));
                for (int i = 0; i < level; i++)
                {
                    reply += "-";
                }
                reply += String.Format("[{0}] - [{1}]\r\n", gcd.user.displayName, gcd.message);

                string checkAccepted = gcd.message.ToLower();
                foreach (string acceptString in Properties.Settings.Default.CrucibleAcceptStrings)
                {
                    if (checkAccepted.Contains(acceptString))
                    {
                        accepted = true;
                        break;
                    }
                }

                if (gcd.replies.Any != null)
                {
                    reply += ProcessReplies(gcd.replies.Any, level + 1);
                }
            }
            return reply;
        }

        bool accepted = false;
        int CommentCount = 1;

        private string bf(string date)
        {
            if (null != date && "" != date)
            {
                return date.Substring(0, date.IndexOf("T"));
            }
            else
            {
                return "";
            }
        }

        string verToString(RemoteVersion[] versions)
        {
            string fv = "";
            if (null != versions)
            {
                foreach (RemoteVersion rv in versions)
                {
                    fv += (string)(revHash[rv.id]) + "\n";
                }
            }
            fv = fv.Trim('\n');
            return fv;
        }
        Hashtable revHash = new Hashtable();

        private RemoteIssue GetIssue(JiraSoapServiceService jss, string token, string jiraId)
        {
            if (loadedJiraItems.ContainsKey(jiraId))
            {
                return (RemoteIssue)(loadedJiraItems[jiraId]);
            }
            else
            {
                try
                {
                    RemoteIssue issue = jss.getIssue(token, jiraId);
                    loadedJiraItems[jiraId] = issue;
                    return issue;
                }
                catch(Exception exp)
                {
                    rp("Loading Jira issue failed..." + exp.Message);
                    return null;
                }
            }
        }
        private void WriteExcelRow(Excel.Worksheet oSheet, int index, reviewData rv, versionedLineCommentData vlcd, 
            JiraSoapServiceService jss, string token)
        {
            int column = 1;
            rp(String.Format("[{0}] Processing comment [{1}] from [{2}]",CommentCount,
                vlcd.permaId.id ,vlcd.user.displayName));
            CommentCount++;

            Excel.Range rng = (Excel.Range)(oSheet.Cells[index, 1]);
            rng.EntireRow.Font.Size = 9;
            rng.EntireRow.WrapText = true;

            oSheet.Cells[index, column++] = rv.author.userName;
            oSheet.Cells[index, column++] = rv.author.displayName;
            oSheet.Cells[index, column++] = bf(rv.closeDate);
            oSheet.Cells[index, column++] = bf(rv.createDate);
            oSheet.Cells[index, column++] = rv.creator.userName;
            oSheet.Cells[index, column++] = rv.creator.displayName;
            oSheet.Cells[index, column++] = rv.description;
            oSheet.Cells[index, column++] = bf(rv.dueDate);
            oSheet.Cells[index, column++] = rv.jiraIssueKey;
            if (rv.jiraIssueKey != null && rv.jiraIssueKey != "")
            {
                RemoteIssue issue = GetIssue(jss, token, rv.jiraIssueKey);
                oSheet.Cells[index, column++] = verToString(issue.affectsVersions);
                oSheet.Cells[index, column++] = GetCvValue(issue, Properties.Settings.Default.JiraAddlCustomField);
            }
            else
            {
                column++;
                column++;
            }
            oSheet.Cells[index, column++] = rv.moderator.userName;
            oSheet.Cells[index, column++] = rv.moderator.displayName;
            oSheet.Cells[index, column++] = rv.name;
            oSheet.Cells[index, column++] = rv.permaId.id;
            oSheet.Cells[index, column++] = rv.projectKey;
            oSheet.Cells[index, column++] = rv.state.ToString();
            oSheet.Cells[index, column++] = rv.summary;
            oSheet.Cells[index, column++] = vlcd.fromLineRange;
            oSheet.Cells[index, column++] = vlcd.toLineRange;
            if (vlcd.lineRanges != null)
            {
                string lineRanges = "";
                foreach (lineRangeDetail lr in vlcd.lineRanges)
                {
                    lineRanges += lr.revision + "-" + lr.range + "\r\n";
                }
                oSheet.Cells[index, column++] = lineRanges;
            }
            else
            {
                column++;
            }
            oSheet.Cells[index, column++] = bf(vlcd.createDate);
            oSheet.Cells[index, column++] = Convert.ToString(vlcd.defectRaised);
            oSheet.Cells[index, column++] = Convert.ToString(vlcd.deleted);
            oSheet.Cells[index, column++] = Convert.ToString(vlcd.draft);
            oSheet.Cells[index, column++] = Convert.ToString(vlcd.message);
            oSheet.Cells[index, column++] = vlcd.user.userName;
            oSheet.Cells[index, column++] = vlcd.user.displayName;
            accepted = false;
            if (vlcd.replies.Any != null)
            {
                oSheet.Cells[index, column++] = ProcessReplies(vlcd.replies.Any, 1);
            }
            else
            {
                column++;
            }
            
            oSheet.Cells[index, column++] = Convert.ToString(accepted);
        }
        private Hashtable loadedJiraItems = new Hashtable();
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                CommentCount = 1;

                rp("Logging in into Crucible...");

                Stream auth = getHttpStream(
                    String.Format(Properties.Settings.Default.CrucibleLoginUrl, Properties.Settings.Default.CrucibleUserName,
                    Properties.Settings.Default.CruciblePassword));
                XmlSerializer asr = new XmlSerializer(typeof(loginResult));
                loginResult lr = (loginResult)asr.Deserialize(auth);

                rp("Login complete...");

                rp("Fetching reviews...");
                Stream rvs = getHttpStream(String.Format(Properties.Settings.Default.CrucibleReviewsUrl,
                    Properties.Settings.Default.CrucibleProject, lr.token));

                XmlSerializer rsr = new XmlSerializer(typeof(reviews));
                reviews reviews = (reviews)rsr.Deserialize(rvs);

                rp("Opening Excel...");

                Excel.Application oXL = new Excel.Application();
                Excel.Workbook workBook = oXL.Workbooks.Add(System.Reflection.Missing.Value);

                rp("Preparing the workbook...");

                PrepareExcelWorkbook(workBook);
                Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Sheets[1];

                rp("Logging in into Jira...");
                JiraSoapServiceService jss = new JiraSoapServiceService();

                string token = jss.login(Properties.Settings.Default.CrucibleUserName, 
                    Properties.Settings.Default.CruciblePassword);

                rp("Loading Jira versions");
                foreach (RemoteVersion ver in jss.getVersions(token,Properties.Settings.Default.JiraProjectName))
                {
                    revHash[ver.id] = ver.name;
                    log(String.Format("Adding [{0}] : [{1}] to hash", ver.name, ver.id));
                }

                rp("Populating data into workbook...");
                int rowIndex = 2;

                foreach (reviewData rv in reviews.reviewData)
                {
                    if(
                        (Properties.Settings.Default.CrucibleFetchAllReviews == false) &&
                        ( rv.state != state.Closed )
                        )
                    {
                        log("Incomplete review, skipping " + rv.permaId.id);
                        continue;
                    }
                    else
                    {
                        rp("Processing " + rv.permaId.id);

                        Stream cms = getHttpStream(String.Format(Properties.Settings.Default.CrucibleCommentUrl, rv.permaId.id, lr.token));

                        XmlSerializer cmr = new XmlSerializer(typeof(comments));
                        comments rcomments = (comments)cmr.Deserialize(cms);
                        if (rcomments.Any != null)
                        {
                            foreach (System.Xml.XmlElement elem in rcomments.Any)
                            {
                                XmlSerializer vlcdr = new XmlSerializer(typeof(versionedLineCommentData));
                                versionedLineCommentData vlcd = (versionedLineCommentData)vlcdr.Deserialize(new StringReader(
                                    decorate(elem.InnerXml, "versionedLineCommentData")));
                                WriteExcelRow(workSheet, rowIndex, rv, vlcd, jss, token);
                                rowIndex++;
                            }
                        }
                        else
                        {
                            rp("Skipping " + rv.permaId.id + ". No review comments detected.");
                        }
                    }
                }
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception exp)
            {
                rp("Sorry, exception occured, after all this is software and there is no CI for this :)\r\n" +
                    exp.Message + "\r\n" + exp.StackTrace);
            }
            rp("Completed...");
        }

        public Stream getHttpStream(string url)
        {
            return getHttpStream(url, "", "");
        }
        public Stream getHttpStream(string url, string uname, string pwd)
        {
            Stream resStream = null;
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            string password = pwd;
            string username = uname;
            if (username.Length != 0)
            {
                string autorization = username + ":" + password;
                byte[] binaryAuthorization = System.Text.Encoding.UTF8.GetBytes(autorization);
                autorization = Convert.ToBase64String(binaryAuthorization);
                autorization = "Basic " + autorization;
                request.Headers.Add("AUTHORIZATION", autorization);
            }

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            // we will read data via the response stream
            resStream = response.GetResponseStream();
            return resStream;
        }

        public string getHttpData(string url, string uname, string pwd)
        {
            string tempString = null;
            // used to build entire input
            StringBuilder sb = new StringBuilder();

            // used on each read operation
            byte[] buf = new byte[8192];

            Stream resStream = getHttpStream(url, uname, pwd);
            int count = 0;

            do
            {
                // fill the buffer with data
                count = resStream.Read(buf, 0, buf.Length);

                // make sure we read some data
                if (count != 0)
                {
                    // translate from bytes to ASCII text
                    tempString = Encoding.ASCII.GetString(buf, 0, count);

                    // continue building the string
                    sb.Append(tempString);
                }
            }
            while (count > 0); // any more data to read?
            return sb.ToString();
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string msg = "[" + DateTime.Now.ToLongTimeString() + "] " + (string)(e.UserState);
            log(msg);
            textBox.Text +=  msg + "\r\n";
            textBox.SelectionStart = textBox.Text.Length;
            textBox.ScrollToCaret();
            textBox.Refresh();
        }

        private void buttonDump_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.CruciblePassword.Length == 0 || Properties.Settings.Default.CrucibleUserName.Length == 0)
            {
                MessageBox.Show("Please check the settings for Crucible username/password", "Can't proceed", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                tabControl1.SelectedIndex = 1;
                return;
            }
            if (!backgroundWorker.IsBusy)
            {
                backgroundWorker.RunWorkerAsync();
                tabControl1.SelectedIndex = 2;
            }
            else
            {
                log("Already running");
            }
        }

        private void propertyGrid1_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(Properties.Settings.Default.CrucibleRESTAPIGuide);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("mailto:renjith.v@nsn.com?subject=Crucible Data Dump Support");
        }
    }
}
