using AuditCodeCoreMessage;
using LeS.FileStore;
using LeSCommonRoutines;
using System.Net;
using HtmlAgilityPack;
using LES.MTML.Generator;
using LES.MTML;
using System.Web;
using System.Globalization;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;
using System.Reflection;


namespace Http_Navibulgar_IProcessor
{
    public class Navibulgar_Routine : LeSCommon
    {
        public readonly LeSFileStore _fileStore;

        #region variables

        public int iRetry = 0, IsAltItemAllowed = 0, IsPriceAveraged = 0, IsUOMChanged = 0, nFileCount = 0, LoginRetry = 0;

        public string sAuditMesage = "", ImgPath = "", VRNO = "", sDoneFile = "", Vessel = "", Port = "", BuyerName = "", SupplierName = "",
            MTML_QuotePath = "", MessageNumber = "", LeadDays = "", Currency = "", MsgNumber = "", MsgRefNumber = "", UCRefNo = "", AAGRefNo = "", cConfigUserId = ""
            , LesRecordID = "", BuyerPhone = "", BuyerEmail = "", BuyerFax = "", supplierName = "", supplierPhone = "", supplierEmail = "", supplierFax = "", cSiteURL = "",
             VesselName = "", PortName = "", PortCode = "", SupplierComment = "", PayTerms = "", PackingCost = "", FreightCharge = "", GrandTotal = "", Allowance = "", TaxCost = "",
          ProcessorName = "", TotalLineItemsAmount = "", BuyerTotal = "", DtDelvDate = "", dtExpDate = "", ToDate = "", FromDate = "", PODate = "", cFilterDuration = "", RFQPath = "",
             MailTo = "", MailBcc = "", MailCC = "", NotificationPath = "", ToCheckPath = "", cPOCPath = "", cPOCurrency = "", MODULE_NAME = "", currDocType = "", IMAGE_PATH = "",
            MAIL_TEXT_PATH = "", AUDITPATH = "", SERVER = "", LINKID, cBranchList;

        public string BUYER_LINK_CODE = "", SUPPLIER_LINK_CODE = "";
        DateTime dt = new DateTime(); DateTime dt1 = new DateTime(); string _dateTo = "", _dateFrom = "";
        bool IsDecline = false, IsSaveQuote = false, IsSubmitQuote = false, SendMail = false, IsSubmitQuote_Buyer = false, IsDownloadFromPortal = false;
        public string[] Actions;
        List<string> xmlFiles = new List<string>();
        List<string> slPendingFiles = new List<string>();
        public LineItemCollection _lineitem = null;
        //RichTextBox _txtData = new RichTextBox();
        public MTMLInterchange _interchange { get; set; }
        Dictionary<string, string> dctSubmitValues = new Dictionary<string, string>();


        #endregion variables


        public Navibulgar_Routine(LeSFileStore fileStore)
        {
            try
            {
                _fileStore = fileStore;
                Initialise();
            }
            catch (Exception ex)
            {
                LogText = "Exception occured in initializing " + ex.GetBaseException;
                throw;
            }
        }

        public void StartingProcess()
        {
            LogText = "Starting the Process";

            try
            {

                foreach (Dictionary<string, string> _dctAppsettings in dctAppSettingsList)
                {

                    dctAppSettings = _dctAppsettings;
                    LoadAppSettings();
                    ProcessUser();
                    LogText = "Process started for " + Userid;

                    LogText = "Uploading files to cloud";
                    //UploadOnCloud("Screenshots", UploadFileType.Attachments, ImgPath);
                    //UploadOnCloud("MTML files", UploadFileType.MtmlInbox, RFQPath);
                    UploadAudits();
                    LogText = "Uploading files to cloud completed";
                }
            }
            catch (Exception e)
            {
                LogText = "Unexpected exception in StartingProcess : " + e.Message;
                AuditMessageData.CreateAuditFile("", MODULE_NAME, "", AuditMessageData.UNEXPECTED_ERROR, BuyerCode, SupplierCode, "", e.Message);
            }
            finally
            {
                LogText = $"{MODULE_NAME} completed!";
                //UploadLogFile();

            }
        }

        private void UploadLogFile()
        {
            try
            {
                if (_fileStore != null)
                {
                    string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
                    string path = Assembly.GetEntryAssembly().Location;

                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                    string path2 = Path.Combine(LogPath, fileNameWithoutExtension + "_" + DateTime.Now.ToString("yyyy") + "_" + DateTime.Now.ToString("MM") + "_" + DateTime.Now.ToString("dd") + ".log");

                    var logRes = _fileStore.UploadLogFile(path2);
                    //AuditMessageData.LogText = "Logs uploaded successfully.";
                }
            }
            catch (Exception e)
            {
                LogText = $"Stack Trace + {e.StackTrace} ";

                AuditMessageData.LogText = "Exception in UploadLogs : " + e.GetBaseException().ToString();

                AuditMessageData.CreateAuditFile("", MODULE_NAME, "", AuditMessageData.UNEXPECTED_ERROR, BuyerCode, SupplierCode, this.currDocType, e.Message);
            }
        }

        private void UploadOnCloud(string type, UploadFileType uploadFileType, string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    int fileCount = Directory.GetFiles(path).Length;
                    if (fileCount > 0)
                    {
                        var validpath = Path.GetFullPath(path);
                        var result = _fileStore.UploadFilesBulk(uploadFileType, validpath);
                        if (result)
                        {
                            LogText = $"{type} uploaded to cloud successfully.";
                        }
                        else
                        {
                            // ## DONE -- Can you pass Supplier Code & Buyer Code in Audit Log here 
                            AuditMessageData.CreateAuditFile("", MODULE_NAME, "", AuditMessageData.UNEXPECTED_ERROR, BuyerCode, SupplierCode, this.currDocType, $"Unable to upload {type} to cloud.");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                LogText = $"Stack Trace + {e.StackTrace} ";

                LogText = "Exception in UploadOnCloud : " + e.GetBaseException().ToString();
                AuditMessageData.CreateAuditFile("", MODULE_NAME, "", AuditMessageData.UNEXPECTED_ERROR, BuyerCode, SupplierCode, this.currDocType, $"Unable to upload {type} to cloud.");
            }
        }

        private void UploadAudits()
        {
            try
            {
                // Updated By Sanjita //
                if (Path.Exists(AuditPath))
                {
                    bool uploaded = _fileStore.UploadFilesBulk(UploadFileType.Audit, AuditPath);
                    if (uploaded)
                        LogText = "AuditLog Files uploaded successfully.";
                    //else LogText = "Unable to upload Audit Log files";
                }

                if (Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AuditMessageData.DefaultAuditPath)))
                {
                    bool uploaded = _fileStore.UploadFilesBulk(UploadFileType.Audit, AuditMessageData.DefaultAuditPath); //upload offline audit logs
                    if (uploaded)
                        LogText = "eSupplierAudit files uploaded successfully.";
                    //else LogText = "Unable to upload Audit Log files";
                }
            }
            catch (Exception e)
            {
                LogText = $"Stack Trace + {e.StackTrace} ";
                LogText = "Exception in UploadAudits : " + e.GetBaseException().ToString();
                // ## DONE -- Can you pass Supplier Code & Buyer Code in Audit Log here 
                // ## DONE -- Also append e.Message in AdditionalMessage 
                AuditMessageData.CreateAuditFile("", MODULE_NAME, "", AuditMessageData.UNEXPECTED_ERROR, BuyerCode, SupplierCode, this.currDocType, e.Message);
            }
        }

        public void LoadAppSettings()
        {
            try
            {
                LogText = ("Loading AppSettings");
                if (dctAppSettings.Count > 0)
                {
                    iRetry = 0;
                    if (dctAppSettings.ContainsKey("LINKID")) LINKID = dctAppSettings["LINKID"].Trim();

                    if (dctAppSettings.ContainsKey("SITE_URL") || string.IsNullOrWhiteSpace(dctAppSettings["SITE_URL"]))
                        cSiteURL = URL = dctAppSettings["SITE_URL"].Trim();

                    if (dctAppSettings.ContainsKey("USERNAME") || string.IsNullOrWhiteSpace(dctAppSettings["USERNAME"]))
                        cConfigUserId = Userid = dctAppSettings["USERNAME"].Trim();

                    if (dctAppSettings.ContainsKey("PASSWORD") || string.IsNullOrWhiteSpace(dctAppSettings["PASSWORD"]))
                        Password = dctAppSettings["PASSWORD"].Trim();

                    if (dctAppSettings.ContainsKey("DOMAIN") || string.IsNullOrWhiteSpace(dctAppSettings["DOMAIN"]))
                        Domain = dctAppSettings["DOMAIN"].Trim();

                    if (dctAppSettings.ContainsKey("MODULE_NAME") || string.IsNullOrWhiteSpace(dctAppSettings["MODULE_NAME"]))
                        MODULE_NAME = dctAppSettings["MODULE_NAME"].Trim();

                    if (dctAppSettings.ContainsKey("BUYERCODE") || string.IsNullOrWhiteSpace(dctAppSettings["BUYERCODE"]))
                        BuyerCode = dctAppSettings["BUYERCODE"].Trim();

                    if (dctAppSettings.ContainsKey("SUPPLIERCODE") || string.IsNullOrWhiteSpace(dctAppSettings["SUPPLIERCODE"]))
                        SupplierCode = dctAppSettings["SUPPLIERCODE"].Trim();

                    if (dctAppSettings.ContainsKey("BUYERNAME") || string.IsNullOrWhiteSpace(dctAppSettings["BUYERNAME"]))
                        BuyerName = dctAppSettings["BUYERNAME"].Trim();

                    if (dctAppSettings.ContainsKey("SUPPLIERNAME") || string.IsNullOrWhiteSpace(dctAppSettings["SUPPLIERNAME"]))
                        SupplierName = dctAppSettings["SUPPLIERNAME"].Trim();

                    if (dctAppSettings.ContainsKey("ACTIONS") || string.IsNullOrWhiteSpace(dctAppSettings["ACTIONS"]))
                        Actions = dctAppSettings["ACTIONS"].Trim().Split(',');

                    if (dctAppSettings.ContainsKey("ATTACHMENT_PATH"))
                    {
                        ImgPath = dctAppSettings["ATTACHMENT_PATH"].Trim();
                        if (!Directory.Exists(ImgPath)) Directory.CreateDirectory(ImgPath);
                    }
                    if (dctAppSettings.ContainsKey("MAIL_TEXT_PATH"))
                    {
                        MAIL_TEXT_PATH = dctAppSettings["MAIL_TEXT_PATH"].Trim();
                        if (!Directory.Exists(MAIL_TEXT_PATH)) Directory.CreateDirectory(MAIL_TEXT_PATH);
                    }
                    if (dctAppSettings.ContainsKey("AUDITPATH"))
                    {
                        AuditPath = dctAppSettings["AUDITPATH"].Trim();
                        if (!Directory.Exists(AuditPath)) Directory.CreateDirectory(AuditPath);
                    }

                    if (dctAppSettings.ContainsKey("BUYER_LINK_CODE"))
                        BUYER_LINK_CODE = dctAppSettings["BUYER_LINK_CODE"].Trim();

                    if (dctAppSettings.ContainsKey("SUPPLIER_LINK_CODE"))
                        SUPPLIER_LINK_CODE = dctAppSettings["SUPPLIER_LINK_CODE"].Trim();

                    if (dctAppSettings.ContainsKey("SERVER"))
                        SERVER = dctAppSettings["SERVER"].Trim();


                    if (dctAppSettings.ContainsKey("QUOTE_POC_PATH"))
                    {
                        MTML_QuotePath = dctAppSettings["QUOTE_POC_PATH"].Trim();
                        if (!Directory.Exists(MTML_QuotePath)) Directory.CreateDirectory(MTML_QuotePath);
                    }

                    if (!dctAppSettings.ContainsKey("LOGINRETRY") ||
                        !int.TryParse(dctAppSettings["LOGINRETRY"].Trim(), out LoginRetry))
                        throw new Exception("LOGINRETRY missing or invalid");

                    if (dctAppSettings.ContainsKey("INQUIRY_FROM_DATE"))
                        FromDate = dctAppSettings["INQUIRY_FROM_DATE"];

                    if (dctAppSettings.ContainsKey("INQUIRY_TO_DATE"))
                        ToDate = dctAppSettings["INQUIRY_TO_DATE"].Trim();

                    if (dctAppSettings.ContainsKey("SAVE_QUOTE"))
                    {
                        IsSaveQuote = dctAppSettings["SAVE_QUOTE"].Trim().ToUpper() == "YES";
                    }

                    if (dctAppSettings.ContainsKey("SUBMIT_QUOTE"))
                    {
                        IsSubmitQuote = dctAppSettings["SUBMIT_QUOTE"].Trim().ToUpper() == "YES";
                    }

                    // Added By Sanjita //
                    if (dctAppSettings.ContainsKey("SEND_MAIL_NOTIFICATION"))
                    {
                        SendMail = dctAppSettings["SEND_MAIL_NOTIFICATION"].Trim().ToUpper() == "YES";
                    }

                    if (dctAppSettings.ContainsKey("SEND_MAIL_PATH"))
                    {
                        NotificationPath = convert.ToString(dctAppSettings["MAIL_NOTIFICATION_PATH"]).Trim();
                    }

                    if (dctAppSettings.ContainsKey("BRANCH_LIST"))
                    {
                        cBranchList = convert.ToString(dctAppSettings["BRANCH_LIST"]).Trim();
                    }




                    if (!Directory.Exists(ImgPath)) Directory.CreateDirectory(ImgPath);
                    //if (!Directory.Exists(MTML_QuotePath + "\\Backup")) Directory.CreateDirectory(MTML_QuotePath + "\\Backup");
                    //if (!Directory.Exists(MTML_QuotePath + "\\Error")) Directory.CreateDirectory(MTML_QuotePath + "\\Error");

                    //ToCheckPath = Path.Combine(MTML_QuotePath , "TO_CHECK");
                    //if (!Directory.Exists(ToCheckPath)) Directory.CreateDirectory(ToCheckPath);

                    if (dctAppSettings.ContainsKey("POC_PATH"))
                    {
                        cPOCPath = dctAppSettings["POC_PATH"].Trim();
                        if (!Directory.Exists(cPOCPath)) Directory.CreateDirectory(cPOCPath);
                        if (!Directory.Exists(cPOCPath + "\\Backup")) Directory.CreateDirectory(cPOCPath + "\\Backup");
                        if (!Directory.Exists(cPOCPath + "\\Error")) Directory.CreateDirectory(cPOCPath + "\\Error");
                    }

                    if (dctAppSettings.ContainsKey("SUBMIT_QUOTE_BUYER"))
                    {
                        IsSubmitQuote_Buyer = dctAppSettings["SUBMIT_QUOTE_BUYER"].Trim().ToUpper() == "YES";
                    }
                    else IsSubmitQuote_Buyer = false;

                    if (dctAppSettings.ContainsKey("FILTER_DURATION"))
                    {
                        cFilterDuration = convert.ToString(dctAppSettings["FILTER_DURATION"]);
                    }

                    //Convert.ToString(ConfigurationManager.AppSettings["FILTER_DURATION"]);

                    if (dctAppSettings.ContainsKey("MTML_PATH"))
                    {
                        RFQPath = dctAppSettings["MTML_PATH"].Trim();
                        if (!Directory.Exists(RFQPath)) Directory.CreateDirectory(RFQPath);
                        //    if (!Directory.Exists(RFQPath + "\\Backup")) Directory.CreateDirectory(RFQPath + "\\Backup");
                        //    if (!Directory.Exists(RFQPath + "\\Error")) Directory.CreateDirectory(RFQPath + "\\Error");
                    }

                    if (dctAppSettings.ContainsKey("DOWNLOAD_FROM_PORTAL"))
                    {
                        if (dctAppSettings["DOWNLOAD_FROM_PORTAL"] == "TRUE")
                        {
                            IsDownloadFromPortal = true;
                        }
                    }

                    if (dctAppSettings.ContainsKey("PROCESSOR_NAME"))
                    {
                        ProcessorName = convert.ToString(dctAppSettings["PROCESSOR_NAME"]);
                    }


                    if (dctAppSettings.ContainsKey("VALIDATE_CERTIFICATE") && dctAppSettings["VALIDATE_CERTIFICATE"].Trim().ToUpper() == "FALSE")
                    {
                        ServicePointManager.ServerCertificateValidationCallback =
                            delegate { return true; };
                    }

                    _fileStore.CreateAppendTextFile($"DownloadedPO_{cBranchList}.txt");
                    _fileStore.DownloadFile($"DownloadedPO_{cBranchList}.txt", DownloadFileType.Other, FileMoveAction.None, dctAppSettings, LINKID);

                    _fileStore.CreateAppendTextFile($"DownloadedRFQ_{cBranchList}.txt");
                    _fileStore.DownloadFile($"DownloadedRFQ_{cBranchList}.txt", DownloadFileType.Other, FileMoveAction.None, dctAppSettings, LINKID);



                    LogText = "Loading AppSetting Completed";
                }
                else
                {
                    LogText = "No App Setting Found";
                    throw new Exception("No App Setting Found");
                }
            }
            catch (Exception ex)
            {
                LogText = "Error in Load App Settings - " + ex.Message;
                throw new Exception("Error in Load App Settings - " + ex.Message);
            }
        }

        public bool ProcessUser()
        {
            bool _result = true;
            if (IsDownloadFromPortal)
            {
                LogText = "Domain: " + Domain + Environment.NewLine;
                LogText = "Process started for " + Userid;
                if (DoLogin("input", "id", "MainContent_txtFrom"))
                {
                    foreach (string sAction in Actions)
                    {
                        try
                        {
                            switch (sAction.ToUpper())
                            {
                                case "RFQ":
                                    currDocType = "RFQ";
                                    ProcessRFQ(string.Empty);
                                    break;
                                case "QUOTE":
                                    currDocType = "QUOTE";
                                    ProcessQuote();
                                    break;
                                case "PO":
                                    currDocType = "PO";
                                    ProcessPO();
                                    break;
                                case "POC":
                                    currDocType = "POC";
                                    //ProcessPOC();
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogText = "Exception while processing user actions " + sAction.ToUpper() + " : " + ex.GetBaseException().ToString();
                        }
                    }
                }
            }
            else
            {
                foreach (string sAction in Actions)
                {
                    try
                    {
                        switch (sAction.ToUpper())
                        {
                            case "RFQ":
                            case "PO":
                                ProcessMailFiles();
                                break;
                            case "QUOTE":
                                //ProcessQuote();
                                break;
                            case "POC":
                                //ProcessPOC();
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogText = "Exception while processing user actions " + sAction.ToUpper() + " : " + ex.GetBaseException().ToString();
                    }
                }
            }
            return _result;
        }

        public override bool DoLogin(string validateNodeType, string validateAttribute, string attributeValue, bool bload = true)
        {
            bool isLoggedin = false;
            try
            {
                //  Login();
                URL = cSiteURL;
                LoadURL("input", "id", "MainContent_LoginUser_LoginButton");
                if (_httpWrapper._dctStateData.Count > 0)
                {
                    dctPostDataValues.Clear();
                    dctPostDataValues.Add("__EVENTTARGET", _httpWrapper._dctStateData["__EVENTTARGET"]);
                    dctPostDataValues.Add("__VIEWSTATE", _httpWrapper._dctStateData["__VIEWSTATE"]);
                    dctPostDataValues.Add("__VIEWSTATEGENERATOR", _httpWrapper._dctStateData["__VIEWSTATEGENERATOR"]);
                    dctPostDataValues.Add("__EVENTVALIDATION", _httpWrapper._dctStateData["__EVENTVALIDATION"]);
                    dctPostDataValues.Add("ctl00%24MainContent%24LoginUser%24UserName", HttpUtility.UrlEncode(Userid));
                    dctPostDataValues.Add("ctl00%24MainContent%24LoginUser%24Password", HttpUtility.UrlEncode(Password));
                    dctPostDataValues.Add("ctl00%24MainContent%24LoginUser%24LoginButton", "Log+In");
                    isLoggedin = base.DoLogin(validateNodeType, validateAttribute, attributeValue, false);

                    if (isLoggedin)
                    {
                        LogText = "Logged in successfully";
                    }
                    else
                    {
                        if (iRetry == LoginRetry)
                        {
                            string filename = ImgPath + "\\Navibulgar_LoginFailed_" + DateTime.Now.ToString("ddMMyyyHHmmssfff") + "_" + Domain + ".png";
                            if (!PrintScreen(filename)) filename = "";
                            LogText = "Login failed";
                            //CreateAuditFile(filename, "Navibulgar_HTTP_Processor", "", "Error", "LeS-1014:Unable to login.", BuyerCode, SupplierCode, AuditPath);
                        }
                        else
                        {
                            iRetry++;
                            LogText = "Login retry";
                            isLoggedin = DoLogin(validateNodeType, validateAttribute, attributeValue, false);
                        }
                    }
                }
                else
                {
                    string filename = ImgPath + "\\Navibulgar_LoginFailed_" + DateTime.Now.ToString("ddMMyyyHHmmssfff") + "_" + Domain + ".png";
                    if (!PrintScreen(filename)) filename = "";
                    LogText = "Unable to load URL" + URL;//
                    //CreateAuditFile(filename, "Navibulgar_HTTP_Processor", "", "Error", "LeS-1016:Unable to load URL" + URL, BuyerCode, SupplierCode, AuditPath);
                }
            }
            catch (Exception e)
            {
                LogText = "Exception while login : " + e.GetBaseException().ToString();
                if (iRetry > LoginRetry)
                {
                    string filename = ImgPath + "\\Navibulgar_LoginFailed_" + DateTime.Now.ToString("ddMMyyyHHmmssfff") + "_" + Domain + ".png";
                    if (!PrintScreen(filename)) filename = "";
                    LogText = "Login failed";
                    //CreateAuditFile(filename, "Navibulgar_HTTP_Processor", "", "Error", "LeS-1014:Unable to login. Error : " + e.Message, BuyerCode, SupplierCode, AuditPath);
                }
                else
                {
                    iRetry++;
                    LogText = "Login retry";
                    isLoggedin = DoLogin(validateNodeType, validateAttribute, attributeValue);
                }
            }
            return isLoggedin;
        }

        public void ProcessMailFiles()
        {
            try
            {
                LogText = "Mail text files Process started";
                if (!string.IsNullOrWhiteSpace(RFQPath))
                {
                    DirectoryInfo _dir = new DirectoryInfo(RFQPath);
                    FileInfo[] _Files = _dir.GetFiles("*.txt");

                    if (_Files != null && _Files.Length > 0)
                    {
                        foreach (FileInfo fInfo in _Files)
                        {
                            //_txtData.Text = File.ReadAllText(fInfo.FullName);

                            //if (_txtData.Text.ToUpper().Contains("REQUEST FOR QUOTATION") || _txtData.Text.ToUpper().Contains("RFQ") || _txtData.Text.ToUpper().Contains("INQUIRY"))
                            //    GetDetails(fInfo.FullName, _txtData);
                            //else
                            //{
                            //    LogText = "Invalid mail file " + fInfo.Name;
                            //    File.Move(fInfo.FullName, RFQPath + "\\Error\\" + fInfo.Name);
                            //    CreateAuditFile(fInfo.FullName, ProcessorName, "", "Error", "Invalid mail file " + fInfo.Name, BuyerCode, SupplierCode, AuditPath);
                            //}
                        }
                    }
                    else LogText = "No Files present.";
                }
                else LogText = "RFQ Path is empty";
            }
            finally
            {
                LogText = "Mail text files Process ended";
            }
        }

        //private void GetDetails(string emlFile, RichTextBox _txtData)
        //{
        //    string cRefNo = "";
        //    for (int i = 0; i < _txtData.Lines.Length; i++)
        //    {
        //        string line = _txtData.Lines[i];
        //        if (line.ToUpper().Contains("USERNAME"))
        //        {
        //            int indx = line.IndexOf("Username:");
        //            Userid = line.Substring(indx).Replace("Username:", "").Replace("<mailto:wris@wrist.com>", "").Trim();
        //        }
        //        else if (line.ToUpper().Contains("PASSWORD"))
        //        {
        //            int indx = line.IndexOf("Password:");
        //            Password = line.Substring(indx).Replace("Password:", "").Trim();
        //        }
        //        else if (line.Contains("Request for quotation"))
        //        {
        //            int indx = line.IndexOf("Request for quotation:");
        //            cRefNo = line.Substring(indx).Replace("Request for quotation:", "").Trim();
        //        }
        //    }
        //    if (!string.IsNullOrEmpty(Userid) && !string.IsNullOrEmpty(Password) && (cConfigUserId == Userid))
        //    {
        //        if (DoLogin("input", "id", "MainContent_txtFrom"))
        //        {
        //            //ProcessRFQ(cRefNo, emlFile);
        //        }
        //    }
        //    else
        //    {
        //        LogText = "UserId or Password not found in file.";
        //        File.Move(emlFile, RFQPath + "\\Error\\" + Path.GetFileName(emlFile));
        //        //CreateAuditFile(emlFile, ProcessorName, "", "Error", "LeS-1014.1:Unable to login as UserName/Password not found in file", BuyerCode, SupplierCode, AuditPath);
        //    }
        //}


        #region ## RFQ ##

        public void ProcessRFQ(string cMailRefNo, string emlFile)
        {
            try
            {
                this.currDocType = "RFQ";
                LogText = "RFQ processing started.";
                List<string> _lstNewRFQs = GetNewRFQs(cMailRefNo);
                if (_lstNewRFQs.Count > 0)
                {
                    //DownloadRFQ(_lstNewRFQs);
                }
                else
                {
                    if (string.IsNullOrEmpty(cMailRefNo)) LogText = "No new RFQ found.";
                    else
                    {
                        LogText = cMailRefNo + " not found in the filtered records.";
                        //CreateAuditFile(emlFile, "Navibulgar_HTTP_Processor", cMailRefNo, "Error", "LeS-1006:Unable to filter " + cMailRefNo + " as it is not found in the filtered records.", BuyerCode, SupplierCode, AuditPath);
                        File.Move(emlFile, RFQPath + "\\Error\\" + Path.GetFileName(emlFile));
                    }
                }
                LogText = "RFQ processing stopped.";
            }
            catch (Exception e)
            {
                //WriteErrorLog_With_Screenshot("Exception in Process RFQ : " + e.GetBaseException().ToString());
                //WriteErrorLog_With_Screenshot("Unable to process file due to " + e.GetBaseException().ToString(), "LeS-1004:");
            }
        }

        public void ProcessRFQ(string cMailRefNo)
        {
            try
            {
                this.currDocType = "RFQ";
                LogText = "RFQ processing started.";
                List<string> _lstNewRFQs = GetNewRFQs(cMailRefNo);
                if (_lstNewRFQs.Count > 0)
                {
                    DownloadRFQ(_lstNewRFQs);
                }
                else
                {
                    LogText = "No new RFQ found.";
                }
                LogText = "RFQ processing stopped.";
            }
            catch (Exception e)
            {
                LogText = "Exception in ProcessRFQ : " + e.GetBaseException().ToString();
                //WriteErrorLog_With_Screenshot("Exception in Process RFQ : " + e.GetBaseException().ToString());
                //WriteErrorLog_With_Screenshot("Unable to process file due to " + e.GetBaseException().ToString(), "LeS-1004:");
            }
        }

        public List<string> GetNewRFQs(string cRefNo)
        {
            List<string> _lstNewRFQs = new List<string>();
            List<string> slProcessedItem = GetProcessedItems(eActions.RFQ);
            _lstNewRFQs.Clear();
            _httpWrapper._CurrentDocument.LoadHtml(_httpWrapper._CurrentResponseString);

            #region for filter inqiry table by dates
            DateTime dtTo = new DateTime();
            DateTime dtFrom = new DateTime();
            if (ToDate == "" && FromDate == "")
            {
                dtTo = DateTime.Now;
                if (cFilterDuration == "DAYS") dtFrom = DateTime.Now.AddDays(-1);//AddDays(-1);[changed by kalpita on 12/07/2021]
                else if (cFilterDuration == "MONTHS") dtFrom = DateTime.Now.AddMonths(-1);//[changed by kalpita on 12/07/2021]
            }
            else
            {
                dtTo = DateTime.MinValue;
                DateTime.TryParseExact(ToDate, "d/M/yyyy", null, DateTimeStyles.None, out dtTo);

                dtFrom = DateTime.MinValue;
                DateTime.TryParseExact(FromDate, "d/M/yyyy", null, DateTimeStyles.None, out dtFrom);
            }

            if (dtTo != DateTime.MinValue)
            {
                dt = DateTime.ParseExact(dtTo.ToShortDateString(), "M/d/yyyy", CultureInfo.InvariantCulture);
                _dateTo = dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            if (dtFrom != DateTime.MinValue)
            {
                dt1 = DateTime.ParseExact(dtFrom.ToShortDateString(), "M/d/yyyy", CultureInfo.InvariantCulture);
                _dateFrom = dt1.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }

            if (_dateTo != "" && _dateFrom != "")
            {
                if (_httpWrapper._dctStateData.Count > 0)
                {
                    dctPostDataValues.Clear();
                    dctPostDataValues.Add("__EVENTTARGET", "ctl00%24MainContent%24txtFrom");//
                    dctPostDataValues.Add("__EVENTARGUMENT", _httpWrapper._dctStateData["__EVENTARGUMENT"]);//
                    dctPostDataValues.Add("__LASTFOCUS", "");
                    dctPostDataValues.Add("__VIEWSTATE", _httpWrapper._dctStateData["__VIEWSTATE"]);
                    dctPostDataValues.Add("__VIEWSTATEGENERATOR", _httpWrapper._dctStateData["__VIEWSTATEGENERATOR"]);
                    dctPostDataValues.Add("__VIEWSTATEENCRYPTED", "");
                    dctPostDataValues.Add("__EVENTVALIDATION", _httpWrapper._dctStateData["__EVENTVALIDATION"]);
                    dctPostDataValues.Add("ctl00%24MainContent%24txtFrom", Uri.EscapeDataString(_dateFrom));
                    dctPostDataValues.Add("ctl00%24MainContent%24txtTo", Uri.EscapeDataString(_dateTo));
                    dctPostDataValues.Add("ctl00%24MainContent%24Button1", "Refresh");

                    if (!_httpWrapper._AddRequestHeaders.ContainsKey("Origin")) _httpWrapper._AddRequestHeaders.Add("Origin", @$"{cSiteURL}");
                    _httpWrapper.Referrer = "";
                }
                #endregion

                URL = $"{cSiteURL}suppliers/OutReq.aspx";
                if (PostURL("table", "id", "MainContent_GridView1"))
                {
                    HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@id='MainContent_GridView1']//tr[@onmouseout]");
                    if (_nodes != null)
                    {
                        if (_nodes.Count > 0)
                        {
                            foreach (HtmlNode _row in _nodes)
                            {
                                HtmlNodeCollection _rowData = _row.ChildNodes;
                                string VRNo = _rowData[1].InnerText.Trim();
                                string Vessel = _rowData[3].InnerText.Trim();
                                string Port = _rowData[4].InnerText.Trim();
                                string _url = _row.GetAttributeValue("onclick", "").Trim();
                                if (_url.Contains(';'))
                                {
                                    string[] _arrUrl = _url.Split(';');
                                    _url = _arrUrl[1].Replace("&#39", "");
                                }

                                string _guid = VRNo + "|" + Vessel + "|" + Port;

                                if (!_lstNewRFQs.Contains(VRNo + "|" + Vessel + "|" + Port + "|" + _url) && !slProcessedItem.Contains(_guid))
                                {
                                    if (IsDownloadFromPortal)//changed by kalpita on 16/07/2021
                                    {
                                        _lstNewRFQs.Add(VRNo + "|" + Vessel + "|" + Port + "|" + _url);
                                    }
                                    else
                                    {
                                        if (VRNo == cRefNo)
                                        {
                                            _lstNewRFQs.Add(VRNo + "|" + Vessel + "|" + Port + "|" + _url);
                                        }
                                    }
                                }
                            }
                        }
                        else
                            LogText = "No new RFQs found.";
                    }
                }
            }
            else
            {
                LogText = "To Date or/and From Date formats are wrong.";
                //CreateAuditFile("", "Navibulgar_HTTP_Processor", "", "Error", "leS-1023:Invalid date format", BuyerCode, SupplierCode, AuditPath);
                AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "leS-1023:Invalid date format");

            }
            return _lstNewRFQs;
        }

        public List<string> GetProcessedItems(eActions eAction)
        {
            try
            {
                List<string> slProcessedItems = new List<string>();
                switch (eAction)
                {
                    case eActions.RFQ: sDoneFile = AppDomain.CurrentDomain.BaseDirectory + this.Domain + "_" + this.SupplierCode + "_GUID.txt"; ; break;
                    case eActions.PO: sDoneFile = AppDomain.CurrentDomain.BaseDirectory + this.Domain + "_PO_" + this.SupplierCode + "_GUID.txt"; ; break;//13/07/2020
                                                                                                                                                          //case eActions.POC: sDoneFile = AppDomain.CurrentDomain.BaseDirectory + this.Domain + "_POC_" + this.SupplierCode + "_GUID.txt"; ; break;//13/07/2020
                    default: break;
                }
                if (File.Exists(sDoneFile))
                {
                    string[] _Items = File.ReadAllLines(sDoneFile);
                    slProcessedItems.AddRange(_Items.ToList());
                }
                return slProcessedItems;
            }
            catch (Exception ex)
            {
                LogText = ex.GetBaseException().ToString();
                throw;
            }
        }

        public void DownloadRFQ(List<string> _lstNewRFQs)
        {
            foreach (string strRFQ in _lstNewRFQs)
            {
                try
                {
                    string[] lst = strRFQ.Split('|');
                    //this can be used to create and maintan the download list.

                    this.VRNO = lst[0];
                    this.Vessel = lst[1];
                    this.Port = lst[2];
                    URL = lst[3];

                    URL = $"{cSiteURL}suppliers/" + URL;
                    LogText = "Processing RFQ for ref no " + this.VRNO;

                    LoadURL("input", "id", "MainContent_btnUpdate");

                    string eFile = this.ImgPath + "\\" + this.VRNO.Replace("/", "_") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + SupplierCode + ".png";
                    if (!PrintScreen(eFile)) eFile = "";
                    MTMLInterchange interchange = new MTMLInterchange();
                    DocHeader docHeader = new DocHeader();

                    interchange.PreparationDate = DateTime.Now.ToString("yyyy-MMM-dd");
                    interchange.PreparationTime = DateTime.Now.ToString("HH:mm");
                    interchange.ControlReference = DateTime.Now.ToString("yyyyMMddHHmmss");
                    interchange.Identifier = DateTime.Now.ToString("yyyyMMddHHmmss");
                    interchange.Recipient = SUPPLIER_LINK_CODE;
                    interchange.Sender = BUYER_LINK_CODE;
                    LineItemCollection lineItems = new LineItemCollection();
                    PartyCollection parties = new PartyCollection();


                    if (GetRFQHeader(ref docHeader, eFile))
                    {
                        interchange.DocumentHeader.MessageNumber = docHeader.References[0].ReferenceNumber;
                        interchange.DocumentHeader = docHeader;
                        if (GetRFQItems(ref lineItems))
                        {
                            docHeader.LineItems = lineItems;
                            if (GetAddress(ref parties))
                            {
                                docHeader.PartyAddresses = parties;
                                string fileName = "RFQ_" + this.VRNO.Replace("/", "_") + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xml";
                                string filePath = Path.Combine(RFQPath, fileName);
                                if (lineItems.Count > 0)
                                {
                                    interchange.DocumentHeader = docHeader;
                                    //_lesXml.FileName = xmlfile;
                                    MTMLClass _class = new MTMLClass();
                                    LogText = "Creating MTML RFQ";
                                    _class.Create(interchange, filePath);
                                    AuditMessageData.CreateAuditFile(fileName, MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "RFQ Generated Successfully for " + UCRefNo);
                                    _fileStore.AppendText($"DownloadedRFQ_{cBranchList}.txt", UCRefNo + "|" + this.Vessel + "|" + this.Port + Environment.NewLine);
                                    //if (File.Exists(_lesXml.FilePath + "\\" + _lesXml.FileName))
                                    //{
                                    //    LogText = xmlfile + " downloaded successfully.";
                                    //    LogText = "";
                                    //    //CreateAuditFile(xmlfile, "Navibulgar_HTTP_Processor", VRNO, "Downloaded", xmlfile + " downloaded successfully.", BuyerCode, SupplierCode, AuditPath);
                                    //    if ((this.VRNO + "|" + this.Vessel + "|" + this.Port).Length > 0)
                                    //        File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + this.Domain + "_" + this.SupplierCode + "_GUID.txt", this.VRNO + "|" + this.Vessel + "|" + this.Port + Environment.NewLine);
                                    //}
                                    //else
                                    //{
                                    //    LogText = "Unable to download file " + xmlfile;
                                    //    string filename = PrintScreenPath + "\\Navibulgar_RFQError_" + DateTime.Now.ToString("ddMMyyyHHmmssfff") + ".png";
                                    //    //CreateAuditFile(filename, "Navibulgar_HTTP_Processor", VRNO, "Error", "Unable to download file " + xmlfile + " for ref " + VRNO + ".", BuyerCode, SupplierCode, AuditPath);
                                    //    //CreateAuditFile(filename, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1004:Unable to process file " + xmlfile + " for ref " + VRNO, BuyerCode, SupplierCode, AuditPath);
                                    //    if (PrintScreen(filename)) filename = "";
                                    //}

                                }
                                else
                                {
                                    LogText = "Unable to get address details";
                                    AuditMessageData.CreateAuditFile(fileName, MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "LeS-1040:Unable to get details-address Field(s) not present" + UCRefNo);

                                    //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1040:Unable to get details-address Field(s) not present", BuyerCode, SupplierCode, AuditPath);
                                }

                            }
                            else
                            {
                                LogText = "Unable to get RFQ item details";
                                //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1040:Unable to get details-RFQ item Field(s) not present", BuyerCode, SupplierCode, AuditPath);
                                AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "LeS-1040:Unable to get details-address Field(s) not present" + UCRefNo);
                            }
                        }
                        else
                        {
                            LogText = "Unable to get RFQ header details";
                            //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1040:Unable to get details-RFQ header Field(s) not present", BuyerCode, SupplierCode, AuditPath);
                            AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "LeS-1040:Unable to get details-address Field(s) not present" + UCRefNo);

                        }
                    }
                }
                catch (Exception ex)
                {
                    //WriteErrorLog_With_Screenshot("Unable to load RFQ '" + VRNO + "' details due to " + ex.GetBaseException().Message.ToString());
                    WriteErrorLog_With_Screenshot("Unable to filter details for " + VRNO + " due to " + ex.GetBaseException().Message.ToString(), "LeS-1006:");
                }
            }
        }

        public bool GetRFQHeader(ref DocHeader docHeader, string eFile)
        {
            bool isResult = false;
            LogText = "Start Getting Header details";
            try
            {

                DateTimePeriodCollection rfqDateTimePeriod = new DateTimePeriodCollection();
                //_lesXml.DocID = DateTime.Now.ToString("yyyyMMddhhmmss");
                //_lesXml.Created_Date = DateTime.Now.ToString("yyyyMMdd");
                if (docHeader.References == null)
                    docHeader.References = new ReferenceCollection();

                if (docHeader.Comments == null)
                    docHeader.Comments = new CommentsCollection();

                if (docHeader.Equipment == null)
                    docHeader.Equipment = new Equipment();
                docHeader.DocType = "REQUESTFORQUOTE";
                //_lesXml.Dialect = "Navigation Maritime Bulgare";
                docHeader.VersionNumber = "1";
                //_lesXml.Date_Document = DateTime.Now.ToString("yyyyMMdd");
                //_lesXml.Date_Preparation = DateTime.Now.ToString("yyyyMMdd"); shifted to interchange date + time
                //_lesXml.Sender_Code = BuyerCode;
                //_lesXml.Recipient_Code = SupplierCode;

                //docHeader.LineItemCount = rfqData1?.data?.enquiryQuoteItems?.Count() ?? 0;
                HtmlNode _InqNo = _httpWrapper.GetElement("span", "id", "MainContent_IDControl");
                if (_InqNo != null)
                {
                    docHeader.MessageNumber = UCRefNo = _InqNo.InnerText.Trim();
                    docHeader.References.Add(new Reference(ReferenceQualifier.UC, _InqNo.InnerText.Trim()));
                    string cBranchDetails = Userid + "_" + Password;
                    docHeader.MessageReferenceNumber = _InqNo.InnerText.Trim() + "|" + cBranchDetails;//added by kalpita on 16/07/2021(wrist)
                }
                HtmlNode _Payterms = _httpWrapper._CurrentDocument.GetElementbyId("MainContent_Label1");//added by kalpita on 16/07/2021(wrist)

                //pending
                if (_Payterms != null)
                {
                    //_lesXml.Remark_PaymentTerms = _Payterms.InnerText.Replace("* Payment terms:", "").Trim();
                    var a = _Payterms.InnerText.Replace("* Payment terms:", "").Trim();
                }
                docHeader.MessageReferenceNumber = URL;

                //if (File.Exists(eFile))
                //    _interchange.MTMLFile = Path.GetFileName(eFile);


                //_lesXml.Active = "1";

                //HtmlNode _Vessel = _httpWrapper.GetElement("span", "id", "MainContent_IDControl0");
                //if (_Vessel != null)
                //    _lesXml.Vessel = _Vessel.InnerText.Trim();  //shift to party

                //_lesXml.BuyerRef = _InqNo.InnerText.Trim(); 

                //HtmlNode _portname = _httpWrapper.GetElement("span", "id", "MainContent_n_PortN");
                //if (_portname != null)
                //{
                //    _lesXml.PortName = _portname.InnerText.Trim();
                //    _lesXml.PortCode = _portname.InnerText.Trim();//added by kalpita on 12/07/2021  //shift to party
                //}

                //HtmlNode _etaDate = _httpWrapper.GetElement("span", "id", "MainContent_Vess_ETA");  //shift to party


                //_lesXml.Currency = "";
                docHeader.Equipment = new Equipment();
                if (URL.Contains("Edit_Spares.aspx"))
                {
                    HtmlNode _equipment = _httpWrapper.GetElement("span", "id", "MainContent_lblEq");
                    if (_equipment != null)
                    {
                        docHeader.Equipment.Name = _equipment.InnerText.Trim();
                    }
                    string b = "";
                    HtmlNode _equipSys = _httpWrapper.GetElement("span", "id", "MainContent_lblSys");
                    if (_equipSys != null) { docHeader.Equipment.Remarks = "System: " + _equipSys.InnerText.Trim(); }


                    HtmlNode _equipNo = _httpWrapper.GetElement("span", "id", "MainContent_lblSNom");
                    if (_equipNo != null) { docHeader.Equipment.Remarks += "Ser. No.: " + _equipNo.InnerText.Trim(); }

                }

                string PUR = "";
                HtmlNode _remarks = _httpWrapper.GetElement("textarea", "id", "MainContent_txtAddInf");
                if (_remarks != null)
                { PUR = _remarks.InnerText.Trim(); }
                docHeader.Comments.Add(new Comments(CommentTypes.PUR, PUR));

                // creationDate
                HtmlNode _etaDate = _httpWrapper.GetElement("span", "id", "MainContent_n_DateC");
                if (_etaDate != null)
                {
                    string strEtaDate = _etaDate.InnerText.Trim();
                    // creationDate
                    if (strEtaDate != "" && strEtaDate != "-")
                    {
                        DateTime date = DateTime.ParseExact(strEtaDate, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                        string _rfqdate = date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                        DateTime _dt = DateTime.MinValue;
                        DateTimePeriod _dtDocDate = new DateTimePeriod();
                        DateTime.TryParseExact(_rfqdate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _dt);
                        _dtDocDate.FormatQualifier = DateTimeFormatQualifiers.CCYYMMDD_102;
                        _dtDocDate.Qualifier = DateTimePeroidQualifiers.DocumentDate_137;
                        _dtDocDate.Value = _dt.ToString("yyyyMMdd");
                        rfqDateTimePeriod.Add(_dtDocDate);
                    }
                }
                docHeader.DateTimePeriods = rfqDateTimePeriod;

                //_lesXml.Total_LineItems_Discount = "0";
                //_lesXml.Total_LineItems_Net = "0";
                //_lesXml.Total_Additional_Discount = "0";
                //_lesXml.Total_Freight = "0";
                //_lesXml.Total_Other = "0";
                //_lesXml.Total_Net_Final = "0";
                LogText = "Getting Header details completed successfully.";
                isResult = true;
                return isResult;
            }
            catch (Exception ex)
            {
                LogText = "Unable to get header details." + ex.GetBaseException().ToString(); isResult = false;
                return isResult;
            }
        }

        public bool GetRFQItems(ref LineItemCollection lineItems)
        {
            bool isResult = false;
            string EquipRemarks = "";
            try
            {
                lineItems.Clear();
                LogText = "Start Getting LineItem details";
                if (URL.Contains("Edit_Spares.aspx"))
                {
                    HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@class='style2E']//tr");
                    if (_nodes != null)
                    {
                        if (_nodes.Count >= 2)
                        {
                            int i = 0;
                            foreach (HtmlNode _row in _nodes)
                            {

                                LineItem _item = new LineItem();
                                try
                                {
                                    HtmlNodeCollection _rowData = _row.ChildNodes;
                                    if (!_rowData[1].InnerText.Trim().Contains("Part No."))
                                    {
                                        // i += 1;
                                        if (_rowData[1].InnerText.Trim().Contains("Subsystem"))
                                        {
                                            EquipRemarks = _rowData[1].InnerText.Trim();
                                        }
                                        if (_rowData.Count == 21)
                                        {
                                            //i += 1;
                                            //_item.Number = Convert.ToString(i);
                                            //_item.OriginatingSystemRef = Convert.ToString(i);
                                            //string _inpQty = _rowData[11].SelectSingleNode("./input").GetAttributeValue("Id", "").Trim();
                                            //_item.OriginatingSystemRef = _inpQty;
                                            ////   _item.OriginatingSystemRef = _rowData[1].InnerText.Trim();
                                            //// _item.Description = _rowData[5].InnerText.Trim();
                                            //_item.Name = _rowData[5].InnerText.Trim();
                                            //_item.Unit = _rowData[9].InnerText.Trim();
                                            //_item.Quantity = _rowData[7].SelectSingleNode("./input").GetAttributeValue("value", "").Trim();
                                            //_item.Discount = "0";
                                            //_item.ListPrice = "0";
                                            //_item.LeadDays = "0";
                                            ////  _item.Remark = "Draw No.: " + _rowData[3].InnerText.Trim();
                                            //_item.SystemRef = Convert.ToString(i);
                                            //_item.ItemRef = _rowData[1].InnerText.Trim().Replace("TAG NO:", "");
                                            //if (EquipRemarks != "")
                                            //{ _item.EquipRemarks = EquipRemarks; EquipRemarks = ""; }
                                            //_lesXml.LineItems.Add(_item);
                                            //if (_rowData[3].InnerText.Trim() != "")
                                            //{
                                            //    _item.Remark = "Draw No: " + _rowData[3].InnerText.Trim();
                                            //}
                                            i += 1;
                                            _item.Number = Convert.ToString(i);
                                            string _inpQty = _rowData[5].SelectSingleNode("./input").GetAttributeValue("Id", "").Trim();
                                            _item.OriginatingSystemRef = Convert.ToString(i); ;
                                            _item.Description = _rowData[2].InnerText.Trim();
                                            _item.MeasureUnitQualifier = _rowData[3].InnerText.Trim();
                                            _item.Quantity = Convert.ToDouble(_rowData[4].InnerText.Trim());
                                            _item.Description = _rowData[2].InnerText.Trim();
                                            _item.Discount_Value = 0.0;
                                            _item.LineItemComment = new Comments();
                                            _item.LineItemComment.Qualifier = CommentTypes.LIN;
                                            _item.LineItemComment.Value = "sample comment";

                                            //_item.ListPrice = "0";
                                            //_item.LeadDays = "0";
                                            //_item.SystemRef = Convert.ToString(i);
                                            var a = _rowData[1].InnerText.Trim().Replace("&nbsp;", "");
                                            lineItems.Add(_item);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                { LogText = ex.GetBaseException().ToString(); }
                            }
                            //_lesXml.Total_LineItems = Convert.ToString(_lesXml.LineItems.Count);
                            isResult = true;
                        }
                        else isResult = false;
                    }
                    else isResult = false;
                }
                else if (URL.Contains("Edit_Req.aspx"))
                {
                    HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@id='MainContent_GridView1']//tr");
                    if (_nodes != null)
                    {
                        if (_nodes.Count >= 2)
                        {
                            int i = 0;
                            foreach (HtmlNode _row in _nodes)
                            {
                                LineItem _item = new LineItem();
                                try
                                {
                                    HtmlNodeCollection _rowData = _row.ChildNodes;
                                    if (!_rowData[1].InnerText.Trim().Contains("Item Code"))
                                    {
                                        i += 1;
                                        _item.Number = Convert.ToString(i);
                                        string _inpQty = _rowData[5].SelectSingleNode("./input").GetAttributeValue("Id", "").Trim();
                                        _item.OriginatingSystemRef = Convert.ToString(i);
                                        _item.Description = _rowData[2].InnerText.Trim();
                                        _item.MeasureUnitQualifier = _rowData[3].InnerText.Trim();
                                        _item.Quantity = Convert.ToDouble(_rowData[4].InnerText.Trim());
                                        _item.Description = _rowData[2].InnerText.Trim();
                                        _item.Discount_Value = 0.0;
                                        //_item.ListPrice = "0";
                                        //_item.LeadDays = "0";
                                        //_item.SystemRef = Convert.ToString(i);
                                        var a = _rowData[1].InnerText.Trim().Replace("&nbsp;", "");
                                        _item.LineItemComment = new Comments();
                                        _item.LineItemComment.Qualifier = CommentTypes.LIN;
                                        _item.LineItemComment.Value = "sample comment";
                                        lineItems.Add(_item);
                                    }
                                }
                                catch (Exception ex)
                                { LogText = ex.GetBaseException().ToString(); }
                            }
                            //_lesXml.Total_LineItems = Convert.ToString(_lesXml.LineItems.Count);
                            isResult = true;
                        }
                        else isResult = false;
                    }
                    else isResult = false;
                }
                else
                {
                    //WriteErrorLog_With_Screenshot("Unable to load URL", "LeS-1016:");//Invalid url

                }
                LogText = "Getting LineItem details successfully";
                return isResult;
            }
            catch (Exception ex)
            {
                LogText = "Exception while getting RFQ Items: " + ex.GetBaseException().ToString(); isResult = false; return isResult;
            }
        }

        public bool GetAddress(ref PartyCollection collection)
        {
            bool isResult = false;
            try
            {
                string vesselName = "";
                string deliveryPlace = "";
                string deliveryCode = "";
                //PartyCollection collection = new PartyCollection();

                HtmlNode _Vessel = _httpWrapper.GetElement("span", "id", "MainContent_IDControl0");
                if (_Vessel != null)
                    vesselName = _Vessel.InnerText.Trim();  //shift to party


                HtmlNode _portname = _httpWrapper.GetElement("span", "id", "MainContent_n_PortN");
                if (_portname != null)
                {
                    deliveryPlace = _portname.InnerText.Trim();
                    deliveryCode = _portname.InnerText.Trim();//added by kalpita on 12/07/2021  //shift to party
                }

                //HtmlNode _etaDate = _httpWrapper.GetElement("span", "id", "MainContent_Vess_ETA");  //shift to party
                Party vessel = new Party();
                vessel.Qualifier = PartyQualifier.UD;
                vessel.Name = vesselName;
                vessel.PartyLocation = new PartyLocation();
                vessel.PartyLocation.Berth = deliveryPlace;
                collection.Add(vessel);

                Party buyer = new Party();
                buyer.Qualifier = PartyQualifier.BY;
                Contact contact = new Contact();
                contact.Name = BuyerName;
                buyer.Contacts.Add(contact);
                collection.Add(buyer);

                Party vendor = new Party();
                vendor.Qualifier = PartyQualifier.VN;
                vendor.Name = SupplierName;
                collection.Add(vendor);

                LogText = "Getting address details successfully";
                isResult = true;
                return isResult;
            }
            catch (Exception ex)
            {
                LogText = "Exception while getting address details: " + ex.GetBaseException().ToString(); isResult = false;
                return isResult;
            }


        }

        #endregion


        #region ## Quote ##
        public string currentMoveDirectory = string.Empty;
        string currentMtmlFile = string.Empty;
        int currentAppSettingCount = 0;
        public void ProcessQuote()
        {
            int j = 0;
            try
            {
                var quoteFiles = _fileStore.DownloadFilesBulk(LeS.FileStore.DownloadFileType.Quote_Poc, FileMoveAction.MoveToInProcess, dctAppSettings, Convert.ToInt32(LINKID));
                currentMoveDirectory = quoteFiles.MovedDirPath; //5 
                string[] outboxfiles = quoteFiles.SavedFiles.ToArray();
                if (outboxfiles.Length > 0)
                {
                    for (j = 0; j < outboxfiles.Length; j++)
                    {
                        string file = outboxfiles[j];
                        currentMtmlFile = file;
                        if (currentMtmlFile.Contains("QUOTE")) { ProcessQuoteMTML(file); }
                    }
                    currentAppSettingCount++;
                }
                else
                {
                    LogText = "No quote document found in directory.";
                }
                //CheckQuoteFiles();//added by kalpita on 09/01/2020

                //LogText = "";
                //LogText = "Quote processing started."; slPendingFiles.Clear();
                //GetXmlFiles();
                //if (xmlFiles.Count > 0)
                //{
                //    LogText = xmlFiles.Count + " Quote files found to process.";
                //    for (j = 0; j < xmlFiles.Count; j++)
                //    {
                //        dctSubmitValues.Clear();
                //        ProcessQuoteMTML(xmlFiles[j]);
                //    }
                //}
                //else LogText = "No quote file found.";
                LogText = "Quote processing stopped.";

            }
            catch (Exception e)
            {
                LogText = e.StackTrace;
                //WriteErrorLogQuote_With_Screenshot("Exception in Process Quote : " + e.GetBaseException().ToString(), xmlFiles[j]);
                WriteErrorLogQuote_With_Screenshot("LeS-1004:Unable to process file due to " + e.GetBaseException().ToString(), xmlFiles[j]);
            }
        }

        public void GetXmlFiles()
        {
            int i = 0;
            xmlFiles.Clear();
            DirectoryInfo _dir = new DirectoryInfo(MTML_QuotePath);
            FileInfo[] _Files = _dir.GetFiles();
            if (_Files != null)
            {
                foreach (FileInfo _MtmlFile in _Files)
                {
                    string cFilename = _MtmlFile.Name;
                    if (cFilename.ToUpper().Contains("QUOTE"))//added by kalpita on 21/07/2020
                    {
                        xmlFiles.Add(_MtmlFile.FullName);
                        if (i == 0) { xmlFiles.Add(_MtmlFile.FullName); }
                        else { File.Move(_MtmlFile.FullName, ToCheckPath + "\\" + _MtmlFile.Name); }
                    }
                    else if (cFilename.ToUpper().Contains("POC")) { File.Move(_MtmlFile.FullName, cPOCPath + "\\" + _MtmlFile.Name); }
                    else { File.Move(_MtmlFile.FullName, ToCheckPath + "\\" + _MtmlFile.Name); }
                    i++;
                }
            }
        }

        public void CheckQuoteFiles()
        {
            DirectoryInfo _dir = new DirectoryInfo(ToCheckPath);
            FileInfo[] _Files = _dir.GetFiles();
            if (_Files != null)
            {
                nFileCount = _Files.Count();
                while (nFileCount > 0)
                {
                    int j = 0;
                    try
                    {
                        foreach (FileInfo _MtmlFile in _Files)
                        {
                            if (j == 0) { File.Move(ToCheckPath + "\\" + _MtmlFile.Name, MTML_QuotePath + "\\" + _MtmlFile.Name); j++; }
                            else { break; }
                        }
                        ProcessQuote();
                        nFileCount--;
                    }
                    catch (Exception e)
                    { }
                }
            }
        }

        public void ProcessQuoteMTML(string MTML_QuoteFile)
        {
            string eFile = "";
            try
            {
                MTMLClass _mtml = new MTMLClass();
                _interchange = _mtml.Load(MTML_QuoteFile);
                LoadInterchangeDetails();
                string[] ArrCredentials = MsgRefNumber.Split('|');
                if (ArrCredentials != null && ArrCredentials.Length > 1)//added by kalpita on 16/07/2021
                {
                    Userid = ArrCredentials[1].Split('_')[0]; Password = ArrCredentials[1].Split('_')[1];
                }
                if (!IsDownloadFromPortal)
                {
                    bool IsLoggedIn = DoLogin("input", "id", "MainContent_txtFrom");
                }
                if (UCRefNo != "")
                {
                    LogText = "Quote processing started for refno: " + UCRefNo;
                    URL = MessageNumber;
                    if (LoadURL("input", "id", "MainContent_btnUpdate"))
                    {
                        HtmlNode _hRefNo = _httpWrapper.GetElement("span", "id", "MainContent_IDControl");
                        if (_hRefNo != null)
                        {
                            if (convert.ToString(_hRefNo.InnerText).Trim() == UCRefNo)
                            {
                                LogText = "Reference Number Matched";
                                HtmlNode _btnSave = _httpWrapper.GetElement("input", "id", "MainContent_btnSave");
                                if (_btnSave != null)
                                {
                                    LogText = "Checking Quote Status";
                                    HtmlNode _lblSubmit = _httpWrapper.GetElement("span", "id", "MainContent_lblSubmit");
                                    if (_lblSubmit == null) // (This Quotation is Submitted!) Label on Edit Page
                                    {
                                        LogText = "Quote is in progress state";

                                        #region  Commented
                                        //if (!_lblSubmit.InnerText.Trim().Contains("This Quotation is Submitted!"))
                                        //{
                                        //}
                                        //else
                                        //{
                                        //    eFile = "Navigation_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                        //    PrintScreen(ImgPath + "\\" + eFile);
                                        //    MoveToError("Quotation is already submitted for refno " + UCRefNo, UCRefNo, MTML_QuoteFile, eFile);
                                        //}
                                        #endregion

                                        double total = 0; string cSaveReqLink = "";
                                        int result = FillQuotation(ref total, out cSaveReqLink);

                                        if (result == 1)
                                        {
                                            eFile = "Navigation_Save_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                            string screenshotpath = Path.Combine(ImgPath,eFile);
                                            PrintScreen(screenshotpath);

                                            string QuoteRef = "";
                                            foreach (Reference Ref in _interchange.DocumentHeader.References)
                                            {
                                                if (Ref.Qualifier == ReferenceQualifier.AAG)
                                                {
                                                    QuoteRef = convert.ToString(Ref.ReferenceNumber).Trim();
                                                    break;
                                                }
                                            }
                                            if (SubmitQuotation(QuoteRef, cSaveReqLink, MessageNumber)) // Updated By Sanjita
                                            {
                                                eFile = "Navigation_Submit_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                                screenshotpath = Path.Combine(ImgPath, eFile);
                                                PrintScreen(screenshotpath);
                                                _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToBackup, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                                                //MoveToBakup(MTML_QuoteFile, "Quote '" + UCRefNo + "' submitted successfully.", eFile);

                                                // Send Mail (Added By Sanjita 01-DEC-18) //
                                                if (SendMail)//added by kalpita on 04/09/2020
                                                {
                                                    //SendMailNotification(_interchange, "QUOTE", UCRefNo, "SUBMITTED", "Quote '" + UCRefNo + "' Submitted Successfully.");
                                                    SendMailNotification_JSON(UCRefNo, SendMailQueueAction.SAVED, DocType.QUOTE, AuditMessageData.SUBMIT_SUCCESS, "", "", BUYER_LINK_CODE, SUPPLIER_LINK_CODE, LINKID, NotificationPath, 0, MODULE_NAME + "_QUOTE");

                                                }
                                            }
                                            else
                                            {
                                                // Added By Sanjita on 29-NOV-18 //
                                                eFile = "Navigation_Error_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                                screenshotpath = Path.Combine(ImgPath, eFile);
                                                PrintScreen(screenshotpath);
                                                _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);
                                                //SendMailNotification_JSON(UCRefNo, SendMailQueueAction.SAVED, DocType.QUOTE, AuditMessageData.UNABLE_TO_SUBMIT, "", "", BUYER_LINK_CODE, SUPPLIER_LINK_CODE, LINKID, NotificationPath, 0, MODULE_NAME + "_QUOTE");

                                                //MoveToError("Unable to submit quote for '" + UCRefNo + "'", UCRefNo, MTML_QuoteFile, eFile);
                                                //MoveToError("LeS-1011:Unable to submit Quote for '" + UCRefNo + "'", UCRefNo, MTML_QuoteFile, eFile);
                                            }
                                        }
                                        else if (result == 0)
                                        {
                                            eFile = "Navigation_SaveDiff_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                            string screenshotpath = Path.Combine(ImgPath, eFile);
                                            PrintScreen(screenshotpath);
                                            _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                                            //MoveToError("Total mismatched as site total " + total + " and mtml total " + GrandTotal, UCRefNo, MTML_QuoteFile, eFile);
                                            //MoveToError("LeS-1008.1:Unable to save quote due to amount mismatched", UCRefNo, MTML_QuoteFile, eFile);
                                        }
                                        else if (result == 2)
                                        {
                                            eFile = "Navigation_SaveError_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                            string screenshotpath = Path.Combine(ImgPath, eFile);
                                            PrintScreen(screenshotpath);
                                            _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                                            //MoveToError("Unable to save quotation." + GrandTotal, UCRefNo, MTML_QuoteFile, eFile);
                                            // MoveToError("LeS-1008:Unable to save quotation", UCRefNo, MTML_QuoteFile, eFile);
                                        }
                                    }
                                    else
                                    {
                                        if (convert.ToString(_lblSubmit.InnerText).Trim().Contains("This Quotation is Submitted!"))
                                        {
                                            eFile = "Navigation_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                            string screenshotpath = Path.Combine(ImgPath, eFile);
                                            PrintScreen(screenshotpath);
                                            _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToBackup, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                                            //MoveToError("Quotation is already submitted for Ref No '" + UCRefNo + "'.", UCRefNo, MTML_QuoteFile, eFile);
                                            //MoveToError("LeS-1011.1:Unable to submit Quote since Quotation is already submitted for " + UCRefNo + ".", UCRefNo, MTML_QuoteFile, eFile);
                                        }
                                        else
                                        {
                                            eFile = "Navigation_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                            string screenshotpath = Path.Combine(ImgPath, eFile);
                                            PrintScreen(screenshotpath);
                                            _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                                            //MoveToError("Submit button is disabled for quote '" + UCRefNo + "'", UCRefNo, MTML_QuoteFile, eFile);
                                            //MoveToError("LeS-1011.1:Unable to Submit Quote since submit button is not active for " + UCRefNo, UCRefNo, MTML_QuoteFile, eFile);
                                        }
                                    }
                                }
                                else
                                {
                                    eFile = "Navigation_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                    string screenshotpath = Path.Combine(ImgPath, eFile);
                                    PrintScreen(screenshotpath);
                                    _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                                    //MoveToError("Unable to get save button for refno " + UCRefNo, UCRefNo, MTML_QuoteFile, eFile);
                                    //MoveToError("LeS-1008.3:Unable to Save Quote due to missing controls for " + UCRefNo, UCRefNo, MTML_QuoteFile, eFile);
                                }
                            }

                            else
                            {
                                eFile = "Navigation_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                                PrintScreen(ImgPath + "\\" + eFile);
                                _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);
                                //MoveToError("Ref No. mismatch betwwen site ref no " + _hRefNo.InnerText.Trim() + " and quote file ref no " + UCRefNo + ".", UCRefNo, MTML_QuoteFile, eFile);
                                //MoveToError("LeS-1003.1:Issue after loading URL as Ref No. mismatched with site Ref No. " + _hRefNo.InnerText.Trim(), UCRefNo, MTML_QuoteFile, eFile);
                            }
                        }
                    }
                    else
                    {
                        eFile = "Navigation_" + SupplierCode + "_" + BuyerCode + "_" + convert.ToFileName(UCRefNo) + "_" + DateTime.Now.ToString("ddMMyyHHmmssff") + ".png";
                        if (PrintScreen(ImgPath + "\\" + eFile)) eFile = "";
                        _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);

                        //MoveToError("Unable to load details page.", UCRefNo, MTML_QuoteFile, eFile);
                        // MoveToError("LeS-1006:Unable to filter details.", UCRefNo, MTML_QuoteFile, eFile);
                    }
                }
            }
            catch (Exception ex)
            {
                LogText = "Error - " + ex.Message;
                LogText = ex.StackTrace;
                //WriteErrorLogQuote_With_Screenshot("Exception while processing Quote '" + UCRefNo + "', Error - " + ex.Message, Path.GetFileName(MTML_QuoteFile));
                WriteErrorLogQuote_With_Screenshot("LeS-1004:Unable to process file for" + UCRefNo + " due to " + ex.Message, Path.GetFileName(MTML_QuoteFile));
            }
            finally
            {
                dctPostDataValues.Clear(); dctSubmitValues.Clear();//added by kalpita on 24/10/2019
            }
        }

        public int FillQuotation(ref double total, out string SaveReqLink)
        {
            SaveReqLink = "";
            //   bool result = false;
            int result = 0;
            try
            {
                // _httpWrapper.IsUrlEncoded = false;
                FillHeaderDetails();
                FillItemDetails();

                dctPostDataValues.Add("ctl00%24MainContent%24btnSave", "Save+%26+Calculate");
                dctPostDataValues.Add("ctl00%24MainContent%24txtHelp", Convert.ToDouble(GrandTotal).ToString());


                string[] ArrURL = URL.Split('=');
                URL = ArrURL[0] + "=" + HttpUtility.UrlEncode(ArrURL[1]);
                SaveReqLink = URL;
                if (PostURL("span", "id", "MainContent_lblGrandTot"))
                {
                    if (convert.ToDouble(GrandTotal) == 0 || IsDecline)
                    {
                        result = 3;
                    }
                    else
                    {
                        HtmlNode _total = _httpWrapper.GetElement("span", "id", "MainContent_lblGrandTot");
                        if (_total != null)
                        {
                            total = convert.ToFloat(_total.InnerText.Trim());
                            //double diff = convert.ToFloat(GrandTotal) - convert.ToFloat(_total.InnerText.Trim());
                            //if (diff >= -2 && diff <= 2)
                            //    result = 1;
                            //else
                            //{
                            //    result = 0;
                            //}


                            // if (Convert.ToInt32(Convert.ToDouble(total)) == Convert.ToInt32(Convert.ToDouble(GrandTotal))) //added by kalpita on 09/09/2019 to check buyer total
                            if (Math.Truncate(Convert.ToDouble(total)) == Math.Truncate(Convert.ToDouble(GrandTotal)))//changed by kalpita 10/05/22
                            {
                                result = 1;
                            }
                            else if (Convert.ToInt32(Convert.ToDouble(total)) < Convert.ToInt32(Convert.ToDouble(GrandTotal)))//changed by kalpita on 11/10/21,10/05/22
                            {
                                double diff = Math.Abs(convert.ToFloat(GrandTotal) - convert.ToFloat(_total.InnerText.Trim()));
                                result = (diff <= 5) ? 1 : 0;
                                //if (diff >= -2 && diff <= 2)
                                //    result = 1;
                                //else
                                //{
                                //    result = 0;
                                //}
                            }
                            else if (BuyerTotal != "" && IsSubmitQuote_Buyer)//added by kalpita on 28/08/2020
                            {
                                if (Convert.ToInt32(Convert.ToDouble(total)) == Convert.ToInt32(Convert.ToDouble(BuyerTotal)))
                                {
                                    result = 1;
                                }
                                else
                                {
                                    double diff = Math.Abs(convert.ToFloat(BuyerTotal) - convert.ToFloat(_total.InnerText.Trim()));
                                    result = (diff <= 5) ? 1 : 0;
                                    //if (Convert.ToDouble(total) > Convert.ToDouble(BuyerTotal))
                                    //    _diff = Convert.ToInt32(Convert.ToDouble(total)) - Convert.ToInt32(Convert.ToDouble(BuyerTotal));
                                    //else if (Convert.ToDouble(BuyerTotal) > Convert.ToDouble(total))
                                    //    _diff = Convert.ToInt32(Convert.ToDouble(BuyerTotal)) - Convert.ToInt32(Convert.ToDouble(total));
                                    //if (_diff <= 1)
                                    //    result = 1;
                                    //else
                                    //    result = 0;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogText = "Exception while filling quote: " + ex.GetBaseException().Message.ToString();
                result = 2;
            }
            return result;
        }

        public void FillHeaderDetails()
        {
            LogText = "Started filling Header Details";
            dctPostDataValues.Clear();

            dctPostDataValues.Add("__EVENTTARGET", _httpWrapper._dctStateData["__EVENTTARGET"]);
            dctPostDataValues.Add("__EVENTARGUMENT", _httpWrapper._dctStateData["__EVENTARGUMENT"]);
            dctPostDataValues.Add("__VIEWSTATE", _httpWrapper._dctStateData["__VIEWSTATE"]);
            dctPostDataValues.Add("__VIEWSTATEGENERATOR", _httpWrapper._dctStateData["__VIEWSTATEGENERATOR"]);
            dctPostDataValues.Add("__EVENTVALIDATION", _httpWrapper._dctStateData["__EVENTVALIDATION"]);
            if (_httpWrapper._dctStateData.ContainsKey("__LASTFOCUS"))
                dctPostDataValues.Add("__LASTFOCUS", _httpWrapper._dctStateData["__LASTFOCUS"]);
            if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATEENCRYPTED"))
                dctPostDataValues.Add("__VIEWSTATEENCRYPTED", _httpWrapper._dctStateData["__VIEWSTATEENCRYPTED"]);
            //supp ref no
            dctPostDataValues.Add("ctl00%24MainContent%24txtTotRef", Uri.EscapeDataString(AAGRefNo));//

            //currency
            string _value = "";
            HtmlNodeCollection _options = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//select[@id='MainContent_cmbCurr']//option");
            if (_options != null)
            {
                if (_options.Count > 0)
                {
                    foreach (HtmlNode _opt in _options)
                    {
                        if (_opt.NextSibling.InnerText.Trim().Split('|')[0].ToUpper().Trim() == Currency.ToUpper())
                        {
                            _value = _opt.GetAttributeValue("value", "");
                            break;
                        }
                    }
                }
            }
            if (_value != "")
            {
                dctPostDataValues.Add("ctl00%24MainContent%24cmbCurr", _value);
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24cmbCurr"))
                    dctSubmitValues.Add("ctl00%24MainContent%24cmbCurr", _value);
                else
                    dctSubmitValues["ctl00%24MainContent%24cmbCurr"] = _value;
            }
            else
                throw new Exception("Unable to set currency for refno: " + UCRefNo);//

            //delivery charges
            double fCharges = 0.0f;//added by kalpita on 28/08/2020
            fCharges += (FreightCharge != "") ? convert.ToFloat(FreightCharge) : 0;
            fCharges += (PackingCost != "") ? convert.ToFloat(PackingCost) : 0;
            fCharges += (TaxCost != "") ? convert.ToFloat(TaxCost) : 0;//added by kalpita on 23/09/2020
            fCharges += GetDeliveryCharge_FromLineItem();
            //if (FreightCharge != "")
            if (fCharges > 0)
            {
                //dctPostDataValues.Add("ctl00%24MainContent%24txtDeliCharg", convert.ToFloat(FreightCharge).ToString("0.00"));
                //if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtDeliCharg"))
                //    dctSubmitValues.Add("ctl00%24MainContent%24txtDeliCharg", convert.ToFloat(FreightCharge).ToString("0.00"));
                dctPostDataValues.Add("ctl00%24MainContent%24txtDeliCharg", fCharges.ToString("0.00"));
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtDeliCharg"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtDeliCharg", fCharges.ToString("0.00"));
            }
            else
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtDeliCharg", "0.00");//
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtDeliCharg"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtDeliCharg", "0.00");

            }

            //All Discount
            if (Allowance != "")
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtAllDisc", convert.ToFloat(Allowance).ToString("0.00"));
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtAllDisc"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtAllDisc", convert.ToFloat(Allowance).ToString("0.00"));
            }
            else
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtAllDisc", "0.00");//
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtAllDisc"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtAllDisc", "0.00");
            }

            //delivery days
            if (LeadDays != "")
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtDaysDeli", LeadDays);
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtDaysDeli"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtDaysDeli", LeadDays);
                else
                    dctSubmitValues["ctl00%24MainContent%24txtDaysDeli"] = LeadDays;
            }
            else
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtDaysDeli", "0.00");//
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtDaysDeli"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtDaysDeli", "0.00");
            }

            //additional info
            HtmlNode _hAddInfo = _httpWrapper.GetElement("textarea", "id", "MainContent_txtAddInf");
            if (_hAddInfo != null)
            {
                string _AddInfo = _hAddInfo.InnerText.Trim();
                dctPostDataValues.Add("ctl00%24MainContent%24txtAddInf", _AddInfo);
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtAddInf"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtAddInf", _AddInfo);
                else
                    dctSubmitValues["ctl00%24MainContent%24txtAddInf"] = _AddInfo;
            }
            else
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtAddInf", "");//
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtAddInf"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtAddInf", "");
            }

            //general info
            string gInfo = SupplierComment.Replace("\r\n", " ");
            gInfo = gInfo.Replace("\n", " ");
            gInfo = gInfo.Replace("&nbsp;", "");

            if (gInfo != "" && dtExpDate != "")
                gInfo = gInfo + ". Expiry Date: " + dtExpDate;
            else if (gInfo == "" && dtExpDate != "")
                gInfo = "Expiry Date: " + dtExpDate;


            if (gInfo != "")
            {
                if (gInfo.Length > 200) gInfo = gInfo.Substring(0, 200);
                {
                    dctPostDataValues.Add("ctl00%24MainContent%24txtGComm", Uri.EscapeDataString(gInfo).Replace("%20", "+"));
                    if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtGComm"))
                        dctSubmitValues.Add("ctl00%24MainContent%24txtGComm", Uri.EscapeDataString(gInfo).Replace("%20", "+"));
                }
            }
            else
            {
                dctPostDataValues.Add("ctl00%24MainContent%24txtGComm", "");//
                if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24txtGComm"))
                    dctSubmitValues.Add("ctl00%24MainContent%24txtGComm", "");
            }

            if (URL.Contains("Edit_Spares.aspx"))
            {
                //hidden element
                dctPostDataValues.Add("ctl00%24MainContent%24hiddenElement", _httpWrapper._dctStateData["MainContent_hiddenElement"]); //
            }

            LogText = "Filling Header Details completed";
        }

        //added by kalpita on 28/08/2020
        private Double GetDeliveryCharge_FromLineItem()
        {
            double fCharge = 0.0f;
            foreach (LES.MTML.LineItem item in _lineitem)
            {
                if (item.ItemType == "Charge")
                {
                    fCharge = convert.ToFloat(item.MonetaryAmount);
                    break;
                }
            }
            return fCharge;
        }

        public void FillItemDetails()
        {
            LogText = "Started filling Item Details";
            if (URL.Contains("Edit_Spares.aspx"))
            {
                HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@class='style2E']//tr");
                if (_nodes != null)
                {
                    if (_nodes.Count >= 2)
                    {
                        int i = 1;
                        LES.MTML.LineItem _item = null;
                        foreach (HtmlNode _tr in _nodes)
                        {
                            string cSiteUnit = "", cSiteQty = "";
                            HtmlNodeCollection _td = _tr.ChildNodes;
                            if (!_td[1].InnerText.Trim().Contains("Part No.") && !_td[1].InnerText.Trim().Contains("Subsystem") && !_td[1].InnerText.Trim().Contains("Comment:") && _td.Count == 21)
                            {
                                foreach (LES.MTML.LineItem item in _lineitem)
                                {

                                    if (_td[11].ChildNodes[1].GetAttributeValue("id", "").Trim() == item.OriginatingSystemRef)
                                    {
                                        _item = item; cSiteUnit = _td[3].InnerText.Trim(); cSiteQty = _td[4].InnerText.Trim();
                                        break;
                                    }
                                }

                                if (_item != null)
                                {
                                    string _price = "", _discount = "", cPrice_Comments = "";
                                    foreach (PriceDetails _priceDetails in _item.PriceList)
                                    {
                                        if (_priceDetails.TypeQualifier == PriceDetailsTypeQualifiers.GRP) _price = _priceDetails.Value.ToString("0.00");
                                        else if (_priceDetails.TypeQualifier == PriceDetailsTypeQualifiers.DPR) _discount = _priceDetails.Value.ToString("0.00");
                                    }
                                    //added by kalpita on 28/08/2020
                                    if (cSiteQty != _item.Quantity.ToString("0.00")) { cPrice_Comments += "Qty Changed : (" + cSiteQty + "-" + _item.Quantity + "),"; }
                                    if (cSiteUnit != _item.MeasureUnitQualifier) { cPrice_Comments += " Unit Changed : (" + cSiteUnit + "-" + _item.MeasureUnitQualifier + "),"; }
                                    cPrice_Comments += (_price == "0.00") ? "Line Item " + _item.Number + " is Cancelled. " : " Unit Price -" + _price + ",";

                                    if (i <= 9)
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24App_Q", _item.Quantity.ToString("0.00"));
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24Loc_Unit", _price);
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24RDisc", _discount);
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24Sum", _item.MonetaryAmount.ToString("0.00"));
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24DropDownList2", "OEM"); // Updated By Sanjita for OEM
                                        if (_item.DeliveryTime != null)
                                            dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24RDaysD", _item.DeliveryTime);
                                        else
                                            dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24RDaysD", "0");

                                        // Submit dic //                                        
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24App_Q", _item.Quantity.ToString("0.00"));
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24Loc_Unit", _price);
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24RDisc", _discount);
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24Sum", _item.MonetaryAmount.ToString("0.00"));
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24DropDownList2", "OEM"); // Updated By Sanjita for OEM
                                        if (_item.DeliveryTime != null)
                                            dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24RDaysD", _item.DeliveryTime);
                                        else
                                            dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24RDaysD", "0");
                                    }
                                    else
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24App_Q", _item.Quantity.ToString("0.00"));
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24Loc_Unit", _price);
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24RDisc", _discount);
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24Sum", _item.MonetaryAmount.ToString("0.00"));
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24DropDownList2", "OEM"); // Updated By Sanjita for OEM
                                        if (_item.DeliveryTime != null)
                                            dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24RDaysD", _item.DeliveryTime);
                                        else
                                            dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24RDaysD", "0");

                                        // Submit Dic //
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24App_Q", _item.Quantity.ToString("0.00"));
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24Loc_Unit", _price);
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24RDisc", _discount);
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24Sum", _item.MonetaryAmount.ToString("0.00"));
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24DropDownList2", "OEM"); // Updated By Sanjita for OEM
                                        if (_item.DeliveryTime != null)
                                            dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24RDaysD", _item.DeliveryTime);
                                        else
                                            dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24RDaysD", "0");
                                    }
                                    string _comments = "";
                                    if (_item.LineItemComment.Value.Trim() != "")
                                    {
                                        _comments = _item.LineItemComment.Value.Trim().Replace("\r\n", " ");
                                        _comments = _item.LineItemComment.Value.Trim().Replace("\n", " ");
                                        _comments = _item.LineItemComment.Value.Trim().Replace("&nbsp;", " ");
                                        _comments += cPrice_Comments.TrimEnd(',');//added by kalpita on 28/08/2020
                                    }
                                    else
                                    {
                                        _comments += cPrice_Comments.TrimEnd(',');//added by kalpita on 28/08/2020
                                    }

                                    if (_comments.Length > 240) _comments = _comments.Substring(0, 240);//added by kalpita on 26/11/2024
                                    if (i <= 9)
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24txtComm", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl0" + i + "%24txtComm", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                    }
                                    else
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24txtComm", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                        dctSubmitValues.Add("ctl00%24MainContent%24Repeater1%24ctl" + i + "%24txtComm", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                    }
                                    i++;
                                }
                                else throw new Exception("Item " + _td[1].InnerText.Trim() + " not found on website for ref no" + UCRefNo);
                            }
                        }
                        i = 1;
                    }
                }
            }
            else if (URL.Contains("Edit_Req.aspx"))
            {
                HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@id='MainContent_GridView1']//tr");
                if (_nodes != null)
                {
                    if (_nodes.Count >= 2)
                    {
                        int i = 2;
                        LES.MTML.LineItem _item = null;
                        foreach (HtmlNode _tr in _nodes)
                        {
                            string cSiteUnit = "", cSiteQty = "";
                            HtmlNodeCollection _td = _tr.ChildNodes;
                            if (!_td[1].InnerText.Trim().Contains("Item Code"))
                            {
                                foreach (LES.MTML.LineItem item in _lineitem)
                                {
                                    if (_td[5].ChildNodes[1].GetAttributeValue("id", "").Trim() == item.OriginatingSystemRef)
                                    {
                                        // if (_td[1].InnerText.Trim() == item.Identification)
                                        // {
                                        // }
                                        cSiteUnit = _td[3].InnerText.Trim(); cSiteQty = _td[4].InnerText.Trim();
                                        _item = item;
                                        break;
                                    }
                                }
                                if (_item != null)
                                {
                                    string _price = "", _discount = "", cPrice_Comments = "";
                                    foreach (PriceDetails _priceDetails in _item.PriceList)
                                    {
                                        if (_priceDetails.TypeQualifier == PriceDetailsTypeQualifiers.GRP) _price = _priceDetails.Value.ToString("0.00");
                                        else if (_priceDetails.TypeQualifier == PriceDetailsTypeQualifiers.DPR) _discount = _priceDetails.Value.ToString("0.00");
                                    }
                                    //added by kalpita on 28/08/2020
                                    if (cSiteQty != _item.Quantity.ToString("0.00")) { cPrice_Comments += "Qty Changed : (" + cSiteQty + "-" + _item.Quantity + "),"; }
                                    if (cSiteUnit != _item.MeasureUnitQualifier) { cPrice_Comments += " Unit Changed : (" + cSiteUnit + "-" + _item.MeasureUnitQualifier + "),"; }
                                    cPrice_Comments += (_price == "0.00") ? "Line Item " + _item.Number + " is cancelled. " : " Unit Price -" + _price + ",";
                                    if (i <= 9)
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox1", _price);
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox2", _discount);
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox3", LeadDays);

                                        //submit dic
                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox1"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox1", _price);
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox1"] = _price;

                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox2"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox2", _discount);
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox2"] = _discount;

                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox3"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox3", LeadDays);
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox3"] = LeadDays;
                                    }
                                    else
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox1", _price);
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox2", _discount);
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox3", LeadDays);

                                        //submit dic
                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox1"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox1", _price);
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox1"] = _price;

                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox2"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox2", _discount);
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox2"] = _discount;

                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox3"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox3", LeadDays);
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox3"] = LeadDays;
                                    }


                                    string _comments = "";
                                    if (_item.LineItemComment.Value.Trim() != "")
                                    {
                                        _comments = _item.LineItemComment.Value.Trim().Replace("\r\n", " ");
                                        _comments = _item.LineItemComment.Value.Trim().Replace("\n", " ");
                                        _comments = _item.LineItemComment.Value.Trim().Replace("&nbsp;", " ");
                                        _comments += cPrice_Comments.TrimEnd(',');//added by kalpita on 28/08/2020
                                        //  _comments += "Quoted Qty: " + _item.Quantity + ", Quoted Unit: " + _item.MeasureUnitQualifier;
                                    }
                                    else
                                    {
                                        _comments += cPrice_Comments.TrimEnd(',');//added by kalpita on 28/08/2020
                                    }
                                    if (_comments.Length > 240) _comments = _comments.Substring(0, 240);//added by kalpita on 26/11/2024
                                    if (i <= 9)
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox5", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox5"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox5", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                        else dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl0" + i + "%24TextBox5"] = Uri.EscapeDataString(_comments).Replace("%20", "+");
                                    }
                                    else
                                    {
                                        dctPostDataValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox5", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                        if (!dctSubmitValues.ContainsKey("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox5"))
                                            dctSubmitValues.Add("ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox5", Uri.EscapeDataString(_comments).Replace("%20", "+"));
                                        else
                                            dctSubmitValues["ctl00%24MainContent%24GridView1%24ctl" + i + "%24TextBox5"] = Uri.EscapeDataString(_comments).Replace("%20", "+");
                                    }

                                    i++;
                                }
                                else throw new Exception("Item " + _td[1].InnerText.Trim() + " not found on website for ref no" + UCRefNo);
                            }
                        }
                        i = 2;
                    }
                }
            }
            LogText = "Filling Item Details completed";
        }

        public bool SubmitQuotation(string QuoteRef, string URL, string MessageNumber)
        {
            bool result = false;
            try
            {
                if (IsSubmitQuote)
                {
                    dctPostDataValues.Clear();
                    dctPostDataValues = dctSubmitValues;

                    if (_httpWrapper._dctStateData.ContainsKey("__EVENTTARGET"))
                        dctPostDataValues.Add("__EVENTTARGET", _httpWrapper._dctStateData["__EVENTTARGET"]);

                    if (_httpWrapper._dctStateData.ContainsKey("__EVENTARGUMENT"))
                        dctPostDataValues.Add("__EVENTARGUMENT", _httpWrapper._dctStateData["__EVENTARGUMENT"]);

                    if (_httpWrapper._dctStateData.ContainsKey("__LASTFOCUS"))
                        dctPostDataValues.Add("__LASTFOCUS", _httpWrapper._dctStateData["__LASTFOCUS"]);

                    if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATE"))
                        dctPostDataValues.Add("__VIEWSTATE", _httpWrapper._dctStateData["__VIEWSTATE"]);

                    if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATEGENERATOR"))
                        dctPostDataValues.Add("__VIEWSTATEGENERATOR", _httpWrapper._dctStateData["__VIEWSTATEGENERATOR"]);

                    if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATEENCRYPTED"))
                        dctPostDataValues.Add("__VIEWSTATEENCRYPTED", _httpWrapper._dctStateData["__VIEWSTATEENCRYPTED"]);

                    if (_httpWrapper._dctStateData.ContainsKey("__EVENTVALIDATION"))
                        dctPostDataValues.Add("__EVENTVALIDATION", _httpWrapper._dctStateData["__EVENTVALIDATION"]);

                    dctPostDataValues.Add("ctl00%24MainContent%24txtTotRef", Uri.EscapeDataString(AAGRefNo));
                    dctPostDataValues.Add("ctl00%24MainContent%24btnUpdate", "Submit+Quotation");
                    HtmlNode _hHelp = _httpWrapper.GetElement("input", "id", "MainContent_txtHelp");
                    if (_hHelp != null)
                    {
                        dctPostDataValues.Add("ctl00%24MainContent%24txtHelp", _hHelp.GetAttributeValue("value", "").Trim());
                    }
                    else throw new Exception("Help textbox not found");

                    //string[] ArrURL = URL.Split('=');
                    //URL = ArrURL[0] + "=" + HttpUtility.UrlDecode(ArrURL[1]);

                    //  URL = MessageNumber;
                    string[] ArrURL = URL.Split('=');
                    //URL = ArrURL[0] + "=" + HttpUtility.UrlEncode(ArrURL[1]);

                    _httpWrapper.Referrer = MessageNumber;
                    if (PostURL("input", "id", "MainContent_txtFrom"))
                    {
                        URL = MessageNumber.Replace("Edit", "View");
                        _httpWrapper.ContentType = "";

                        if (LoadURL("input", "id", "MainContent_txtTotRef"))
                        {
                            #region  // Commented By Sanjita on 29-NOV-18 //
                            //string _chkReadonly = _httpWrapper.GetElement("input", "id", "MainContent_txtTotRef").GetAttributeValue("readonly", "");
                            //if (_chkReadonly == "readonly")
                            //{
                            //    result = true;
                            //}                            
                            #endregion

                            string updatedQuoteRef = _httpWrapper.GetElement("input", "id", "MainContent_txtTotRef").GetAttributeValue("value", "");
                            HtmlNode quoteStatus = _httpWrapper.GetElement("span", "id", "MainContent_lblStatus");
                            if (quoteStatus != null)
                            {
                                string strQuoteStatus = convert.ToString(quoteStatus.InnerText);
                                if (updatedQuoteRef == QuoteRef.Trim() && strQuoteStatus.Trim().ToUpper() == "OPEN")
                                {
                                    result = true;
                                }
                            }
                            else
                            {
                                quoteStatus = _httpWrapper.GetElement("span", "id", "MainContent_lblSubmit");
                                if (quoteStatus != null)
                                {
                                    string strQuoteStatus = convert.ToString(quoteStatus.InnerText);
                                    if (updatedQuoteRef == QuoteRef.Trim() && strQuoteStatus.Trim().ToUpper().Contains("SUBMIT"))
                                    {
                                        result = true;
                                    }
                                }
                            }
                        }
                    }
                }
                else result = true;
            }
            catch (Exception ex)
            {
                LogText = "Exception while submitting quote; " + ex.Message;
                LogText = ex.StackTrace;
                //LogText = "Exception while submitting quote: " + ex.GetBaseException().Message.ToString(); result = false;
                throw ex;
            }
            return result;
        }

        public bool DeclineQuotation(string QuoteRef)
        {
            bool result = false;
            try
            {
                dctPostDataValues.Clear();
                dctPostDataValues = dctSubmitValues;

                if (_httpWrapper._dctStateData.ContainsKey("__EVENTTARGET"))
                    dctPostDataValues.Add("__EVENTTARGET", _httpWrapper._dctStateData["__EVENTTARGET"]);

                if (_httpWrapper._dctStateData.ContainsKey("__EVENTARGUMENT"))
                    dctPostDataValues.Add("__EVENTARGUMENT", _httpWrapper._dctStateData["__EVENTARGUMENT"]);

                if (_httpWrapper._dctStateData.ContainsKey("__LASTFOCUS"))
                    dctPostDataValues.Add("__LASTFOCUS", _httpWrapper._dctStateData["__LASTFOCUS"]);

                if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATE"))
                    dctPostDataValues.Add("__VIEWSTATE", _httpWrapper._dctStateData["__VIEWSTATE"]);

                if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATEGENERATOR"))
                    dctPostDataValues.Add("__VIEWSTATEGENERATOR", _httpWrapper._dctStateData["__VIEWSTATEGENERATOR"]);

                if (_httpWrapper._dctStateData.ContainsKey("__VIEWSTATEENCRYPTED"))
                    dctPostDataValues.Add("__VIEWSTATEENCRYPTED", _httpWrapper._dctStateData["__VIEWSTATEENCRYPTED"]);

                if (_httpWrapper._dctStateData.ContainsKey("__EVENTVALIDATION"))
                    dctPostDataValues.Add("__EVENTVALIDATION", _httpWrapper._dctStateData["__EVENTVALIDATION"]);

                dctPostDataValues.Add("ctl00%24MainContent%24txtTotRef", Uri.EscapeDataString(AAGRefNo));
                dctPostDataValues.Add("ctl00%24MainContent%24btnCancel", "Can%27t+Quote");
                HtmlNode _hHelp = _httpWrapper.GetElement("input", "id", "MainContent_txtHelp");
                if (_hHelp != null)
                {
                    dctPostDataValues.Add("ctl00%24MainContent%24txtHelp", _hHelp.GetAttributeValue("value", "").Trim());
                }
                else throw new Exception("Help textbox not found");

                string[] ArrURL = URL.Split('=');
                URL = ArrURL[0] + "=" + HttpUtility.UrlEncode(ArrURL[1]);
                _httpWrapper.Referrer = URL;
                if (PostURL("input", "id", "MainContent_txtFrom"))
                {
                    URL = MessageNumber.Replace("Edit", "View");
                    _httpWrapper.ContentType = "";

                    if (LoadURL("input", "id", "MainContent_txtTotRef"))
                    {
                        string updatedQuoteRef = _httpWrapper.GetElement("input", "id", "MainContent_txtTotRef").GetAttributeValue("value", "");
                        HtmlNode quoteStatus = _httpWrapper.GetElement("span", "id", "MainContent_lblStatus");
                        if (quoteStatus != null)
                        {
                            string strQuoteStatus = convert.ToString(quoteStatus.InnerText);
                            if (updatedQuoteRef == QuoteRef.Trim() && strQuoteStatus.Trim().ToUpper() == "OPEN")
                            {
                                result = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogText = "Exception while declining quote; " + ex.Message;
                LogText = ex.StackTrace;
                throw ex;
            }
            return result;
        }

        public void WriteErrorLogQuote_With_Screenshot(string AuditMsg, string _File)
        {
            LogText = AuditMsg;
            string eFile = PrintScreenPath + "\\Navigation_" + this.currDocType + "Error_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".png";
            if (!PrintScreen(eFile)) eFile = _File;

            //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", AuditMsg, BuyerCode, SupplierCode, AuditPath);
            string Server = SERVER;
            string Processor = MODULE_NAME;
            try
            {
                string auditPath = AuditPath;
                if (!Directory.Exists(auditPath)) Directory.CreateDirectory(auditPath);

                string auditData = "";
                if (auditData.Trim().Length > 0) auditData += Environment.NewLine;
                auditData += "0|"; // Buyer ID
                auditData += "0|"; // Supplier ID            
                auditData += "Navibulgar_HTTP_Processor|";
                auditData += Path.GetFileName(_File) + "|";
                auditData += UCRefNo + "|";
                auditData += "Error" + "|";
                auditData += DateTime.Now.ToString("yy-MM-dd HH:mm") + " : " + AuditMsg + "|";
                auditData += "0|"; // Linkid
                auditData += Server + "|"; // Server 
                auditData += BuyerCode + "|"; // Server 
                auditData += SupplierCode + "|"; // Server 
                auditData += Processor; // Processor 

                if (auditData.Trim().Length > 0)
                {
                    File.WriteAllText(AuditPath + "\\Navibulgar_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".txt", auditData);
                    System.Threading.Thread.Sleep(5);
                }
            }
            catch (Exception ex)
            {
                LogText = ex.GetBaseException().ToString();
                _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);
            }
            _fileStore.MoveFile(currentMtmlFile, FileMoveAction.MoveToInProcess, FileMoveAction.MoveToError, currentMtmlFile, currentMoveDirectory, dctAppSettings);
            //if (!Directory.Exists(MTML_QuotePath + "\\Error")) Directory.CreateDirectory(MTML_QuotePath + "\\Error");
            //File.Move(MTML_QuotePath + "\\" + _File, MTML_QuotePath + "\\Error\\" + _File);
        }

        public void LoadInterchangeDetails()
        {
            try
            {
                IsDecline = false; dtExpDate = ""; BuyerCode = ""; SupplierCode = ""; currDocType = ""; MessageNumber = ""; LeadDays = ""; MsgNumber = ""; MsgRefNumber = ""; UCRefNo = ""; IsAltItemAllowed = 0; IsPriceAveraged = 0; IsUOMChanged = 0;
                AAGRefNo = ""; LesRecordID = ""; BuyerName = ""; BuyerPhone = ""; BuyerEmail = ""; BuyerFax = ""; supplierName = ""; supplierPhone = ""; supplierEmail = ""; supplierFax = ""; VesselName = ""; PortName = ""; PortCode = "";
                SupplierComment = ""; PayTerms = ""; PackingCost = "0"; FreightCharge = "0"; TotalLineItemsAmount = "0"; Allowance = "0"; GrandTotal = "0"; BuyerTotal = "0"; TaxCost = "0"; DtDelvDate = "";

                LogText = "Started Loading interchange object.";
                if (_interchange != null)
                {
                    if (_interchange.Recipient != null)
                        BuyerCode = _interchange.Recipient;

                    if (_interchange.Sender != null)
                        SupplierCode = _interchange.Sender;

                    if (_interchange.DocumentHeader.DocType != null)
                        currDocType = _interchange.DocumentHeader.DocType;

                    if (_interchange.DocumentHeader != null)
                    {
                        if (_interchange.DocumentHeader.IsDeclined)
                            IsDecline = _interchange.DocumentHeader.IsDeclined;

                        if (_interchange.DocumentHeader.MessageNumber != null)
                            MessageNumber = _interchange.DocumentHeader.MessageNumber;

                        if (_interchange.DocumentHeader.LeadTimeDays != null)
                            LeadDays = _interchange.DocumentHeader.LeadTimeDays;

                        Currency = _interchange.DocumentHeader.CurrencyCode;

                        MsgNumber = _interchange.DocumentHeader.MessageNumber;
                        MsgRefNumber = _interchange.DocumentHeader.MessageReferenceNumber;

                        if (_interchange.DocumentHeader.IsAltItemAllowed != null) IsAltItemAllowed = Convert.ToInt32(_interchange.DocumentHeader.IsAltItemAllowed);
                        if (_interchange.DocumentHeader.IsPriceAveraged != null) IsPriceAveraged = Convert.ToInt32(_interchange.DocumentHeader.IsPriceAveraged);
                        if (_interchange.DocumentHeader.IsUOMChanged != null) IsUOMChanged = Convert.ToInt32(_interchange.DocumentHeader.IsUOMChanged);


                        for (int i = 0; i < _interchange.DocumentHeader.References.Count; i++)
                        {
                            if (_interchange.DocumentHeader.References[i].Qualifier == ReferenceQualifier.UC)
                                UCRefNo = _interchange.DocumentHeader.References[i].ReferenceNumber.Trim();
                            else if (_interchange.DocumentHeader.References[i].Qualifier == ReferenceQualifier.AAG)
                                AAGRefNo = _interchange.DocumentHeader.References[i].ReferenceNumber.Trim();
                        }
                    }
                    if (_interchange.BuyerSuppInfo != null)
                    {
                        LesRecordID = Convert.ToString(_interchange.BuyerSuppInfo.RecordID);
                    }

                    #region read interchange party addresses

                    for (int j = 0; j < _interchange.DocumentHeader.PartyAddresses.Count; j++)
                    {
                        if (_interchange.DocumentHeader.PartyAddresses[j].Qualifier == PartyQualifier.BY)
                        {
                            BuyerName = _interchange.DocumentHeader.PartyAddresses[j].Name;
                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts != null)
                            {
                                if (_interchange.DocumentHeader.PartyAddresses[j].Contacts.Count > 0)
                                {
                                    if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList.Count > 0)
                                    {
                                        for (int k = 0; k < _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList.Count; k++)
                                        {
                                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Qualifier == CommunicationMethodQualifiers.TE)
                                                BuyerPhone = _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Number;
                                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Qualifier == CommunicationMethodQualifiers.EM)
                                                BuyerEmail = _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Number;
                                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Qualifier == CommunicationMethodQualifiers.FX)
                                                BuyerFax = _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Number;
                                        }
                                    }
                                }
                            }
                        }

                        else if (_interchange.DocumentHeader.PartyAddresses[j].Qualifier == PartyQualifier.VN)
                        {
                            supplierName = _interchange.DocumentHeader.PartyAddresses[j].Name;
                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts != null)
                            {
                                if (_interchange.DocumentHeader.PartyAddresses[j].Contacts.Count > 0)
                                {
                                    if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList.Count > 0)
                                    {
                                        for (int k = 0; k < _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList.Count; k++)
                                        {
                                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Qualifier == CommunicationMethodQualifiers.TE)
                                                supplierPhone = _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Number;
                                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Qualifier == CommunicationMethodQualifiers.EM)
                                                supplierEmail = _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Number;
                                            if (_interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Qualifier == CommunicationMethodQualifiers.FX)
                                                supplierFax = _interchange.DocumentHeader.PartyAddresses[j].Contacts[0].CommunMethodList[k].Number;
                                        }
                                    }
                                }
                            }
                        }

                        else if (_interchange.DocumentHeader.PartyAddresses[j].Qualifier == PartyQualifier.UD)
                        {
                            VesselName = _interchange.DocumentHeader.PartyAddresses[j].Name;
                            if (_interchange.DocumentHeader.PartyAddresses[j].PartyLocation.Berth != "")
                                PortName = _interchange.DocumentHeader.PartyAddresses[j].PartyLocation.Berth;

                            if (_interchange.DocumentHeader.PartyAddresses[j].PartyLocation.Port != null)
                                PortCode = _interchange.DocumentHeader.PartyAddresses[j].PartyLocation.Port;
                        }
                    }

                    #endregion

                    #region read comments

                    if (_interchange.DocumentHeader.Comments != null)
                    {
                        for (int i = 0; i < _interchange.DocumentHeader.Comments.Count; i++)
                        {
                            if (_interchange.DocumentHeader.Comments[i].Qualifier == CommentTypes.SUR)
                                SupplierComment = _interchange.DocumentHeader.Comments[i].Value;
                            else if (_interchange.DocumentHeader.Comments[i].Qualifier == CommentTypes.ZTP)
                                PayTerms = _interchange.DocumentHeader.Comments[i].Value;
                        }
                    }

                    #endregion

                    #region read Line Items

                    if (_interchange.DocumentHeader.LineItemCount > 0)
                    {
                        _lineitem = _interchange.DocumentHeader.LineItems;
                    }

                    #endregion

                    #region read Interchange Monetory Amount

                    if (_interchange.DocumentHeader.MonetoryAmounts != null)
                    {
                        for (int i = 0; i < _interchange.DocumentHeader.MonetoryAmounts.Count; i++)
                        {
                            if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.PackingCost_106)
                                PackingCost = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                            else if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.FreightCharge_64)
                                FreightCharge = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                            else if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.TotalLineItemsAmount_79)
                                TotalLineItemsAmount = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                            else if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.AllowanceAmount_204)
                                Allowance = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                            else if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.GrandTotal_259)
                                GrandTotal = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                            else if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.BuyerItemTotal_90)//16-12-2017
                                BuyerTotal = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                            else if (_interchange.DocumentHeader.MonetoryAmounts[i].Qualifier == MonetoryAmountQualifier.TaxCost_99)//23/09/2020 by kalpita
                                TaxCost = _interchange.DocumentHeader.MonetoryAmounts[i].Value.ToString();
                        }
                    }

                    #endregion

                    #region read date time period

                    if (_interchange.DocumentHeader.DateTimePeriods != null)
                    {
                        for (int i = 0; i < _interchange.DocumentHeader.DateTimePeriods.Count; i++)
                        {
                            if (_interchange.DocumentHeader.DateTimePeriods[i].Qualifier == DateTimePeroidQualifiers.DocumentDate_137)
                            {
                                if (_interchange.DocumentHeader.DateTimePeriods[i].Value != null)
                                { DateTime dtDocDate = FormatMTMLDate(_interchange.DocumentHeader.DateTimePeriods[i].Value); }
                            }

                            else if (_interchange.DocumentHeader.DateTimePeriods[i].Qualifier == DateTimePeroidQualifiers.DeliveryDate_69)
                            {
                                if (_interchange.DocumentHeader.DateTimePeriods[i].Value != null)
                                {
                                    DateTime dtDelDate = FormatMTMLDate(_interchange.DocumentHeader.DateTimePeriods[i].Value);
                                    if (dtDelDate != DateTime.MinValue)
                                    {
                                        DtDelvDate = dtDelDate.ToString("MM/dd/yyyy");
                                    }
                                    if (dtDelDate == null)
                                    {
                                        DateTime dt = FormatMTMLDate(DateTime.Now.AddDays(Convert.ToDouble(LeadDays)).ToString());
                                        if (dt != DateTime.MinValue)
                                        {
                                            DtDelvDate = dt.ToString("MM/dd/yyyy");
                                        }
                                    }
                                }
                            }

                            if (_interchange.DocumentHeader.DateTimePeriods[i].Qualifier == DateTimePeroidQualifiers.ExpiryDate_36)
                            {
                                if (_interchange.DocumentHeader.DateTimePeriods[i].Value != null)
                                {
                                    DateTime ExpDate = FormatMTMLDate(_interchange.DocumentHeader.DateTimePeriods[i].Value);
                                    if (ExpDate != DateTime.MinValue)
                                    {
                                        dtExpDate = ExpDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);//15-3-18
                                        //  dtExpDate = ExpDate.ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                                    }
                                }
                            }
                        }
                    }

                    #endregion
                    LogText = "stopped Loading interchange object.";
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Exception on LoadInterchangeDetails : " + ex.GetBaseException().ToString());
            }
        }

        public DateTime FormatMTMLDate(string DateValue)
        {
            DateTime Dt = DateTime.MinValue;
            if (DateValue != null && DateValue != "")
            {
                if (DateValue.Length > 5)
                {
                    int year = Convert.ToInt32(DateValue.Substring(0, 4));
                    int Month = Convert.ToInt32(DateValue.Substring(4, 2));
                    int day = Convert.ToInt32(DateValue.Substring(6, 2));
                    Dt = new DateTime(year, Month, day);
                }
            }
            return Dt;
        }
        #endregion

        #region PO
        public void ProcessPO()
        {
            try
            {
                this.currDocType = "PO";
                LogText = "PO processing started.";
                List<string> _lstNewPOs = GetNewPOs();
                if (_lstNewPOs.Count > 0)
                {
                    GetPODetails(_lstNewPOs);
                }
                else
                {
                    LogText = "No new PO found.";
                }
                LogText = "PO processing stopped.";
            }
            catch (Exception e)
            {
                WriteErrorLog_With_Screenshot("Unable to process due to " + e.GetBaseException().ToString(), "LeS-1004:");
            }
        }


        public List<string> GetNewPOs()
        {
            List<string> _lstNewPOs = new List<string>();
            List<string> slProcessedItem = GetProcessedItems(eActions.PO);
            _lstNewPOs.Clear();
            _httpWrapper._CurrentDocument.LoadHtml(_httpWrapper._CurrentResponseString);

            if (_httpWrapper._dctStateData.Count > 0)
            {
                dctPostDataValues.Clear();
                dctPostDataValues.Add("__EVENTTARGET", "ctl00%24MainContent%24txtFrom");
                dctPostDataValues.Add("__EVENTARGUMENT", _httpWrapper._dctStateData["__EVENTARGUMENT"]);
                dctPostDataValues.Add("__LASTFOCUS", "");
                dctPostDataValues.Add("__VIEWSTATE", _httpWrapper._dctStateData["__VIEWSTATE"]);
                dctPostDataValues.Add("__VIEWSTATEGENERATOR", _httpWrapper._dctStateData["__VIEWSTATEGENERATOR"]);
                dctPostDataValues.Add("__VIEWSTATEENCRYPTED", "");
                dctPostDataValues.Add("__EVENTVALIDATION", _httpWrapper._dctStateData["__EVENTVALIDATION"]);
                if (!_httpWrapper._AddRequestHeaders.ContainsKey("Origin")) _httpWrapper._AddRequestHeaders.Add("Origin", @$"{cSiteURL}");
                _httpWrapper.Referrer = "";
            }

            URL = $"{cSiteURL}suppliers/UnfinOrd.aspx";
            if (PostURL("table", "id", "MainContent_GridView1"))
            {
                HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@id='MainContent_GridView1']//tr[@onmouseout]");
                if (_nodes != null)
                {
                    if (_nodes.Count > 0)
                    {
                        foreach (HtmlNode _row in _nodes)
                        {
                            HtmlNodeCollection _rowData = _row.ChildNodes;
                            string VRNo = _rowData[1].InnerText.Trim();
                            string Date = _rowData[2].InnerText.Trim();
                            string Vessel = _rowData[3].InnerText.Trim();
                            string Port = _rowData[4].InnerText.Trim();
                            string _url = _row.GetAttributeValue("onclick", "").Trim();
                            if (_url.Contains(';'))
                            {
                                string[] _arrUrl = _url.Split(';');
                                _url = _arrUrl[1].Replace("&#39", "");
                            }
                            string[] _dateArr = Date.Split('/');
                            if (Convert.ToInt32(_dateArr[2]) == DateTime.Now.Year)
                            {
                                string _guid = VRNo + "|" + Date + "|" + Vessel + "|" + Port;
                                if (!_lstNewPOs.Contains(VRNo + "|" + Date + "|" + Vessel + "|" + Port + "|" + _url) && !slProcessedItem.Contains(_guid))
                                {
                                    _lstNewPOs.Add(VRNo + "|" + Date + "|" + Vessel + "|" + Port + "|" + _url);
                                }
                            }
                        }
                    }
                    else
                        LogText = "No new POs found.";
                }
            }
            return _lstNewPOs;
        }

        public void GetPODetails(List<string> _lstNewPOs)
        {
            foreach (string strPO in _lstNewPOs)
            {
                try
                {
                    string cRefNo = "", cVessel = "", cPort = "", cOrderDate = "", cVslEtaDate = "", cComment = "";
                    string[] lst = strPO.Split('|');
                    UCRefNo = this.VRNO = lst[0];
                    this.Vessel = lst[1]; this.PODate = lst[2];
                    this.Port = lst[3];

                    URL = lst[4];
                    string _POguid = this.VRNO + "|" + this.PODate + "|" + this.Vessel + "|" + this.Port;

                    URL = $"{cSiteURL}suppliers/" + URL;
                    LoadURL("input", "id", "MainContent_btnConfirm");
                    LogText = "Processing PO for Ref No. " + this.VRNO;

                    string eFile = Path.Combine(this.ImgPath, this.VRNO.Replace("/", "_") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + SupplierCode + ".png");
                    if (!PrintScreen(eFile)) eFile = "";
                    //LeSXML.LeSXML _lesXml = new LeSXML.LeSXML();
                    cPOCurrency = "";


                    MTMLInterchange interchange = new MTMLInterchange();
                    DocHeader docHeader = new DocHeader();
                    DateTimePeriodCollection PODateTimePeriod = new DateTimePeriodCollection();

                    interchange.PreparationDate = DateTime.Now.ToString("yyyy-MMM-dd");
                    interchange.PreparationTime = DateTime.Now.ToString("HH:mm");
                    interchange.ControlReference = DateTime.Now.ToString("yyyyMMddHHmmss");
                    interchange.Identifier = DateTime.Now.ToString("yyyyMMddHHmmss");
                    interchange.Recipient = SUPPLIER_LINK_CODE;
                    interchange.Sender = BUYER_LINK_CODE;
                    LineItemCollection lineItems = new LineItemCollection();
                    PartyCollection parties = new PartyCollection();

                    HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@id='MainContent_GridView2']//tr");
                    if (_nodes != null)
                    {
                        if (_nodes.Count >= 2)
                        {
                            int i = 0;
                            foreach (HtmlNode _row in _nodes)
                            {
                                if (i > 0)
                                {
                                    HtmlNodeCollection _rowData = _row.ChildNodes;
                                    UCRefNo = cRefNo = Convert.ToString(_rowData[1].InnerText).Trim();
                                    cOrderDate = Convert.ToString(_rowData[2].InnerText).Trim();
                                    cPort = Convert.ToString(_rowData[3].InnerText).Trim();
                                    cVessel = Convert.ToString(_rowData[4].InnerText).Trim();
                                    cVslEtaDate = Convert.ToString(_rowData[5].InnerText).Trim().Replace("&nbsp;", "");
                                    cComment = Convert.ToString(_rowData[6].InnerText).Trim().Replace("&nbsp;", "");
                                }
                                i++;
                            }
                        }
                    }
                    //document date
                    if (cOrderDate != "" && cOrderDate != "-")
                    {
                        DateTime dt = DateTime.MinValue;
                        //DateTime.TryParseExact(cOrderDate, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out dt);
                        //_lesXml.Date_Document = dt.ToString("yyyyMMdd");

                        DateTime date = DateTime.ParseExact(cOrderDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        string _poetadate = date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                        DateTime _dt = DateTime.MinValue;
                        DateTimePeriod _dtDocDate = new DateTimePeriod();
                        DateTime.TryParseExact(_poetadate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _dt);
                        _dtDocDate.FormatQualifier = DateTimeFormatQualifiers.CCYYMMDD_102;
                        _dtDocDate.Qualifier = DateTimePeroidQualifiers.DocumentDate_137;
                        _dtDocDate.Value = _dt.ToString("yyyyMMdd");
                        PODateTimePeriod.Add(_dtDocDate);
                    }

                    if (GetPOHeader(ref docHeader, eFile, cVslEtaDate))
                    {
                        if (GetPOItems(ref lineItems))
                        {
                            docHeader.LineItems = lineItems;
                            docHeader.LineItemCount = lineItems.Count();

                            if (GetPOAddress(ref parties, cVslEtaDate, cVessel, cPort))
                            {
                                docHeader.PartyAddresses = parties;
                                string pofile = "PO_" + this.VRNO.Replace("/", "_") + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xml";
                                if (lineItems.Count() > 0)
                                {
                                    docHeader.CurrencyCode = cPOCurrency;          
                                    string filePath = Path.Combine(RFQPath, pofile);
                                    interchange.DocumentHeader = docHeader;
                                    MTMLClass _class = new MTMLClass();
                                    LogText = "Creating MTML RFQ";
                                    _class.Create(interchange, filePath);

                                    if (File.Exists(filePath))
                                    {
                                        LogText = pofile + " downloaded successfully.";
                                        LogText = "";
                                        //CreateAuditFile(xmlfile, "Navibulgar_HTTP_Processor", VRNO, "Downloaded", xmlfile + " downloaded successfully.", BuyerCode, SupplierCode, AuditPath);
                                        AuditMessageData.CreateAuditFile(pofile, MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "PO Generated Successfully for " + UCRefNo);

                                        //if ((this.VRNO + "|" + this.Vessel + "|" + this.Port).Length > 0)
                                        //    File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + this.Domain + "_PO_" + this.SupplierCode + "_GUID.txt", this.VRNO + "|" + this.Vessel + "|" + this.Port + Environment.NewLine);

                                        _fileStore.AppendText($"DownloadedPO_{cBranchList}.txt", this.VRNO + "|" + this.Vessel + "|" + this.Port + Environment.NewLine);
                                    }
                                    else
                                    {
                                        LogText = "Unable to download file " + pofile;
                                        string filename = PrintScreenPath + "\\Navibulgar_POError_" + DateTime.Now.ToString("ddMMyyyHHmmssfff") + ".png";
                                        //CreateAuditFile(filename, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1004:Unable to process file " + pofile + " for ref " + VRNO, BuyerCode, SupplierCode, AuditPath);
                                        AuditMessageData.CreateAuditFile(pofile, MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "PO Generated Successfully for " + UCRefNo);
                                        if (PrintScreen(filename)) filename = "";
                                    }
                                }
                            }
                            else
                            {
                                LogText = "Unable to get address details";
                                //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1040:Unable to get details-address Field(s) not present", BuyerCode, SupplierCode, AuditPath);
                                AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "LeS-1040:Unable to get details-address Field(s) not present");

                            }
                        }
                        else
                        {
                            LogText = "Unable to get PO item details";
                            //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1040:Unable to get details-PO item Field(s) not present", BuyerCode, SupplierCode, AuditPath);
                            AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "LeS-1040:Unable to get details-PO item Field(s) not present");

                        }
                    }
                    else
                    {
                        LogText = "Unable to get PO header details";
                        //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", "LeS-1040:Unable to get details-PO header Field(s) not present", BuyerCode, SupplierCode, AuditPath);
                        AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, "LeS-1040:Unable to get details-PO header Field(s) not present");
                    }
                

                
                }
                catch (Exception ex)
                {
                    WriteErrorLog_With_Screenshot("Unable to filter details for " + VRNO + " due to " + ex.GetBaseException().Message.ToString(), "LeS-1006:");
                }
            }
        }

        public bool GetPOHeader(ref DocHeader docHeader, string eFile, string cVslEtaDate)
        {
            bool isResult = false;
            //string cRefNo = "", cVessel = "", cPort = "", cOrderDate = "", cVslEtaDate = "", cComment = "";
            LogText = "Start Getting Header details";
            try
            {


                if (docHeader.References == null)
                    docHeader.References = new ReferenceCollection();

                if (docHeader.Comments == null)
                    docHeader.Comments = new CommentsCollection();

                if (docHeader.Equipment == null)
                    docHeader.Equipment = new Equipment();
                docHeader.DocType = "ORDER";

                docHeader.MessageReferenceNumber = URL;

                //_lesXml.DocID = DateTime.Now.ToString("yyyyMMddhhmmss");
                //_lesXml.Created_Date = DateTime.Now.ToString("yyyyMMdd");
                //_lesXml.Doc_Type = "PO";
                //_lesXml.Dialect = "Navigation Maritime Bulgare";
                //_lesXml.Version = "1";
                //_lesXml.Date_Document = DateTime.Now.ToString("yyyyMMdd");
                //_lesXml.Date_Preparation = DateTime.Now.ToString("yyyyMMdd");
                //_lesXml.Sender_Code = BuyerCode;
                //_lesXml.Recipient_Code = SupplierCode;
                //_lesXml.Active = "1";
                //_lesXml.DocLinkID = URL;

                //if (File.Exists(eFile))
                //    _lesXml.OrigDocFile = Path.GetFileName(eFile);
                //_lesXml.DocReferenceID = cRefNo;
                //_lesXml.BuyerRef = cRefNo;
                //_lesXml.Vessel = cVessel;
                //_lesXml.PortName = cPort;
                //_lesXml.Remark_Header = cComment;


                //_lesXml.Total_LineItems_Discount = "0.00";
                //_lesXml.Total_Additional_Discount = "0.00";


                LogText = "Getting Header details completed successfully.";
                isResult = true;
                return isResult;
            }
            catch (Exception ex)
            {
                LogText = "Unable to get header details." + ex.GetBaseException().ToString(); isResult = false;
                return isResult;
            }
        }

        public bool GetPOItems(ref LineItemCollection lineItemss)
        {
            bool isResult = false;
            string EquipRemarks = "", cCurrency = "";
            try
            {
                LogText = "Start Getting LineItem details";

                HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//table[@id='MainContent_GridView3']//tr");
                //HtmlNodeCollection _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//span[starts-with(@id,'MainContent_Repeater1_Label1_')]/ancestor::tr");
                //var _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//span[starts-with(@id,'MainContent_Repeater1_Label1_')]/parent::td/parent::tr");
                //var _nodes = _httpWrapper._CurrentDocument.DocumentNode.SelectNodes("//span[starts-with(@id,'MainContent_Repeater1_Label1_')]/ancestor::tr");

                if (_nodes != null)
                {
                    if (_nodes.Count >= 2)
                    {
                        int i = 0;
                        foreach (HtmlNode _row in _nodes)
                        {
                            LineItem _item = new LineItem();
                            try
                            {
                                HtmlNodeCollection _rowData = _row.ChildNodes;
                                if (!_rowData[1].InnerText.Trim().Contains("Item Code"))
                                {
                                    i += 1;
                                    _item.Number = Convert.ToString(i);
                                    _item.OriginatingSystemRef = Convert.ToString(i);
                                    _item.SYS_ITEMNO = i;
                                    _item.OriginatingSystemRef = Convert.ToString(_rowData[1].InnerText).Trim();
                                    _item.Description = Convert.ToString(_rowData[3].InnerText).Trim().Replace("&nbsp;", "");
                                    _item.MeasureUnitQualifier = Convert.ToString(_rowData[5].InnerText).Trim();
                                    _item.Quantity = Convert.ToDouble(_rowData[4].InnerText);
                                    _item.Description = Convert.ToString(_rowData[2].InnerText).Trim();
                                    _item.Discount_Value = 0.0;
                                    _item.PriceList.Add(new PriceDetails(PriceDetailsTypeCodes.Quoted_QT, PriceDetailsTypeQualifiers.LIST, Convert.ToSingle(_rowData[6].InnerText)));
                                    _item.PriceList.Add(new PriceDetails(PriceDetailsTypeCodes.Quoted_QT, PriceDetailsTypeQualifiers.GRP, Convert.ToSingle(_rowData[6].InnerText)));
                                    //_item.LeadDays = "0";
                                    cCurrency += Convert.ToString(_rowData[6].InnerText).Trim() + ",";
                                    _item.Identification = Convert.ToString(_rowData[1].InnerText).Trim();
                                    //var check = Convert.ToString(_rowData[11].InnerText).Trim();
                                    lineItemss.Add(_item);

                                    //----------------------------------------------------------------
                      
                            
                                }
                               
                            }
                            catch (Exception ex)
                            { LogText = ex.GetBaseException().ToString(); }
                        }
                        //_lesXml.Total_LineItems = Convert.ToString(_lesXml.LineItems.Count);
                        cPOCurrency = cCurrency.Split(',').Distinct().First();

                        isResult = true;
                    }
                    else isResult = false;
                }
                else isResult = false;

                LogText = "Getting LineItem details successfully";
                return isResult;
            }
            catch (Exception ex)
            {
                LogText = "Exception while getting PO Items: " + ex.GetBaseException().ToString(); isResult = false; return isResult;
            }
        }

        public bool GetPOAddress(ref PartyCollection collection, string cVslEtaDate, string cVessel, string cPort)
        {
            bool isResult = false;
            try
            {
                //_lesXml.Addresses.Clear();
                //LogText = "Start Getting address details";
                //LeSXML.Address _xmlAdd = new LeSXML.Address();

                //_xmlAdd.Qualifier = "BY";
                //_xmlAdd.AddressName = BuyerName;
                //_lesXml.Addresses.Add(_xmlAdd);

                //_xmlAdd = new LeSXML.Address();
                //_xmlAdd.Qualifier = "VN";
                //_xmlAdd.AddressName = SupplierName;
                //_lesXml.Addresses.Add(_xmlAdd);

                //LogText = "Getting address details successfully";
                
                collection.Clear();

                //Vendor
                Party vendor = new Party();
                vendor.Name = SupplierName;
                vendor.Qualifier = PartyQualifier.VN;
                collection.Add(vendor);

                //Buyer 
                Party buyer = new Party();
                Contact contactBuyer = new Contact();
                buyer.Qualifier = PartyQualifier.BY;
 
                buyer.Name = BuyerName;
                collection.Add(buyer);

                //Vessel
                Party vessel = new Party();
                vessel.Qualifier = PartyQualifier.UD;
                vessel.Name = cVessel;
                //vessel.Identification = "";
                vessel.PartyLocation = new PartyLocation();
                vessel.PartyLocation.Qualifier = PartyLocationQualifier.ZUC;
                vessel.PartyLocation.Port = cPort;

                collection.Add(vessel);
                isResult = true;
                return isResult;
            }
            catch (Exception ex)
            {
                LogText = "Exception while getting address details: " + ex.GetBaseException().ToString(); isResult = false;
                return isResult;
            }
        }

        #endregion


        public void WriteErrorLog_With_Screenshot(string AuditMsg, string ErrorNo)
        {
            LogText = AuditMsg;
            string eFile = Path.Combine(ImgPath , "Navibulgar_" + this.currDocType + "Error_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + SupplierCode + ".png");
            if (!PrintScreen(eFile)) eFile = "";
            //CreateAuditFile(eFile, "Navibulgar_HTTP_Processor", VRNO, "Error", ErrorNo + AuditMsg, BuyerCode, SupplierCode, AuditPath);
            AuditMessageData.CreateAuditFile("", MODULE_NAME, UCRefNo, AuditMessageData.DOWNLOAD_SUCCESS, BuyerCode, SupplierCode, currDocType, AuditMsg);


        }

    }
}
