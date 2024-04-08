using LSExtensionWindowLib;
using LSSERVICEPROVIDERLib;
//using PreviewLetter.Controls;
using ONE1_richTextCtrl;
using Patholab_Common;

using Patholab_DAL_V1;
using Patholab_DAL_V1.Enums;

using System;
using System.Collections.Generic;

using System.Linq;
using System.Windows.Forms;

using Path = System.IO.Path;
using Timer = System.Windows.Forms.Timer;

using System.Windows.Input;
using CrystalDecisions.CrystalReports.Engine;
using System.IO;
using CrystalDecisions.Shared;
using System.ComponentModel;

namespace PreviewLetter
{
    public class PreviewLetterCls
    {
        private DataLayer _dal;
        private INautilusDBConnection _ntlsCon;

        public void runPreviewLetter(long SDG_ID, string U_PDF_PATH, INautilusDBConnection ntlsCon, DataLayer dal)
        {
            try
            {
                if (U_PDF_PATH != null && File.Exists(U_PDF_PATH))
                {
                    var pdf = new PdfViewerFrm(U_PDF_PATH);
                    pdf.ShowDialog();
                }
                else
                {
                    _dal = dal;
                    _ntlsCon = ntlsCon;
                    RunReport(SDG_ID);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error on Load pdf! " + ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.WriteLogFile(ex);
            }
        }

        ReportDocument CR;
        bool IsProxy = false;
        public void RunReport(long sdg_id)
        {
            try
            {
                string serverName;
                string nautilusUserName;
                string nautilusPassword;

                serverName = _ntlsCon.GetServerDetails();
                IsProxy = _ntlsCon.GetServerIsProxy();
                if (IsProxy)
                {
                    nautilusUserName = "";
                    nautilusPassword = "";
                }
                else
                {
                    nautilusUserName = _ntlsCon.GetUsername();
                    nautilusPassword = _ntlsCon.GetPassword();
                }


                var reportPath = _dal.FindBy<PHRASE_HEADER>(ph => ph.NAME.Equals("System Parameters")).FirstOrDefault().PHRASE_ENTRY.Where(pe => pe.PHRASE_NAME.Equals("Preview Letter")).FirstOrDefault().PHRASE_DESCRIPTION;

                if (File.Exists(reportPath))
                {
                    //load
                    CR = new ReportDocument();
                    CR.Load(reportPath);
                }
                else
                {
                    MessageBox.Show("Can't find pdf path from phrase.");
                    return;
                }

                CR.SetParameterValue("sdg id", sdg_id);


                Tables crTables;
                var crTableLoginInfo = new TableLogOnInfo();
                var crConnectionInfo = new ConnectionInfo();

                crConnectionInfo.ServerName = serverName;
                if (IsProxy)
                {
                    crConnectionInfo.IntegratedSecurity = true;
                }
                else
                {
                    crConnectionInfo.UserID = nautilusUserName;
                    crConnectionInfo.Password = nautilusPassword;
                }

                crTables = CR.Database.Tables;
                foreach (Table crTable in crTables)
                {
                    crTableLoginInfo = crTable.LogOnInfo;
                    crTableLoginInfo.ConnectionInfo = crConnectionInfo;
                    crTable.ApplyLogOnInfo(crTableLoginInfo);
                }

                CrystalReportsV1.Form1 f = new CrystalReportsV1.Form1(CR);

                f.ShowDialog();

            }
            catch (Exception e)
            {
                MessageBox.Show("Error on RunReport : " + e.Message);
            }
        }

    }
}
