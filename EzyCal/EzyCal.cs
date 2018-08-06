
using System;
using System.IO;
using System.IO.Ports;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Threading;
using System.Diagnostics;
using System.Timers;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Win32;

namespace EzyCal
{
    public partial class EzyCal : Form
    {
        bool bResultsSavedAccess = false;
        bool bResultsSavedSql = false;
        public enum MSG { INFO, WARNING, ERROR, FAILURE, DEBUG };
        DataTable DataTableBatch;
        public static bool bClickStep = false;
        public static Actuals Actuals = new Actuals();
        public static Actuals Desired = new Actuals();
        public static StepValues StepValues = new StepValues();
        public static Electricals Electricals = new Electricals();
        public static ErrorCounter ErrorCounter = new ErrorCounter();
        public static Counters Counters = new Counters();
        public static Settings Settings = new Settings();
        public static Bench Bench = new Bench();
        public static Meter [] TestBoard = new Meter[50];
        public const int iMaxSteps = 500;
        Button[] buttons = new Button[50];
        Random Random = new Random();
        public static int iMeters = 0;
        DateTime AppStartTime = DateTime.Now;
        public static int iProcedureListRow = 0;
        public static int iProcedureRunRow = 0;
        int iClientNo;
        BackgroundWorker BgWorker;
        string strSuccess = "》";  // a unicode (char*) return 
        const short BOX_ERRCOUNTER = 1;
        const short BOX_UIPFSOURCE = 2;
        const short BOX_REFMTR232 = 7;
        double dCurConst = 1000000000;
        const int iPASS = 11111111; // Marker for pass
        const int iFAIL = 22222222; // Marker for fail
        const double dDEFAULT = 11.11; // Marker for default
        static string strSqlConnection = @"SERVER=AUSYDVS12; DATABASE=FTSU1200; TRUSTED_CONNECTION=FALSE; USER ID=FTS; PASSWORD=FTSPassword; CONNECTION TIMEOUT=10;";

        CtrComm.YCIVCtrClass _CtrComm = new CtrComm.YCIVCtrClass();

        public EzyCal()
        {
            InitializeComponent();

            BgWorker = new BackgroundWorker();
            BgWorker.DoWork += new DoWorkEventHandler(BgWorker_DoWork);
            BgWorker.ProgressChanged += new ProgressChangedEventHandler(BgWorker_Progress);
            BgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BgWorker_Completed);
            BgWorker.WorkerReportsProgress = true;
            BgWorker.WorkerSupportsCancellation = true;

            ToolStripMenuItem_Nor.CheckState = CheckState.Checked;
            ToolStripMenuItem_Continuous.CheckState = CheckState.Checked;
            ToolStripMenuItem_Allow10s.CheckState = CheckState.Checked;

            // Double buffering
            typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, DGV_Results, new object[] { true });

            if (ReadIniFile())
            {
                return;
            }

            for (int iPos = 0; iPos < 50; iPos++)
            {
                TestBoard[iPos] = new Meter();
                TestBoard[iPos].Active = false;
                TestBoard[iPos].Saved = false;
                TestBoard[iPos].Failed = 0;
                TestBoard[iPos].Status = "Status?";
                TestBoard[iPos].MeterType = "MeterType?";
                TestBoard[iPos].MSN = "MSN?";
                TestBoard[iPos].OwnerNo = "OwnerNo?";
                TestBoard[iPos].ContractNo = "ContractNo?";
                TestBoard[iPos].Client = "Client?";
                TestBoard[iPos].ClientNo = "Client?";
                TestBoard[iPos].Firmware = "Firmware?";
                TestBoard[iPos].MAC = "MAC?";
                TestBoard[iPos].DateTime = "DateTime?";
                TestBoard[iPos].Program = "Program?";
                TestBoard[iPos].Ripple = "Ripple?";

                for (int iStep = 0; iStep < iMaxSteps; iStep++)
                {
                    TestBoard[iPos].ErrorV[iStep] = dDEFAULT;
                    TestBoard[iPos].ErrorR[iStep] = 0;
                }
            }

            CreateTestBoard();
            UpdateGridViewBatch();

            foreach (DataGridViewColumn DGV_Col in DGV_ProcedureList.Columns)
            {
                DGV_Col.HeaderCell.Style.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Pixel);
            }

            int iStationID = GetStationID();
            LogMessage(MSG.DEBUG, "Station ID : " + "0x" + iStationID.ToString("X8") + " (" + iStationID + ")");

            textBoxWinInstallDate.Text = GetWindowsInstallationDateTime(String.Empty).ToString("yyyy-MM-dd HH:mm:ss");
            textBoxPCName.Text = System.Environment.MachineName;

            LogMessage(MSG.DEBUG, "EzyCal()");
        }

        private void Button_Start_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_Start_Click()");

            if (BgWorker.IsBusy)
            {
                LogMessage(MSG.INFO, "Program is busy");
                return;
            }

            StartProcedures();
        }

        private void StartProcedures()
        {
            int iStepStart = 0;

            LogMessage(MSG.DEBUG, "StartProcedures()");

            ClearMessageLogs();

            if (DGV_ProcedureRun.Rows.Count < 1)
            {
                LogMessage(MSG.WARNING, "No procedure defined");
                return;
            }

            if (iMeters == 0)
            {
                LogMessage(MSG.WARNING, "No meters defined");
                return;
            }

            // Find selected step
            for (int iRow = 0; iRow < DGV_ProcedureRun.Rows.Count; iRow++)
            {
                if (DGV_ProcedureRun.Rows[iRow].Selected && iRow > 0)
                {
                    iStepStart = iRow;
                }
            }

            progressBar.Maximum = DGV_ProcedureRun.Rows.Count;
            progressBar.Value = iStepStart;

            ButtonsEnabled(false);

            BgWorker.RunWorkerAsync(iStepStart);
        }

        // Main procedure. Don't update GUI controls in here
        void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = RunProcedure(e, (int)e.Argument);
        }

        void BgWorker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            UpdateTestBoard();

            if (e.Cancelled)
            {
                LogMessage(MSG.INFO, "Procedure cancelled by user");
            }
            else if (e.Error != null)
            {
                LogMessage(MSG.ERROR, "Procedure FAILED! see supervisor");
                LogMessage(MSG.ERROR, e.Error.Message);
            }
            else if ((int)e.Result == -1)
            {
                LogMessage(MSG.ERROR, "Procedure FAILED! see supervisor");
            }
            else
            {
                LogMessage(MSG.INFO, "Procedure completed successfully");
                DGV_ProcedureRun.CurrentCell = DGV_ProcedureRun.Rows[0].Cells[0];
                DGV_ProcedureRun.Rows[0].Selected = true;
            }

            ButtonsEnabled(true);
        }

        // Update GUI from worker thread
        void BgWorker_Progress(object sender, ProgressChangedEventArgs e)
        {
            int iStep = e.ProgressPercentage;

            if ((iStep + 1) <= progressBar.Maximum)
            {
                progressBar.Value = iStep + 1;
            }
        }

        private int RunProcedure(DoWorkEventArgs e, int iStepStart)
        {
            int iStep = 0;
            int iResultsRow = 0;

            LogMessage(MSG.DEBUG, "RunProcedure()");

            int iReturn = SleepDoEvents(e, true, 1000);

            if (iReturn != 0)
            {
                return iReturn;
            }

            LogMessage(MSG.INFO, "Current directory: " + Directory.GetCurrentDirectory());

            // Find results row to update
            for (int i = 0; i < iStepStart; i++)
            {
                if ((int)ConvertToDouble(DGV_ProcedureRun.Rows[i].Cells[20].Value.ToString()) > 1)
                {
                    iResultsRow++;
                }
            }

            // Iterate through procedure test steps
            for (iStep = iStepStart; iStep < DGV_ProcedureRun.Rows.Count; iStep++)
            {
                UpdateGuiProcedureStep(iStep);

                LogMessage(MSG.INFO, "==< Step " + (iStep + 1) + " >== ( " + DGV_ProcedureRun.Rows[iStep].Cells[2].Value.ToString() + " )");

                GetStepParams(iStep);

                // Set voltage, run A cmds, import Imp.res
                iReturn = Incus(e, iStep);

                if (iReturn != 0)
                {
                    PowerDown(e, iStep);
                    return iReturn;
                }

                // Set current & power factor, run error tests & B cmds
                iReturn = Maleus(e, iStep, iResultsRow);

                if (iReturn != 0)
                {
                    PowerDown(e, iStep);
                    return iReturn;
                }

                // Turn off current, run C cmds
                iReturn = Stapes(e, iStep, iResultsRow);

                if (iReturn != 0)
                {
                    PowerDown(e, iStep);
                    return iReturn;
                }

                if (StepValues.Storing > 1)
                {
                    iResultsRow++;
                }

                BgWorker.ReportProgress(iStep);

                if (BgWorker.CancellationPending)
                {
                    PowerDown(e, iStep);
                    e.Cancel = true;
                    return 1;
                }

                if (ToolStripMenuItem_SingleStep.Checked)
                {
                    bClickStep = false;

                    while (bClickStep == false)
                    {
                        Thread.Sleep(100);

                        if (BgWorker.CancellationPending)
                        {
                            PowerDown(e, iStep);
                            e.Cancel = true;
                            return 1;
                        }
                    }
                }
            }

            return 0;
        }

        // Turn off amplifier
        private void PowerDown(DoWorkEventArgs e, int iStep)
        {
            double[] dI = new double[3];
            double[] dP = new double[3];
            double[] dV = new double[3];

            SetCurrent(e, false, iStep, true, dI, dP);
            SetVoltage(e, false, iStep, 0, dV);
        }

        private void GetStepParams(int iStep)
        {
            LogMessage(MSG.DEBUG, "GetStepParams(" + (iStep + 1) + ")");

            //StepValues.ProcedureID = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[0].Value.ToString());
            StepValues.PStepNo = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[1].Value.ToString());
            StepValues.Name = DGV_ProcedureRun.Rows[iStep].Cells[2].Value.ToString();
            StepValues.UA = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[3].Value.ToString());
            StepValues.UB = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[4].Value.ToString());
            StepValues.UC = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[5].Value.ToString());
            StepValues.IA = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[6].Value.ToString());
            StepValues.IB = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[7].Value.ToString());
            StepValues.IC = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[8].Value.ToString());
            StepValues.IsImax = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[9].Value.ToString());
            StepValues.PHI = (int)ConvertToDouble(Regex.Replace(DGV_ProcedureRun.Rows[iStep].Cells[10].Value.ToString(), "[^0-9.]", ""));
            //StepValues.FREQ = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[11].Value.ToString());
            //StepValues.Waveform = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[12].Value.ToString());
            //StepValues.PhaseSeq = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[13].Value.ToString());
            StepValues.TestTypeID = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[14].Value.ToString());
            StepValues.Measurement = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[15].Value.ToString());
            StepValues.NumPulses = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[16].Value.ToString());
            StepValues.ULIMIT = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[17].Value.ToString());
            StepValues.LLIMIT = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[18].Value.ToString());
            StepValues.ChannelNo = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[19].Value.ToString());
            StepValues.Storing = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[20].Value.ToString());
            StepValues.FileIE = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[21].Value.ToString());
            //StepValues.Duration = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[22].Value.ToString());
            StepValues.Timeout = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[23].Value.ToString());
            StepValues.Finally = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[24].Value.ToString());
            StepValues.ACMDS = DGV_ProcedureRun.Rows[iStep].Cells[25].Value.ToString();
            StepValues.BCMDS = DGV_ProcedureRun.Rows[iStep].Cells[26].Value.ToString();
            StepValues.CCMDS = DGV_ProcedureRun.Rows[iStep].Cells[27].Value.ToString();
            StepValues.WithAmp = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[28].Value.ToString());
            //StepValues.MinimumTime = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[29].Value.ToString());
            StepValues.BaseUb = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[30].Value.ToString());
            StepValues.BaseIb = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[31].Value.ToString());
            StepValues.BaseImax = (int)ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[32].Value.ToString());

            UpdateGuiStepValues(iStep);

            string strLenUB = DGV_ProcedureRun.Rows[iStep].Cells[4].Value.ToString();
            string strLenUC = DGV_ProcedureRun.Rows[iStep].Cells[5].Value.ToString();
            string strLenIB = DGV_ProcedureRun.Rows[iStep].Cells[7].Value.ToString();
            string strLenIC = DGV_ProcedureRun.Rows[iStep].Cells[8].Value.ToString();

            // B & C set to A if not defined
            if (strLenUB.Length == 0) { StepValues.UB = StepValues.UA; }
            if (strLenUC.Length == 0) { StepValues.UC = StepValues.UA; }
            if (strLenIB.Length == 0) { StepValues.IB = StepValues.IA; }
            if (strLenIC.Length == 0) { StepValues.IC = StepValues.IA; }
        }

        private int Incus(DoWorkEventArgs e, int iStep)
        {
            double[] dV = new double[3];

            LogMessage(MSG.DEBUG, "Incus()");

            dV[0] = (StepValues.UA / 100.0) * StepValues.BaseUb;
            dV[1] = (StepValues.UB / 100.0) * StepValues.BaseUb;
            dV[2] = (StepValues.UC / 100.0) * StepValues.BaseUb;

            int iResult = SetVoltage(e, true, iStep, 1, dV);

            if (iResult != 0)
            {
                return iResult;
            }

            // For dial test (TODO)
            if (StepValues.TestTypeID == 5)
            {
                LogMessage(MSG.ERROR, "Dial test NOT implemented");
                return -1;
            }

            iResult = RunCMDs(e, ParseCmds(StepValues.ACMDS));

            if (iResult != 0)
            {
                return iResult;
            }

            if ((StepValues.FileIE & 0x1) == 0x1)
            {
                ReadImportFile(iStep, 0);
            }

            return 0;
        }

        private int Maleus(DoWorkEventArgs e, int iStep, int iResultsRow)
        {
            LogMessage(MSG.DEBUG, "Maleus()");

            double[] dI = new double[3];
            double[] dP = new double[3];

            dI[0] = (StepValues.IA / 100.0) * (StepValues.IsImax == 2 ? StepValues.BaseImax : StepValues.BaseIb);
            dI[1] = (StepValues.IB / 100.0) * (StepValues.IsImax == 2 ? StepValues.BaseImax : StepValues.BaseIb);
            dI[2] = (StepValues.IC / 100.0) * (StepValues.IsImax == 2 ? StepValues.BaseImax : StepValues.BaseIb);

            dP[0] = dP[1] = dP[2] = StepValues.PHI;

            int iResult = SetCurrent(e, true, iStep, false, dI, dP);

            if (iResult != 0)
            {
                return iResult;
            }

            iResult = RunCMDs(e, ParseCmds(StepValues.BCMDS));

            if (iResult != 0)
            {
                return iResult;
            }

            // Start & creep tests
            if (StepValues.TestTypeID == 2 || StepValues.TestTypeID == 3)
            {
                iResult = DoStartTest(e, iStep, iResultsRow);

                if (iResult != 0)
                {
                    return iResult;
                }
            }

            // Error test
            if (StepValues.TestTypeID == 4)
            {
                iResult = DoErrorTest(e, iStep, iResultsRow);

                if (iResult != 0)
                {
                    return iResult;
                }
            }

            // Other tests
            if (StepValues.TestTypeID != 2 && StepValues.TestTypeID != 3 && StepValues.TestTypeID != 4)
            {
                iResult = SleepDoEvents(e, true, StepValues.Timeout * 1000);

                if (iResult != 0)
                {
                    return iResult;
                }
            }

            // Export file
            if ((StepValues.FileIE & 0x2) == 0x2)
            {
                if (CreateExportFile(iStep))
                {
                    return -1;
                }
            }

            // Drop current in normal cases
            if (StepValues.WithAmp != 1)
            {
                if (StepValues.Finally != 2)
                {
                    iResult = SetCurrent(e, true, iStep, true, dI, dP);

                    if (iResult != 0)
                    {
                        return iResult;
                    }
                }
            }

            return 0;
        }

        private int Stapes(DoWorkEventArgs e, int iStep, int iResultsRow)
        {
            LogMessage(MSG.DEBUG, "Stapes()");

            // Import2 File
            if ((StepValues.FileIE & 0x4) == 0x4)
            {
                ReadImportFile(iStep, 1);
                //ReadMeterClock(iStep);
            }

            if (StepValues.Storing > 1)
            {
                UpdateResults(iStep, iResultsRow);
            }

            int iResult = RunCMDs(e, ParseCmds(StepValues.CCMDS));

            if (iResult != 0)
            {
                return iResult;
            }

            // For dial test? (TODO)
            if (StepValues.TestTypeID == 5)
            {
                LogMessage(MSG.ERROR, "Dial test NOT implemented");
                return -1;
            }

            // Wait if required
            if (StepValues.Finally == 1 || StepValues.Finally == 2)
            {
                MessageBox.Show("Press OK button to continue testing", "Important Note", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }

            // Drop current
            if (StepValues.WithAmp == 1)
            {
                double[] dI = new double[3];
                double[] dP = new double[3];

                iResult = SetCurrent(e, true, iStep, true, dI, dP);

                if (iResult != 0)
                {
                    return iResult;
                }
            }

            return 0;
        }

        public int SetVoltage(DoWorkEventArgs e, bool bCancel, int iStep, int iMode, double[] dV)
        {
            bool bResult = true;
            int iResult = 0;
            int iRetries;

            LogMessage(MSG.DEBUG, "SetVoltage()");

            if (iMode == 0)
            {
                dV[0] = 0;
                dV[1] = 0;
                dV[2] = 0;
            }

            /*
            if ((dV[0] == Electricals.uA) && (dV[1] == Electricals.uB) && (dV[2] == Electricals.uC) && (iMode != 0))
            {
                LogMessage(MSG.DEBUG, "Voltage already correct");
                return 0;
            }
            */

            // Simulation mode
            Desired.uA = dV[0];
            Desired.uB = dV[1];
            Desired.uC = dV[2];

            Counters.iSetVoltage++;

            UpdateGuiDesiredVoltage(iStep);

            // Set RefMtr Volt range
            if (iMode != 0)
            {
                bResult = OpenBox(BOX_REFMTR232);

                if (bResult)
                {
                    double U = 0.0;
                    int iConfig = 10 * (Electricals.lineType);
                    iConfig += 2;

                    if (Bench.numPhase == 1)
                    {
                        U = dV[0];
                    }
                    else
                    {
                        U = Math.Max(Math.Max(dV[0], dV[1]), dV[2]);
                    }

                    if (U > Bench.valMaxMin.uMax)
                    {
                        U = Bench.valMaxMin.uMax;
                    }

                    bool[] b = new bool[] { false, false, false };

                    iRetries = 2;

                    while (iRetries-- > 0)
                    {
                        if (!b[0] && 0 == SetConfig(iConfig))
                        {
                            b[0] = true;
                        }

                        // 1=Active, 2=Reactive
                        if (!b[1] && 0 == SetPQS(1))
                        {
                            b[1] = true;
                        }

                        if (!b[2] && 0 == SetURange(U))
                        {
                            b[2] = true;
                        }

                        Thread.Sleep(200);

                        if (b[0] && b[1] && b[2])
                        {
                            break;
                        }
                    }

                    if (!(b[0] && b[1] && b[2]))
                    {
                        bResult = false;
                    }
                }
            }

            if (bResult)
            {
                bResult = OpenBox(BOX_UIPFSOURCE);
            }

            if (bResult)
            {
                double[] U = new double[3];
                bool[] b = new bool[] { false, false, false };

                if (Bench.numPhase == 1)
                {
                    b[1] = b[2] = true;
                }

                // Safety
                U[0] = dV[0] > Bench.valMaxMin.uMax ? Bench.valMaxMin.uMax : dV[0];
                U[1] = dV[1] > Bench.valMaxMin.uMax ? Bench.valMaxMin.uMax : dV[1];
                U[2] = dV[2] > Bench.valMaxMin.uMax ? Bench.valMaxMin.uMax : dV[2];

                iRetries = 5;

                while (iRetries-- > 0)
                {
                    // Start profile
                    int iNumStops = 6;
                    double[] dPercentU = new double[] { 0.25, 0.5, 0.6, 0.7, 0.8, 1.0 };
                    int[] iDelayU = new int[] { 10, 10, 20, 20, 20, 0 };

                    for (int iStops = 0; iStops < iNumStops; iStops++)
                    {
                        // iMode: 0 = drop U, 1 = normal start, 2 = step start
                        if (iMode == 0 || iMode == 1)
                        {
                            iStops = iNumStops - 1; // For drop voltage or normal start
                        }

                        for (int p = 0; p < Bench.numPhase; p++)
                        {
                            if (b[p]) continue;

                            LogMessage(MSG.INFO, "VoltageOut(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + (220 + p) + "," + U[p] * dPercentU[iStops] + ",false)");

                            if (BgWorker.CancellationPending && bCancel)
                            {
                                e.Cancel = true;
                                return 1;
                            }

                            if (Settings.bRunCtrComm)
                            {
                                b[p] = CheckResult(_CtrComm.VoltageOut(Bench.ctrlSio, Bench.ctrlSioFmt, (short)(220 + p), U[p] * dPercentU[iStops], false));
                            }
                            else
                            {
                                b[p] = true;
                            }

                            Thread.Sleep(300);
                        }

                        if (iStops < (iNumStops - 1))
                        {
                            b[0] = b[1] = b[2] = false;

                            if (Bench.numPhase == 1)
                            {
                                b[1] = b[2] = true;
                            }

                            iResult = SleepDoEvents(e, bCancel, 1000 * iDelayU[iStops]);

                            if (iResult != 0)
                            {
                                return iResult;
                            }
                        }
                    }

                    if (b[0] && b[1] && b[2])
                    {
                        break;
                    }
                }

                if (!(b[0] && b[1] && b[2]))
                {
                    bResult = false;
                }
            }

            if (!bResult)
            {
                return -1;
            }

            iResult = CheckTargetU(e, bCancel, iStep, dV);

            if (iResult != 0)
            {
                Counters.iSetVoltageErrors++;
                return iResult;
            }

            Electricals.uA = dV[0];
            Electricals.uB = dV[1];
            Electricals.uC = dV[2];

            return 0;
        }

        public int CheckTargetU(DoWorkEventArgs e, bool bCancel, int iStep, double[] dV)
        {
            int iResult = 0;

            LogMessage(MSG.DEBUG, "CheckTargetU(" + dV[0] + ", " + dV[1] + ", " + dV[2] + ")");

            for (int c = 0; c < 10; c++)
            {
                iResult = SleepDoEvents(e, bCancel, 1000);

                if (iResult != 0)
                {
                    return iResult;
                }

                iResult = GetActuals(e, bCancel, iStep);

                if (iResult != 0)
                {
                    return iResult;
                }

                if (dV[0] == 0.0)
                {
                    Electricals.bA = (Math.Abs(Actuals.uA) < 1.0);
                }
                else
                {
                    Electricals.bA = (Math.Abs((Actuals.uA / dV[0]) - 1) <= 0.1);
                }

                if (dV[1] == 0.0)
                {
                    Electricals.bB = (Math.Abs(Actuals.uB) < 1.0);
                }
                else
                {
                    Electricals.bB = (Math.Abs((Actuals.uB / dV[1]) - 1) <= 0.1);
                }

                if (dV[2] == 0.0)
                {
                    Electricals.bC = (Math.Abs(Actuals.uC) < 1.0);
                }
                else
                {
                    Electricals.bC = (Math.Abs((Actuals.uC / dV[2]) - 1) <= 0.1);
                }

                UpdateGuiActualVoltageColor(iStep);

                if (Bench.numPhase == 1)
                {
                    if (Electricals.bA)
                    {
                        return 0;
                    }
                }
                else
                {
                    if (Electricals.bA && Electricals.bB && Electricals.bC)
                    {
                        return 0;
                    }
                }
            }

            return -1;
        }

        public int SetCurrent(DoWorkEventArgs e, bool bCancel, int iStep, bool bOff, double[] dI, double[] dP)
        {
            bool bResult = true;
            int iResult = 0;
            int iRetries;

            LogMessage(MSG.DEBUG, "SetCurrent(" + bOff + ")");

            if (bOff)
            {
                dI[0] = 0;
                dI[1] = 0;
                dI[2] = 0;
            }

            /*
            if ((dI[0] == Electricals.iA) && (dI[1] == Electricals.iB) && (dI[2] == Electricals.iC) && !bOff)
            {
                LogMessage(MSG.DEBUG, "Current already correct");
                return 0;
            }
            */

            Desired.iA = dI[0];
            Desired.iB = dI[1];
            Desired.iC = dI[2];
            Desired.phiA = dP[0];
            Desired.phiB = dP[1];
            Desired.phiC = dP[2];

            Counters.iSetCurrent++;

            UpdateGuiDesiredCurrent(iStep);

            // Set RefMtr current range
            if (!bOff)
            {
                bResult = OpenBox(BOX_REFMTR232);

                if (bResult)
                {
                    double I = 0.0;

                    if (Bench.numPhase == 1)
                    {
                        I = dI[0];
                    }
                    else
                    {
                        I = Math.Max(Math.Max(dI[0], dI[1]), dI[2]);
                    }

                    if (I > Bench.valMaxMin.iMax)
                    {
                        I = Bench.valMaxMin.iMax;
                    }

                    bool r = false;

                    iRetries = 4;

                    while (iRetries-- > 0)
                    {
                        if (!r && 0 == SetIRange(I))
                        {
                            r = true;
                        }

                        Thread.Sleep(200);

                        if (r)
                        {
                            break;
                        }
                    }

                    if (!r)
                    {
                        bResult = false;
                    }
                }
            }

            if (bResult)
            {
                bResult = OpenBox(BOX_UIPFSOURCE);
            }

            // Set I source
            if (bResult && !bOff)
            {
                bool[] r = new bool[] { false, false, false };

                if (Bench.numPhase == 1)
                {
                    r[1] = r[2] = true;
                }

                iRetries = 5;

                while (iRetries-- > 0)
                {
                    for (int p = 0; p < Bench.numPhase; p++)
                    {
                        if (r[p]) continue;

                        LogMessage(MSG.INFO, "CosOut(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + (220 + p) + "," + dP[p] + ")");

                        if (BgWorker.CancellationPending && bCancel)
                        {
                            e.Cancel = true;
                            return 1;
                        }

                        if (Settings.bRunCtrComm)
                        {
                            r[p] = CheckResult(_CtrComm.CosOut(Bench.ctrlSio, Bench.ctrlSioFmt, (short)(220 + p), dP[p]));
                        }
                        else
                        {
                            r[p] = true;
                        }

                        Thread.Sleep(200);
                    }

                    if (r[0] && r[1] && r[2])
                    {
                        break;
                    }
                }

                if (!(r[0] && r[1] && r[2]))
                {
                    bResult = false;
                }
            }

            if (bResult)
            {
                double[] I = new double[3];
                bool[] r = new bool[] { false, false, false };

                if (Bench.numPhase == 1)
                {
                    r[1] = r[2] = true;
                }

                // Safety
                I[0] = Math.Min(dI[0], Bench.valMaxMin.iMax);
                I[1] = Math.Min(dI[1], Bench.valMaxMin.iMax);
                I[2] = Math.Min(dI[2], Bench.valMaxMin.iMax);

                bool bBigAmp = false;

                if (I[0] > 105.0 || I[1] > 105.0 || I[2] > 105.0)
                {
                    bBigAmp = true;
                }

                iRetries = 5;

                while (iRetries-- > 0)
                {
                    // Start profile
                    int iNumStops = 4;
                    double[] dPercentI = new double[] { 0.5, 0.7, 0.88, 1.0 };
                    int[] iDelayU = new int[] { 5, 5, 8, 0 };

                    for (int iStops = 0; iStops < iNumStops; iStops++)
                    {
                        // For normal but low ampere
                        if (!bBigAmp)
                        {
                            iStops = iNumStops - 1;
                        }

                        for (int p = 0; p < Bench.numPhase; p++)
                        {
                            if (r[p]) continue;

                            LogMessage(MSG.INFO, "CurrentOut(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + (220 + p) + "," + I[p] * (dPercentI[iStops]) + ",false)");

                            if (BgWorker.CancellationPending && bCancel)
                            {
                                e.Cancel = true;
                                return 1;
                            }

                            if (Settings.bRunCtrComm)
                            {
                                r[p] = CheckResult(_CtrComm.CurrentOut(Bench.ctrlSio, Bench.ctrlSioFmt, (short)(220 + p), I[p] * (dPercentI[iStops]), false));
                            }
                            else
                            {
                                r[p] = true;
                            }

                            Thread.Sleep(200);
                        }

                        if (iStops < (iNumStops - 1))
                        {
                            r[0] = r[1] = r[2] = false;

                            if (Bench.numPhase == 1)
                            {
                                r[1] = r[2] = true;
                            }

                            iResult = SleepDoEvents(e, bCancel, 1000 * iDelayU[iStops]);

                            if (iResult != 0)
                            {
                                return iResult;
                            }
                        }
                    }

                    if (r[0] && r[1] && r[2])
                    {
                        break;
                    }
                }

                if (!(r[0] && r[1] && r[2]))
                {
                    bResult = false;
                }
            }

            // For accurate measurement
            if (bResult && bOff)
            {
                Thread.Sleep(800); // Must

                bResult = OpenBox(BOX_REFMTR232);

                if (bResult)
                {
                    bool r = false;

                    iRetries = 4;

                    while (iRetries-- > 0)
                    {
                        if (!r && 0 == SetIRange(0.0001))
                        {
                            r = true;
                        }

                        Thread.Sleep(200);

                        if (r)
                        {
                            break;
                        }
                    }

                    if (!r)
                    {
                        bResult = false;
                    }
                }
            }

            if (!bResult)
            {
                return -1;
            }

            iResult = CheckTargetI(e, bCancel, iStep, dI);

            if (iResult != 0)
            {
                Counters.iSetCurrentErrors++;
                return iResult;
            }

            Electricals.iA = dI[0];
            Electricals.iB = dI[1];
            Electricals.iC = dI[2];

            return 0;
        }

        public int CheckTargetI(DoWorkEventArgs e, bool bCancel, int iStep, double[] dI)
        {
            int iResult = 0;

            LogMessage(MSG.DEBUG, "CheckTargetI(" + dI[0] + ", " + dI[1] + ", " + dI[2] + ")");

            for (int c = 0; c < 10; c++)
            {
                iResult = SleepDoEvents(e, bCancel, 1000);

                if (iResult != 0)
                {
                    return iResult;
                }

                iResult = GetActuals(e, bCancel, iStep);

                if (iResult != 0)
                {
                    return iResult;
                }

                if (dI[0] == 0.0)
                {
                    Electricals.bA = (Math.Abs(Actuals.iA) < 0.25);
                }
                else
                {
                    Electricals.bA = (Math.Abs((Actuals.iA / dI[0]) - 1) <= 0.1);
                }

                if (dI[1] == 0.0)
                {
                    Electricals.bB = (Math.Abs(Actuals.iB) < 0.25);
                }
                else
                {
                    Electricals.bB = (Math.Abs((Actuals.iB / dI[1]) - 1) <= 0.1);
                }

                if (dI[2] == 0.0)
                {
                    Electricals.bC = (Math.Abs(Actuals.iC) < 0.25);
                }
                else
                {
                    Electricals.bC = (Math.Abs((Actuals.iC / dI[2]) - 1) <= 0.1);
                }

                UpdateGuiActualCurrentColor(iStep);

                if (Bench.numPhase == 1)
                {
                    if (Electricals.bA)
                    {
                        return 0;
                    }
                }
                else
                {
                    if (Electricals.bA && Electricals.bB && Electricals.bC)
                    {
                        return 0;
                    }
                }
            }

            return -1;
        }

        private int RunCMDs(DoWorkEventArgs e, string strCmds)
        {
            LogMessage(MSG.DEBUG, "RunCMDs()");

            if (string.IsNullOrEmpty(strCmds))
            {
                return 0;
            }

            string[] strAryCmds = strCmds.Split('|');

            for (int n = 0; n < strAryCmds.Length; n++)
            {
                bool bShellCmd = false;

                if (strAryCmds[n].Length > 0)
                {
                    string strCmd = strAryCmds[n];

                    Process p = new Process();
                    p.StartInfo.UseShellExecute = true;
                    p.StartInfo.WorkingDirectory = Settings.strRootPath;

                    string[] strAryCmd = strCmd.Split(' ');
                    string strArgList = null;

                    if (strAryCmd.Length > 1)
                    {
                        for (int a = 1; a < strAryCmd.Length; a++)
                        {
                            strArgList += strAryCmd[a] + " ";
                        }

                        strCmd = strAryCmd[0];
                        p.StartInfo.Arguments = strArgList;
                    }

                    // Special case for OS shell commands
                    if (String.Equals(strAryCmd[0].Trim(), "del", StringComparison.OrdinalIgnoreCase))
                    {
                        bShellCmd = true;
                    }

                    // Procedure command
                    if (String.Equals(strAryCmd[0].Trim(), "manual", StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show(strArgList);
                        continue;
                    }

                    // Procedure command
                    if (String.Equals(strAryCmd[0].Trim(), "wait", StringComparison.OrdinalIgnoreCase))
                    {
                        int iMs = 0;
                        TimeSpan result;

                        if (TimeSpan.TryParse(strArgList, out result))
                        {
                            iMs = (int)result.TotalSeconds * 1000;
                        }

                        int iResult = SleepDoEvents(e, true, iMs);

                        if (iResult != 0)
                        {
                            return iResult;
                        }

                        continue;
                    }

                    // Replace hard coded paths with strRootPath
                    var pattern = @"C:\\Programme\\EMH\\CamCal\\";
                    var rgx = new Regex(pattern, RegexOptions.IgnoreCase);
                    var file = @"\" + rgx.Replace(strCmd, "", 1);

                    // Determine if file is .exe or .bat for Process()
                    string strPathFile = Settings.strRootPath + file;
                    string strPathFileBat = Settings.strRootPath + file + ".bat";
                    string strPathFileExe = Settings.strRootPath + file + ".exe";

                    if (!bShellCmd)
                    {
                        if (File.Exists(strPathFile) == false)
                        {
                            if (File.Exists(strPathFileBat) == false)
                            {
                                if (File.Exists(strPathFileExe) == false)
                                {
                                    if (Settings.bFilesExist)
                                    {
                                        LogMessage(MSG.ERROR, "File NOT found: " + strPathFile);
                                        return -1;
                                    }
                                }
                                else
                                {
                                    strPathFile = strPathFileExe;
                                }
                            }
                            else
                            {
                                strPathFile = strPathFileBat;
                            }
                        }
                    }

                    p.StartInfo.WindowStyle = ProcessWindowStyle.Normal;

                    if (ToolStripMenuItem_Min.Checked)
                    {
                        p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                    }

                    if (bShellCmd)
                    {
                        p.StartInfo.FileName = "cmd.exe";
                        p.StartInfo.Arguments = "/C" + " " + strCmd + " " + strArgList;
                    }
                    else
                    {
                        p.StartInfo.FileName = strPathFile;
                    }

                    LogMessage(MSG.INFO, p.StartInfo.FileName + " " + p.StartInfo.Arguments);

                    if (BgWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        return 1;
                    }

                    if (Settings.bRunCmds)
                    {
                        try
                        {
                            p.Start();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, ex.ToString());
                            return -1;
                        }

                        while (!p.HasExited)
                        {
                            Thread.Sleep(100);

                            if (BgWorker.CancellationPending)
                            {
                                e.Cancel = true;
                                return 1;
                            }
                        }

                        if (Settings.bProcessExitCode)
                        {
                            if (p.ExitCode != 0)
                            {
                                return -1;
                            }
                        }
                    }
                }
            }

            return 0;
        }

        private string ParseCmds(string strCmds)
        {
            string strParsedCmds = null;

            LogMessage(MSG.DEBUG, "ParseCmds()");

            if (string.IsNullOrEmpty(strCmds))
            {
                return strParsedCmds;
            }

            string[] strAryCmds = strCmds.Split('|');

            for (int n = 0; n < strAryCmds.Length; n++)
            {
                if (strAryCmds[n].Length > 0)
                {
                    string strCmd = strAryCmds[n];
                    string[] strAryCmd = strCmd.Split(' ');
                    string strArgList = null;

                    if (strAryCmd.Length > 1)
                    {
                        for (int a = 1; a < strAryCmd.Length; a++)
                        {
                            strArgList += " ";

                            if (Regex.IsMatch(strAryCmd[a], @"\[V"))
                            {
                                switch (strAryCmd[a])
                                {
                                    case "[V0]": strArgList += ((StepValues.UA / 100.0) * StepValues.BaseUb).ToString(); break; // applied A phase voltage
                                    case "[V1]": strArgList += ((StepValues.UB / 100.0) * StepValues.BaseUb).ToString(); break; // applied B phase voltage
                                    case "[V2]": strArgList += ((StepValues.UC / 100.0) * StepValues.BaseUb).ToString(); break; // applied C phase voltage
                                    case "[V3]": strArgList += Actuals.uA.ToString(); break; // instant A phase voltage
                                    case "[V4]": strArgList += Actuals.uB.ToString(); break; // instant B phase voltage
                                    case "[V5]": strArgList += Actuals.uC.ToString(); break; // instant C phase voltage
                                    case "[V6]": strArgList += StepValues.BaseUb.ToString(); break; // rated A phase voltage
                                    case "[V7]": strArgList += StepValues.BaseUb.ToString(); break; // rated B phase voltage
                                    case "[V8]": strArgList += StepValues.BaseUb.ToString(); break; // rated C phase voltage
                                    default: LogMessage(MSG.ERROR, "Command argument NOT found: " + strAryCmd[a]); break;
                                }
                            }
                            else if (Regex.IsMatch(strAryCmd[a], @"\[I"))
                            {
                                switch (strAryCmd[a])
                                {
                                    case "[I0]": strArgList += ((StepValues.IA / 100.0) * StepValues.BaseIb).ToString(); break; // applied A phase current
                                    case "[I1]": strArgList += ((StepValues.IB / 100.0) * StepValues.BaseIb).ToString(); break; // applied B phase current
                                    case "[I2]": strArgList += ((StepValues.IC / 100.0) * StepValues.BaseIb).ToString(); break; // applied C phase current
                                    case "[I3]": strArgList += Actuals.iA.ToString(); break; // instant A phase current
                                    case "[I4]": strArgList += Actuals.iB.ToString(); break; // instant B phase current
                                    case "[I5]": strArgList += Actuals.iC.ToString(); break; // instant C phase current
                                    case "[I6]": strArgList += StepValues.BaseIb.ToString(); break; // rated A phase current
                                    case "[I7]": strArgList += StepValues.BaseIb.ToString(); break; // rated B phase current
                                    case "[I8]": strArgList += StepValues.BaseIb.ToString(); break; // rated C phase current
                                    default: LogMessage(MSG.ERROR, "Command argument NOT found: " + strAryCmd[a]); break;
                                }
                            }
                            else
                            {
                                strArgList += strAryCmd[a];
                            }
                        }
                    }

                    strParsedCmds += strAryCmd[0] + strArgList + "|";
                }
            }

            strParsedCmds = strParsedCmds.TrimEnd('|');

            return strParsedCmds;
        }

        private int DoStartTest(DoWorkEventArgs e, int iStep, int iResultsRow)
        {
            int iResult;

            LogMessage(MSG.DEBUG, "DoStartTest()");

            iResult = OnMark(e);

            if (iResult != 0)
            {
                return iResult;
            }

            iResult = SleepDoEvents(e, true, 15000);

            if (iResult != 0)
            {
                return iResult;
            }

            iResult = StartTest(e);

            if (iResult != 0)
            {
                return iResult;
            }

            // No-load test gives pass/fail after timeout period, add extra time.
            long lTotalTime = StepValues.Timeout * 1000 + 20000;

            lTotalTime = lTotalTime / Settings.iSleepDivide;

            Stopwatch Stopwatch = Stopwatch.StartNew();

            while (lTotalTime > Stopwatch.ElapsedMilliseconds)
            {
                //GetActuals(); Why?

                iResult = SleepDoEvents(e, true, 1000);

                if (iResult != 0)
                {
                    return iResult;
                }

                iResult = ReadErrors(e, iStep);

                if (iResult != 0)
                {
                    return iResult;
                }

                UpdateDisplayErrors(iStep);

                if (StepValues.Storing > 1)
                {
                    UpdateResults(iStep, iResultsRow);
                }

                iResult = SleepDoEvents(e, true, 1000);

                if (iResult != 0)
                {
                    return iResult;
                }

                bool bExit = true;

                // Exit once we have ALL results
                for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                {
                    if (TestBoard[iPos].Active)
                    {
                        double dVal = TestBoard[iPos].ErrorV[iStep];

                        if ((dVal != iPASS) && (dVal != iFAIL))
                        {
                            bExit = false;
                        }
                    }
                }

                if (bExit)
                {
                    return 0;
                }
            }

            return 0;
        }

        private int DoErrorTest(DoWorkEventArgs e, int iStep, int iResultsRow)
        {
            LogMessage(MSG.DEBUG, "DoErrorTest()");

            int iResult = SetErrorCntr(e);

            if (iResult != 0)
            {
                return iResult;
            }

            int iTotalTime = StepValues.Timeout * 1000;

            iTotalTime = iTotalTime / Settings.iSleepDivide;

            // Approx. min reads of error counter as GENY API calls can take 1 minute.
            int iMinReadErrors = iTotalTime / (500 + (iMeters * 100));

            Stopwatch Stopwatch = Stopwatch.StartNew();

            // Read errors for timeout period and at least iMinReadErrors times
            while ((iTotalTime > Stopwatch.ElapsedMilliseconds) || (iMinReadErrors > 0))
            {
                iMinReadErrors--;

                //GetActuals(); Why?

                //Thread.Sleep(200); Why?

                iResult = ReadErrors(e, iStep);

                if (iResult != 0)
                {
                    return iResult;
                }

                UpdateDisplayErrors(iStep);

                if (StepValues.Storing > 1)
                {
                    UpdateResults(iStep, iResultsRow);
                }

                //Thread.Sleep(200); Why?
            }

            return 0;
        }

        private int OnMark(DoWorkEventArgs e)
        {
            bool bResult = true;
            int iRetries;
            int iSuccess = 0;

            LogMessage(MSG.DEBUG, "OnMark()");

            bResult = OpenBox(BOX_ERRCOUNTER);

            if (bResult)
            {
                iRetries = 3;

                while (iRetries-- > 0)
                {
                    LogMessage(MSG.INFO, "ErrCounterSeBiao(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + 199 + "," + StepValues.ChannelNo + ",1)");

                    if (BgWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        return 1;
                    }

                    Thread.Sleep(200);

                    if (Settings.bRunCtrComm)
                    {
                        if (CheckResult(_CtrComm.ErrCounterSeBiao(
                            Bench.ctrlSio,
                            Bench.ctrlSioFmt,
                            199, // Broadcast mode
                            (short)StepValues.ChannelNo,
                            (short)1)))
                        {
                            iSuccess++;
                        }
                    }
                    else
                    {
                        // Debug mode
                        iSuccess++;
                    }

                    if (iSuccess >= 1)
                    {
                        break;
                    }
                }
            }

            if (iSuccess < 1)
            {
                return -1;
            }

            return 0;
        }

        private int StartTest(DoWorkEventArgs e)
        {
            bool bResult = true;
            int iRetries;
            int iSuccess = 0;

            LogMessage(MSG.DEBUG, "StartTest()");

            bResult = OpenBox(BOX_ERRCOUNTER);

            if (bResult)
            {
                iRetries = 10;

                while (iRetries-- > 0)
                {
                    if (StepValues.TestTypeID == 2)
                    {
                        LogMessage(MSG.INFO, "ErrCounterQiDong(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + 199 + "," + StepValues.Timeout + "," + StepValues.NumPulses + "," + StepValues.ChannelNo + ",1)");
                    }
                    else
                    {
                        LogMessage(MSG.INFO, "ErrCounterQianDong(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + 199 + "," + StepValues.Timeout + "," + StepValues.NumPulses + "," + StepValues.ChannelNo + ",1)");
                    }

                    if (BgWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        return 1;
                    }

                    Thread.Sleep(200);

                    if (Settings.bRunCtrComm)
                    {
                        if (StepValues.TestTypeID == 2)
                        {
                            if (CheckResult(_CtrComm.ErrCounterQiDong(
                                Bench.ctrlSio,
                                Bench.ctrlSioFmt,
                                199, // Broadcast mode
                                (short)StepValues.Timeout,
                                (short)StepValues.NumPulses,
                                (short)StepValues.ChannelNo,
                                (short)1)))
                            {
                                iSuccess++;
                            }
                        }
                        else
                        {
                            if (CheckResult(_CtrComm.ErrCounterQianDong(
                                Bench.ctrlSio,
                                Bench.ctrlSioFmt,
                                199, // Broadcast mode
                                (short)StepValues.Timeout,
                                (short)StepValues.NumPulses,
                                (short)StepValues.ChannelNo,
                                (short)1)))
                            {
                                iSuccess++;
                            }
                        }
                    }
                    else
                    {
                        // Debug mode
                        iSuccess++;
                    }

                    if (iSuccess >= 1)
                    {
                        break;
                    }
                }
            }

            if (iSuccess < 1)
            {
                return -1;
            }

            return 0;
        }

        public int GetActuals(DoWorkEventArgs e, bool bCancel, int iStep)
        {
            int iRetries;
            bool bResult = true;

            LogMessage(MSG.DEBUG, "GetActuals()");

            Counters.iReadRefMeter++;

            Actuals.numPhase = Bench.numPhase;

            bResult = OpenBox(BOX_REFMTR232);

            if (bResult)
            {
                iRetries = 2;

                while (iRetries-- > 0)
                {
                    if (Settings.bRunCtrComm)
                    {
                        if (BgWorker.CancellationPending && bCancel)
                        {
                            e.Cancel = true;
                            return 1;
                        }

                        if (Bench.numPhase == 1)
                        {
                            Actuals.uB = 0;
                            Actuals.uC = 0;
                            Actuals.iB = 0;
                            Actuals.iC = 0;
                            Actuals.phiB = 0;
                            Actuals.phiC = 0;

                            bResult = ReadActuals1();
                        }
                        else
                        {
                            bResult = ReadActuals3();
                        }
                    }
                    else
                    {
                        CreateActuals();
                    }

                    Thread.Sleep(200);

                    if (bResult)
                    {
                        break;
                    }

					Counters.iReadRefMeterErrors++;
                }
            }

            if (bResult)
            {
                UpdateGuiActuals(iStep);
            }
            else
            {
                return -1;
            }

            return 0;
        }

        private int SetErrorCntr(DoWorkEventArgs e)
        {
            bool bResult = true;
            int iRetries;
            int iSuccess = 0;

            LogMessage(MSG.DEBUG, "SetErrorCntr()");

            Counters.iSetErrorCounter++;

            bResult = OpenBox(BOX_REFMTR232);

            // Set RefMtr measurement element
            if (bResult)
            {
                iRetries = 2;

                while (iRetries-- > 0)
                {
                    bResult = SetPQS(StepValues.Measurement) == 0 ? true : false;

                    Thread.Sleep(200);

                    if (bResult)
                    {
                        break;
                    }
                }
            }

            if (bResult)
            {
                bResult = OpenBox(BOX_ERRCOUNTER);
            }

            if (bResult)
            {
                iRetries = 10;

                while (iRetries-- > 0)
                {
                    LogMessage(MSG.INFO, "ErrCounterTest(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + 199 + "," + StepValues.NumPulses + "," + (dCurConst / ErrorCounter.mtrConst) * StepValues.NumPulses + "," + StepValues.ChannelNo + ",1," + StepValues.ULIMIT + "," + StepValues.LLIMIT + ")");

                    if (BgWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        return 1;
                    }

                    Thread.Sleep(500); // GENY recommend

                    if (Settings.bRunCtrComm)
                    {
                        if (CheckResult(_CtrComm.ErrCounterTest(
                            Bench.ctrlSio,
                            Bench.ctrlSioFmt,
                            199, // Broadcast mode
                            (short)StepValues.NumPulses,
                            (dCurConst / ErrorCounter.mtrConst) * StepValues.NumPulses,
                            (short)StepValues.ChannelNo,
                            (short)1,
                            StepValues.ULIMIT,
                            StepValues.LLIMIT)))
                        {
                            iSuccess++;
                        }
						else
						{
							Counters.iSetErrorCounterErrors++;
						}
                    }
                    else
                    {
                        // Debug mode
                        iSuccess++;
                    }

                    // GENY recommend 3 successful attempts
                    if (iSuccess >= 3)
                    {
                        break;
                    }
                }
            }

            if (iSuccess < 3)
            {
                return -1;
            }

            return 0;
        }

        private int ReadErrors(DoWorkEventArgs e, int iStep)
        {
            bool bResult = true;
            string strResult;

            LogMessage(MSG.DEBUG, "ReadErrors()");

            Counters.iReadErrorCounter++;

            bResult = OpenBox(BOX_ERRCOUNTER);

            if (!bResult)
            {
                return -1;
            }

            LogMessage(MSG.INFO, "ErrCounterReadData(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + ")");

            for (short iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    short sErrTimes = 0;
                    string strDataStr = "";

                    if (BgWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        return 1;
                    }

                    if (Settings.bRunCtrComm)
                    {
                        if (CheckResult(_CtrComm.ErrCounterReadData(
                            Bench.ctrlSio,
                            Bench.ctrlSioFmt,
                            iPos,
                            ref sErrTimes,
                            ref strDataStr,
                            0)))
                        {
                            LogMessage(MSG.DEBUG, "ErrCounterReadData(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + iPos + ")" + " = " + strDataStr);

                            if (strDataStr == null)
                            {
                                strResult = "a+11.11"; // Marker
                            }
                            else
                            {
                                strResult = strDataStr;
                            }

                            if (strResult.Length >= 2 && strResult.Length <= 16)
                            {
                                switch (StepValues.TestTypeID)
                                {
                                    case 2: // Start Test
                                    case 3: // Creep Test

                                        strResult = strResult.ToLower();

                                        if (Regex.IsMatch(strResult, "pass"))
                                        {
                                            TestBoard[iPos].ErrorV[iStep] = iPASS;
                                        }

                                        if (Regex.IsMatch(strResult, "fail"))
                                        {
                                            TestBoard[iPos].ErrorV[iStep] = iFAIL;
                                        }

                                        break;

                                    case 4: // Error Test

                                        if (strResult.StartsWith("{") || strResult.StartsWith("}") || strResult.StartsWith("a"))
                                        {
                                            TestBoard[iPos].ErrorV[iStep] = ConvertToDouble(strResult.Substring(2));
                                        }
                                        else
                                        {
                                            TestBoard[iPos].ErrorV[iStep] = dDEFAULT; // Marker
                                        }

                                        break;
                                }
                            }
                        }
                        else
                        {
							Counters.iReadErrorCounterErrors++;
                            //return -1; // We'll never complete a batch!
                        }
                    }
                    else
                    {
                        GenerateErrors(iStep, iPos);
                        LogMessage(MSG.DEBUG, "ErrCounterReadData(" + Bench.ctrlSio + "," + Bench.ctrlSioFmt + "," + iPos + ")" + " = " + TestBoard[iPos].ErrorV[iStep]);
                    }
                }
            }

            return 0;
        }

        private void UpdateGuiStepValues(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiStepValues), new object[] { iStep });
                return;
            }

            if (StepValues.LLIMIT == dDEFAULT) // Marker
            {
                labelLimitsL.Text = "";
            }
            else
            {
                labelLimitsL.Text = StepValues.LLIMIT.ToString();
            }

            if (StepValues.ULIMIT == dDEFAULT) // Marker
            {
                labelLimitsU.Text = "";
            }
            else
            {
                labelLimitsU.Text = StepValues.ULIMIT.ToString();
            }

            labelBaseUb.Text = StepValues.BaseUb.ToString();
            labelBaseIb.Text = StepValues.BaseIb.ToString();
            labelBaseIm.Text = StepValues.BaseImax.ToString();
        }

        private void UpdateGuiProcedureStep(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiProcedureStep), new object[] { iStep });
                return;
            }

            DGV_ProcedureRun.CurrentCell = DGV_ProcedureRun.Rows[iStep].Cells[0];
            DGV_ProcedureRun.Rows[iStep].Selected = true;
            listBox01to24.Items.Clear();
            listBox25to48.Items.Clear();
        }

        private void UpdateGuiActuals(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiActuals), new object[] { iStep });
                return;
            }

            textBox_AVA.Text = Convert.ToDouble(Actuals.uA).ToString();
            textBox_AVB.Text = Convert.ToDouble(Actuals.uB).ToString();
            textBox_AVC.Text = Convert.ToDouble(Actuals.uC).ToString();
            textBox_AIA.Text = Convert.ToDouble(Actuals.iA).ToString();
            textBox_AIB.Text = Convert.ToDouble(Actuals.iB).ToString();
            textBox_AIC.Text = Convert.ToDouble(Actuals.iC).ToString();
            textBox_APHIA.Text = Convert.ToDouble(Actuals.phiA).ToString();
            textBox_APHIB.Text = Convert.ToDouble(Actuals.phiB).ToString();
            textBox_APHIC.Text = Convert.ToDouble(Actuals.phiC).ToString();

            labelFreq.Text = Convert.ToDouble(Actuals.freq).ToString("F1");
            labelUA.Text = Convert.ToDouble(Actuals.uA).ToString("F1");
            labelUB.Text = Convert.ToDouble(Actuals.uB).ToString("F1");
            labelUC.Text = Convert.ToDouble(Actuals.uC).ToString("F1");
            labelIA.Text = Convert.ToDouble(Actuals.iA).ToString("F2");
            labelIB.Text = Convert.ToDouble(Actuals.iB).ToString("F2");
            labelIC.Text = Convert.ToDouble(Actuals.iC).ToString("F2");
            labelPhiA.Text = Convert.ToDouble(Actuals.phiA).ToString("F1");
            labelPhiB.Text = Convert.ToDouble(Actuals.phiB).ToString("F1");
            labelPhiC.Text = Convert.ToDouble(Actuals.phiC).ToString("F1");
            labelPW.Text = Convert.ToDouble(Actuals.totalP).ToString("F1");
            labelPQ.Text = Convert.ToDouble(Actuals.totalQ).ToString("F1");
            labelPS.Text = Convert.ToDouble(Actuals.totalS).ToString("F1");
        }

        private void UpdateGuiDesiredVoltage(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiDesiredVoltage), new object[] { iStep });
                return;
            }

            textBox_VA.Text = Convert.ToInt32(Desired.uA).ToString();
            textBox_VB.Text = Convert.ToInt32(Desired.uB).ToString();
            textBox_VC.Text = Convert.ToInt32(Desired.uC).ToString();
        }

        private void UpdateGuiActualVoltageColor(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiActualVoltageColor), new object[] { iStep });
                return;
            }

            if (Electricals.bA) textBox_AVA.BackColor = Color.Green;
            if (Electricals.bB) textBox_AVB.BackColor = Color.Green;
            if (Electricals.bC) textBox_AVC.BackColor = Color.Green;
        }

        private void UpdateGuiDesiredCurrent(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiDesiredCurrent), new object[] { iStep });
                return;
            }

            textBox_IA.Text = Convert.ToInt32(Desired.iA).ToString();
            textBox_IB.Text = Convert.ToInt32(Desired.iB).ToString();
            textBox_IC.Text = Convert.ToInt32(Desired.iC).ToString();
            textBox_PHIA.Text = Convert.ToInt32(Desired.phiA).ToString();
            textBox_PHIB.Text = Convert.ToInt32(Desired.phiB).ToString();
            textBox_PHIC.Text = Convert.ToInt32(Desired.phiC).ToString();
        }

        private void UpdateGuiActualCurrentColor(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateGuiActualCurrentColor), new object[] { iStep });
                return;
            }

            if (Electricals.bA) textBox_AIA.BackColor = Color.Green;
            if (Electricals.bB) textBox_AIB.BackColor = Color.Green;
            if (Electricals.bC) textBox_AIC.BackColor = Color.Green;
            if (Electricals.bA) textBox_APHIA.BackColor = Color.Green;
            if (Electricals.bB) textBox_APHIB.BackColor = Color.Green;
            if (Electricals.bC) textBox_APHIC.BackColor = Color.Green;
        }

        public bool WriteToFile(string strFname, string strText)
        {
            int iRetries = 5;

            LogMessage(MSG.DEBUG, "WriteToFile(" + strFname + ")");

            while (iRetries-- > 0)
            {
                try
                {
                    using (StreamWriter swFile = new StreamWriter(Settings.strRootPath + @"\" + strFname))
                    {
                        swFile.WriteLine(strText);
                    }

                    return false;
                }
                catch
                {
                    LogMessage(MSG.WARNING, "File : " + strFname + "- Conflict with other class instance or application, retrying...");
                    Thread.Sleep(100);
                    Application.DoEvents();
                }
            }

            LogMessage(MSG.ERROR, "File : " + strFname + "- Conflict with other class instance or application, giving up!");

            return true;
        }

        private void GenerateErrors(int iStep, int iPos)
        {
            double dVal;

            if (Settings.bErrorsGenerate)
            {
                if (StepValues.TestTypeID == 2 || StepValues.TestTypeID == 3)
                {
                    TestBoard[iPos].ErrorV[iStep] = (iPos % 2) == 1 ? iFAIL : iPASS;

                    if ((iPos % 15) == 0)
                    {
                        TestBoard[iPos].ErrorV[iStep] = dDEFAULT;
                    }
                }
                else
                {
                    if (Settings.bErrorsType)
                    {
                        // Values = Position.StepNo
                        dVal = (iStep + 1) * 0.01;
                        TestBoard[iPos].ErrorV[iStep] = iPos + dVal;
                    }
                    else
                    {
                        // Values = +/- random within dErrorsRange
                        dVal = Random.Next(0, 1 + (int)(Settings.dErrorsRange * 100) * 2);
                        dVal = dVal - (int)(Settings.dErrorsRange * 100);
                        dVal = dVal / 100;
                        TestBoard[iPos].ErrorV[iStep] = dVal;
                    }
                }
            }
            else
            {
                if (StepValues.TestTypeID == 2 || StepValues.TestTypeID == 3)
                {
                    TestBoard[iPos].ErrorV[iStep] = iPASS;
                }
                else
                {
                    // Values = +/- random within 0.2
                    dVal = Random.Next(0, 1 + (int)(0.2 * 100) * 2);
                    dVal = dVal - (int)(0.2 * 100);
                    dVal = dVal / 100;
                    TestBoard[iPos].ErrorV[iStep] = dVal;
                }
            }
        }

		private bool CreateActuals()
        {
            Actuals.uA = Desired.uA;
            Actuals.iA = Desired.iA;
            Actuals.phiA = Desired.phiA;

            if (Bench.numPhase > 1)
            {
                Actuals.uB = Desired.uB;
                Actuals.uC = Desired.uC;
                Actuals.iB = Desired.iB;
                Actuals.iC = Desired.iC;
                Actuals.phiB = Desired.phiB;
                Actuals.phiC = Desired.phiC;
            }

            Actuals.totalP = Desired.uA * Desired.iA;
            Actuals.totalQ = 0.5 * Desired.uA * Desired.iA;
            Actuals.totalS = 1000;

            return false;
        }

        public bool CheckResult(string strReturn)
        {
            LogMessage(MSG.DEBUG, "CheckResult()");

            if (strReturn == null)
            {
                LogMessage(MSG.ERROR, "CtrComm API Returned : NULL");
                return false;
            }

            if (strReturn.CompareTo(strSuccess) == 0)
            {
                return true;
            }

            //LogMessage(MSG.ERROR, "CtrComm API Returned : " + strReturn); // FIX: Occasionally throws exception
            LogMessage(MSG.ERROR, "CtrComm API Returned : ERROR");

            return false;
        }

        public int SetConfig(int iConfig)
        {
            int iResult = 0;
            byte[] ba = new byte[10];

	        if (Bench.numPhase == 1)
	        {
		        return 0;
	        }

            ba[0]=0x02;
            ba[1]=0xAA;
            ba[2] = (byte)'S';
            ba[3] = (byte)'e';
            ba[4] = (byte)'t';
            ba[5] = (byte)':';
            ba[6] = (iConfig / 10 == 3 ? (byte)'2' : (byte)'0');
            ba[7] = (byte)'0';
            ba[8] = (byte)'0';
            ba[9]=0xAA;

            SerialPort ComPort = new SerialPort();

            ComPort.PortName = "COM" + Bench.ctrlSio;
            ComPort.BaudRate = 19200;
            ComPort.DataBits = 8;
            ComPort.StopBits = (StopBits)1;
            ComPort.Parity = (Parity)0;
            ComPort.Handshake = (Handshake)0;
            ComPort.ReadTimeout = 2000;
            ComPort.WriteTimeout = 2000;

            LogMessage(MSG.DEBUG, "ComPort.Open(" + ComPort.PortName + "," + ComPort.BaudRate + "," + ComPort.DataBits + "," + ComPort.StopBits + "," + ComPort.Parity + "," + ComPort.Handshake + ")");

            try
            {
                ComPort.Open();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 1;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Write()");
                ComPort.Write(ba, 0, 10);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
            }

            Thread.Sleep(500);

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Close()");
                ComPort.Close();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 1;
            }

            return iResult;
        }

        public int SetPQS(int iMeasure)
        {
            int iResult = 0;
            byte[] ba = new byte[10];

            ba[0] = 0x02;
            ba[1] = 0xAA;
            ba[9] = 0xAA;

            switch (iMeasure)
            {
                case 2:
                    ba[2] = (byte)'R';
                    ba[3] = (byte)'e';
                    ba[4] = (byte)'a';
                    ba[5] = (byte)'c';
                    ba[6] = (byte)'t';
                    ba[7] = (byte)' ';
                    ba[8] = (byte)' ';
                    break;
                default:
                    ba[2] = (byte)'A';
                    ba[3] = (byte)'c';
                    ba[4] = (byte)'t';
                    ba[5] = (byte)'i';
                    ba[6] = (byte)'v';
                    ba[7] = (byte)'e';
                    ba[8] = (byte)' ';
                    break;
            }

            SerialPort ComPort = new SerialPort();

            ComPort.PortName = "COM" + Bench.ctrlSio;
            ComPort.BaudRate = 19200;
            ComPort.DataBits = 8;
            ComPort.StopBits = (StopBits)1;
            ComPort.Parity = (Parity)0;
            ComPort.Handshake = (Handshake)0;
            ComPort.ReadTimeout = 2000;
            ComPort.WriteTimeout = 2000;

            LogMessage(MSG.DEBUG, "ComPort.Open(" + ComPort.PortName + "," + ComPort.BaudRate + "," + ComPort.DataBits + "," + ComPort.StopBits + "," + ComPort.Parity + "," + ComPort.Handshake + ")");

            try
            {
                ComPort.Open();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 1;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Write()");
                ComPort.Write(ba, 0, 10);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
            }

            Thread.Sleep(500);

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Close()");
                ComPort.Close();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 1;
            }

            return iResult;
        }

        public int SetURange(double dVolt)
        {
            return 0;
        }

        public int SetIRange(double dCurrent)
        {
            int iResult = 0;
            byte[] ba = new byte[10];
            double dConstant = 1000000000;

	        ba[0] = 0x02;
	        ba[1] = 0xAA;
	        ba[2] = (byte)'I';
	        ba[3] = (byte)'R';
	        ba[4] = (byte)'1';
	        ba[5] = (byte)'A';
	        ba[6] = (byte)' ';
	        ba[7] = (byte)' ';
	        ba[8] = (byte)' ';
	        ba[9] = 0xAA;

	        if(dCurrent > 1.2)
	        {
		        ba[5] = (byte)'0';
                ba[6] = (byte)'A';
		        dConstant=100000000;

		        if(dCurrent >12.0)
		        {
			        ba[6] = (byte)'0';
                    ba[7] = (byte)'A';
			        dConstant=10000000;
		        }
	        }

            SerialPort ComPort = new SerialPort();

            ComPort.PortName = "COM" + Bench.ctrlSio;
            ComPort.BaudRate = 19200;
            ComPort.DataBits = 8;
            ComPort.StopBits = (StopBits)1;
            ComPort.Parity = (Parity)0;
            ComPort.Handshake = (Handshake)0;
            ComPort.ReadTimeout = 2000;
            ComPort.WriteTimeout = 2000;

            LogMessage(MSG.DEBUG, "ComPort.Open(" + ComPort.PortName + "," + ComPort.BaudRate + "," + ComPort.DataBits + "," + ComPort.StopBits + "," + ComPort.Parity + "," + ComPort.Handshake + ")");

            try
            {
                ComPort.Open();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 1;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Write()");
                ComPort.Write(ba, 0, 10);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
            }

            Thread.Sleep(1000);

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Close()");
                ComPort.Close();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 1;
            }

            dCurConst = dConstant;

            return iResult;
        }

        public bool ReadActuals1()
        {
            byte[] ba = new byte[10];
            byte[] data = new byte[360];

            ba[0] = 0x02;
            ba[1] = 0xAA;
            ba[2] = (byte)'D';
            ba[3] = (byte)'a';
            ba[4] = (byte)'t';
            ba[5] = (byte)'a';
            ba[6] = (byte)' ';
            ba[7] = (byte)' ';
            ba[8] = (byte)' ';
            ba[9] = 0xAA;

            SerialPort ComPort = new SerialPort();

            ComPort.PortName = "COM" + Bench.ctrlSio;
            ComPort.BaudRate = 19200;
            ComPort.DataBits = 8;
            ComPort.StopBits = (StopBits)1;
            ComPort.Parity = (Parity)0;
            ComPort.Handshake = (Handshake)0;
            ComPort.ReadTimeout = 5000;
            ComPort.WriteTimeout = 5000;

            LogMessage(MSG.DEBUG, "ComPort.Open(" + ComPort.PortName + "," + ComPort.BaudRate + "," + ComPort.DataBits + "," + ComPort.StopBits + "," + ComPort.Parity + "," + ComPort.Handshake + ")");

            try
            {
                ComPort.Open();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Write()");
                ComPort.Write(ba, 0, 10);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            Thread.Sleep(1000);

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Read()");
                ComPort.Read(data, 0, 360);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Close()");
                ComPort.Close();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            // Replace non-printable chars
            for (int i = 0; i < 360; i++)
            {
                if ((data[i] < 0x20) || (data[i] > 0x7e))
                {
                    data[i] = 0x2a;
                }
            }

            string strRefMeter = Encoding.ASCII.GetString(data);

            if (strRefMeter.Length < 360)
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid length: " + strRefMeter.Length);
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            LogMessage(MSG.DEBUG, strRefMeter);

            int iIndexD = strRefMeter.IndexOf("Data");
            int iIndexA = strRefMeter.IndexOf("A:");
            int iIndexB = strRefMeter.IndexOf("B:");
            int iIndexF = strRefMeter.IndexOf("FEQ=");
            int iIndexP = strRefMeter.IndexOf("PSUM=");

            if ((iIndexD == -1))
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: Data");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if ((iIndexA == -1))
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: A:");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if ((iIndexB == -1))
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: B:");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if ((iIndexF == -1))
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: FREQ");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if ((iIndexP == -1))
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: PSUM");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            string strA = strRefMeter.Substring(iIndexA + 2, (iIndexB - iIndexA) - 2);
            string strF = strRefMeter.Substring(iIndexF + 4, (iIndexP - iIndexF) - 4);

            LogMessage(MSG.DEBUG, "strA = " + strA);
            LogMessage(MSG.DEBUG, "strF = " + strF);

            string[] strAValues = strA.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (strAValues.Length == 9)
            {
                Actuals.uA = ConvertToDouble(strAValues[0]);
                Actuals.iA = ConvertToDouble(strAValues[2]);
                Actuals.phiA = ConvertToDouble(strAValues[6]);
                Actuals.totalP = ConvertToDouble(strAValues[4]);
                Actuals.totalQ = ConvertToDouble(strAValues[5]);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: Phase A");
                LogMessage(MSG.ERROR, strA);
                return false;
            }

            if (strF.Length > 2)
            {
                Actuals.freq = ConvertToDouble(strF);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: FREQ");
                LogMessage(MSG.ERROR, strF);
                return false;
            }

            return true;
        }

        public bool ReadActuals3()
        {
            byte[] ba = new byte[10];
            byte[] data = new byte[372];

            ba[0] = 0x02;
            ba[1] = 0xAA;
            ba[2] = (byte)'D';
            ba[3] = (byte)'a';
            ba[4] = (byte)'t';
            ba[5] = (byte)'a';
            ba[6] = (byte)' ';
            ba[7] = (byte)' ';
            ba[8] = (byte)' ';
            ba[9] = 0xAA;

            SerialPort ComPort = new SerialPort();

            ComPort.PortName = "COM" + Bench.ctrlSio;
            ComPort.BaudRate = 19200;
            ComPort.DataBits = 8;
            ComPort.StopBits = (StopBits)1;
            ComPort.Parity = (Parity)0;
            ComPort.Handshake = (Handshake)0;
            ComPort.ReadTimeout = 5000;
            ComPort.WriteTimeout = 5000;

            LogMessage(MSG.DEBUG, "ComPort.Open(" + ComPort.PortName + "," + ComPort.BaudRate + "," + ComPort.DataBits + "," + ComPort.StopBits + "," + ComPort.Parity + "," + ComPort.Handshake + ")");

            try
            {
                ComPort.Open();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Write()");
                ComPort.Write(ba, 0, 10);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            Thread.Sleep(1000);

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Read()");
                ComPort.Read(data, 0, 372);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            try
            {
                LogMessage(MSG.DEBUG, "ComPort.Close()");
                ComPort.Close();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return false;
            }

            // Replace non-printable chars
            for (int i = 0; i < 372; i++)
            {
                if ((data[i] < 0x20) || (data[i] > 0x7e))
                {
                    data[i] = 0x2a;
                }
            }

            string strRefMeter = Encoding.ASCII.GetString(data);

            if (strRefMeter.Length < 372)
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid length: " + strRefMeter.Length);
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            LogMessage(MSG.DEBUG, strRefMeter);

            int iIndexD = strRefMeter.IndexOf("Data");
            int iIndexA = strRefMeter.IndexOf("A:");
            int iIndexB = strRefMeter.IndexOf("B:");
            int iIndexC = strRefMeter.IndexOf("C:");
            int iIndexF = strRefMeter.IndexOf("FEQ=");
            int iIndexP = strRefMeter.IndexOf("PSUM=");
            int iIndexQ = strRefMeter.IndexOf("QSUM=");
            int iIndexS = strRefMeter.IndexOf("SSUM=");
            int iIndexE = strRefMeter.IndexOf("COSSUM=");

            if (iIndexD == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: Data");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexA == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: A:");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexB == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: B:");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexC == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: C:");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexF == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: FEQ=");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexP == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: PSUM=");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexQ == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: QSUM=");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexS == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: SSUM=");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            if (iIndexE == -1)
            {
                LogMessage(MSG.ERROR, "Reference meter - Tag not found: COSSUM=");
                LogMessage(MSG.ERROR, strRefMeter);
                return false;
            }

            string strA = strRefMeter.Substring(iIndexA + 2, (iIndexB - iIndexA) - 3);
            string strB = strRefMeter.Substring(iIndexB + 2, (iIndexC - iIndexB) - 3);
            string strC = strRefMeter.Substring(iIndexC + 2, (iIndexF - iIndexC) - 3);
            string strF = strRefMeter.Substring(iIndexF + 4, (iIndexP - iIndexF) - 5);
            string strP = strRefMeter.Substring(iIndexP + 5, (iIndexQ - iIndexP) - 6);
            string strQ = strRefMeter.Substring(iIndexQ + 5, (iIndexS - iIndexQ) - 6);
            string strS = strRefMeter.Substring(iIndexS + 5, (iIndexE - iIndexS) - 6);

            LogMessage(MSG.DEBUG, "strA = " + strA);
            LogMessage(MSG.DEBUG, "strB = " + strB);
            LogMessage(MSG.DEBUG, "strC = " + strC);
            LogMessage(MSG.DEBUG, "strF = " + strF);
            LogMessage(MSG.DEBUG, "strP = " + strP);
            LogMessage(MSG.DEBUG, "strQ = " + strQ);
            LogMessage(MSG.DEBUG, "strS = " + strS);

            string[] strAValues = strA.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (strAValues.Length == 9)
            {
                Actuals.uA = ConvertToDouble(strAValues[0]);
                Actuals.iA = ConvertToDouble(strAValues[2]);
                Actuals.phiA = ConvertToDouble(strAValues[7]);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: Phase A");
                LogMessage(MSG.ERROR, strA);
                return false;
            }

            string[] strBValues = strB.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (strBValues.Length == 9)
            {
                Actuals.uB = ConvertToDouble(strBValues[0]);
                Actuals.iB = ConvertToDouble(strBValues[2]);
                Actuals.phiB = ConvertToDouble(strBValues[7]);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: Phase B");
                LogMessage(MSG.ERROR, strB);
                return false;
            }

            string[] strCValues = strC.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (strCValues.Length == 9)
            {
                Actuals.uC = ConvertToDouble(strCValues[0]);
                Actuals.iC = ConvertToDouble(strCValues[2]);
                Actuals.phiC = ConvertToDouble(strCValues[7]);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: Phase C");
                LogMessage(MSG.ERROR, strC);
                return false;
            }

            if (strF.Length > 2)
            {
                Actuals.freq = ConvertToDouble(strF);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: FREQ");
                LogMessage(MSG.ERROR, strF);
                return false;
            }

            if (strP.Length > 2)
            {
                Actuals.totalP = ConvertToDouble(strP);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: PSUM");
                LogMessage(MSG.ERROR, strP);
                return false;
            }

            if (strQ.Length > 2)
            {
                Actuals.totalQ = ConvertToDouble(strQ);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: QSUM");
                LogMessage(MSG.ERROR, strQ);
                return false;
            }

            if (strS.Length > 2)
            {
                Actuals.totalS = ConvertToDouble(strS);
            }
            else
            {
                LogMessage(MSG.ERROR, "Reference meter - Invalid fields: SSUM");
                LogMessage(MSG.ERROR, strS);
                return false;
            }

            return true;
        }

        public bool OpenBox(short sBox)
        {
            bool bResult = true;
            int iRetries = 4;

            while (iRetries-- > 0)
            {
                LogMessage(MSG.DEBUG, "OpenBox(" + Bench.ctrlSio + "," + sBox + "," + Bench.ctrlSioFmt + ")");

                if (Settings.bRunCtrComm)
                {
                    bResult = CheckResult(_CtrComm.OpenBox(Bench.ctrlSio, sBox, Bench.ctrlSioFmt));
                }

                Thread.Sleep(200);

                if (bResult)
                {
                    break;
                }
            }

            return bResult;
        }

        private void CreateTestBoard()
        {
            int iCols;
            int iRows;

            LogMessage(MSG.DEBUG, "CreateTestBoard()");

            switch (Bench.numPosition)
            {
                case 20: iCols = 10 + 2; iRows = 4; break;
                case 24: iCols = 12 + 2; iRows = 4; break;
                case 32: iCols = 16 + 2; iRows = 4; break;
                case 48: iCols = 12 + 2; iRows = 7; break;
                default: iCols = 10 + 2; iRows = 4; break;
            }

            // Create buttons for meters
            for (int i = 0; i < Bench.numPosition; i++)
            {
                buttons[i] = new Button();
                buttons[i].Anchor = ((AnchorStyles)((((AnchorStyles.Top | AnchorStyles.Bottom)
                | AnchorStyles.Left)
                | AnchorStyles.Right)));
                buttons[i].BackColor = Color.Transparent;
                //buttons[i].BackgroundImage = global::EzyCal.Properties.Resources.MeterStart;
                buttons[i].BackgroundImage = null;
                buttons[i].BackgroundImageLayout = ImageLayout.Stretch;
                buttons[i].FlatStyle = FlatStyle.Popup;
                buttons[i].Font = new Font("Microsoft Sans Serif", 8F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
                buttons[i].ForeColor = Color.Black;
                //buttons[i].Location = new Point(70, 103);
                buttons[i].Margin = new Padding(10, 10, 10, 10);
                buttons[i].Name = "button" + (i + 1).ToString();
                //buttons[i].Size = new Size(63, 81);
                buttons[i].Text = (i + 1).ToString();
                buttons[i].UseVisualStyleBackColor = false;
                buttons[i].Click += new System.EventHandler(this.Button_Meter_Click);
            }

            TableLayoutPanel TLP_TestBoard = new TableLayoutPanel();
            TLP_TestBoard.SuspendLayout();
            tabControl1.TabPages["Board"].Controls.Add(TLP_TestBoard);
            TLP_TestBoard.BackColor = System.Drawing.Color.White;
            TLP_TestBoard.BackgroundImageLayout = ImageLayout.None;
            TLP_TestBoard.ColumnCount = iCols;

            float dWidth = 100 / iCols;

            for (int i = 0; i < iCols; i++)
            {
                TLP_TestBoard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, dWidth));
            }

            int iMeter = 0;

            for (int i = 1; i < (iCols - 1); i++)
            {
                TLP_TestBoard.Controls.Add(buttons[iMeter++], i, 1);
            }

            switch (Bench.numPosition)
            {
                case 20: iMeter = 19; ; break;
                case 24: iMeter = 23; ; break;
                case 32: iMeter = 31; ; break;
                case 48: iMeter = 23; ; break;
                default: iMeter = 19; ; break;
            }

            for (int i = 1; i < (iCols - 1); i++)
            {
                TLP_TestBoard.Controls.Add(buttons[iMeter--], i, 2);
            }

            if (Bench.numPosition == 48)
            {
                iMeter = 24;

                for (int i = 1; i < (iCols - 1); i++)
                {
                    TLP_TestBoard.Controls.Add(buttons[iMeter++], i, 4);
                }

                iMeter = 47;

                for (int i = 1; i < (iCols - 1); i++)
                {
                    TLP_TestBoard.Controls.Add(buttons[iMeter--], i, 5);
                }
            }

            TLP_TestBoard.Dock = DockStyle.Fill;
            TLP_TestBoard.ForeColor = System.Drawing.Color.Black;
            TLP_TestBoard.Margin = new Padding(2);
            TLP_TestBoard.Name = "TLP_TestBoard";
            TLP_TestBoard.Padding = new Padding(2, 2, 2, 2);
            TLP_TestBoard.RowCount = iRows;

            for (int i = 0; i < iRows; i++)
            {
                TLP_TestBoard.RowStyles.Add(new RowStyle(SizeType.Percent, 10));
            }

            TLP_TestBoard.Size = new System.Drawing.Size(987, 453);
            TLP_TestBoard.TabIndex = 55;
            TLP_TestBoard.ResumeLayout(false);
            TLP_TestBoard.PerformLayout();
        }

        private double ConvertToDouble(string strInput)
        {
            double dValue = 0;

            if (string.IsNullOrEmpty(strInput))
            {
                return dDEFAULT; // Marker
            }

            try
            {
                dValue = Convert.ToDouble(strInput);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return dDEFAULT; // Marker
            }

            return dValue;
        }

        // For property number
        private UInt64 ConvertToUInt64(string strInput)
        {
            UInt64 i64Value = 0;

            if (string.IsNullOrEmpty(strInput))
            {
                return 0;
            }

            try
            {
                i64Value = Convert.ToUInt64(strInput, 10);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return 0;
            }

            return i64Value;
        }

        private void UpdateDisplayErrors(int iStep)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateDisplayErrors), new object[] { iStep });
                return;
            }

            LogMessage(MSG.DEBUG, "UpdateDisplayErrors(" + iStep + ")");

            string strListBoxValue = null;
            int iRow = 0;

            listBox01to24.Items.Clear();
            listBox25to48.Items.Clear();

            // Don't rely on StepValues, worker thread can be one step ahead of GUI update
            double dUlimit = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[17].Value.ToString());
            double dLlimit = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[18].Value.ToString());

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    double dVal = TestBoard[iPos].ErrorV[iStep];

                    if (dVal == iPASS)
                    {
                        strListBoxValue = (iPos).ToString("D2") + ":" + "Pass";
                    }
                    else if (dVal == iFAIL)
                    {
                        LogMessage(MSG.DEBUG, "Position: " + (iPos).ToString("D2") + ", Step: " + (iStep + 1).ToString("D2") + ", Value: Fail");
                        strListBoxValue = (iPos).ToString("D2") + ":" + "Fail <<";
                    }
                    else if ((dVal < dLlimit) || (dVal > dUlimit))
                    {
                        LogMessage(MSG.DEBUG, "Position: " + (iPos).ToString("D2") + ", Step: " + (iStep + 1).ToString("D2") + ", Value: " + dVal);
                        strListBoxValue = (iPos).ToString("D2") + ":" + dVal.ToString("+00.00;-00.00") + " <<";
                    }
                    else
                    {
                        strListBoxValue = (iPos).ToString("D2") + ":" + dVal.ToString("+00.00;-00.00");
                    }

                    if (iRow < 24)
                    {
                        listBox01to24.Items.Add(strListBoxValue);
                    }
                    else
                    {
                        listBox25to48.Items.Add(strListBoxValue);
                    }

                    iRow++;
                }
            }
        }

        private void Button_FindOrder_Click(object sender, EventArgs e)
        {
            ButtonsEnabled(false);
            LogMessage(MSG.DEBUG, "Button_FindOrder_Click()");
            FindOrder();
            ButtonsEnabled(true);
        }

        private void FindOrder()
        {
            LogMessage(MSG.DEBUG, "FindOrder()");

            if (textBoxContractNo.Text.Length != 6)
            {
                LogMessage(MSG.ERROR, "Invalid Purchase Order number");
                return;
            }

            SqlDataReader SqlReader = null;
            SqlConnection SqlConnection = new SqlConnection(strSqlConnection);

            try
            {
                LogMessage(MSG.DEBUG, "SqlConnection.Open()");
                SqlConnection.Open();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return;
            }

            try
            {
                string strSqlCommand = "SELECT * FROM ProductionInstructions WHERE [Production Order] = " + textBoxContractNo.Text;
                SqlCommand SqlCommand = new SqlCommand(strSqlCommand, SqlConnection);
                LogMessage(MSG.DEBUG, strSqlCommand);
                LogMessage(MSG.DEBUG, "SqlCommand.ExecuteReader()");
                SqlReader = SqlCommand.ExecuteReader();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return;
            }

            try
            {
                LogMessage(MSG.DEBUG, "SqlReader.Read()");

                if (SqlReader.Read())
                {
                    comboBoxClient.Items.Clear();
                    comboBoxClient.Items.Add(SqlReader["Customer"].ToString());
                    comboBoxClient.SelectedItem = SqlReader["Customer"].ToString();

                    comboBoxMeterType.Items.Clear();
                    comboBoxMeterType.Items.Add(SqlReader["Meter Type"].ToString());
                    comboBoxMeterType.SelectedItem = SqlReader["Meter Type"].ToString();

                    comboBoxFirmware.Items.Clear();
                    comboBoxFirmware.Items.Add(SqlReader["Firmware"].ToString());
                    comboBoxFirmware.SelectedItem = SqlReader["Firmware"].ToString();

                    comboBoxCustomerProgram.Items.Clear();
                    comboBoxCustomerProgram.Items.Add(SqlReader["Customer Program"].ToString());
                    comboBoxCustomerProgram.SelectedItem = SqlReader["Customer Program"].ToString();

                    comboBoxRippleProgram.Items.Clear();
                    comboBoxRippleProgram.Items.Add(SqlReader["Ripple Program"].ToString());
                    comboBoxRippleProgram.SelectedItem = SqlReader["Ripple Program"].ToString();

                    textBoxOwnersNo.Text = SqlReader["CSN From"].ToString();
                }
                else
                {
                    LogMessage(MSG.ERROR, "Production Order Not Found : " + textBoxContractNo.Text);
                }
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return;
            }

            try
            {
                LogMessage(MSG.DEBUG, "SqlConnection.Close()");
                SqlConnection.Close();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
            }
        }

        private void Button_Meter_Click(object sender, EventArgs e)
        {
            ButtonsEnabled(false);

            int iPos = (int)ConvertToDouble(((Button)sender).Text);

            LogMessage(MSG.DEBUG, "Button_Meter_Click()");

            if (TestBoard[iPos].Active)
            {
                TestBoard[iPos].Active = false;

                using (StreamWriter StreamWriter = File.AppendText(Settings.strRootPath + @"\Bypassed.txt"))
                {
                    StreamWriter.WriteLine(TestBoard[iPos].OwnerNo);
                }	
            }
            else
            {
                TestBoard[iPos].Active = true;

                string[] strLines = null;

                if (File.Exists(Settings.strRootPath + @"\Bypassed.txt") == true)
                {
                    strLines = File.ReadAllLines(Settings.strRootPath + @"\Bypassed.txt");
                    Array.Sort(strLines);

                    using (StreamWriter StreamWriter = new StreamWriter(Settings.strRootPath + @"\Bypassed.txt"))
                    {
                        foreach (string strLine in strLines)
                        {
                            if (strLine != TestBoard[iPos].OwnerNo)
                            {
                                StreamWriter.WriteLine(strLine);
                            }
                        }
                    }
                }
            }

            UpdateGridViewBatch();

            ButtonsEnabled(true);
        }

        private void UpdateGridViewBatch()
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action(UpdateGridViewBatch));
                return;
            }

            LogMessage(MSG.DEBUG, "UpdateGridViewBatch()");

            DGV_Batch.Columns.Clear();

            // Create a New DataTable to store the Data
            DataTableBatch = new DataTable("Batch");

            // Create the Columns in the DataTable
            DataColumn c01 = new DataColumn("Pos");
            DataColumn c02 = new DataColumn("Status");
            DataColumn c03 = new DataColumn("Meter Type");
            DataColumn c04 = new DataColumn("MSN");
            DataColumn c05 = new DataColumn("Owner No");
            DataColumn c06 = new DataColumn("Contract No");
            DataColumn c07 = new DataColumn("Client");
            DataColumn c08 = new DataColumn("Client No");
            DataColumn c09 = new DataColumn("Firmware");
            DataColumn c10 = new DataColumn("MAC");
            DataColumn c11 = new DataColumn("DateTime");
            DataColumn c12 = new DataColumn("Program");
            DataColumn c13 = new DataColumn("Ripple");

            // Add the Created Columns to the Datatable
            DataTableBatch.Columns.Add(c01);
            DataTableBatch.Columns.Add(c02);
            DataTableBatch.Columns.Add(c03);
            DataTableBatch.Columns.Add(c04);
            DataTableBatch.Columns.Add(c05);
            DataTableBatch.Columns.Add(c06);
            DataTableBatch.Columns.Add(c07);
            DataTableBatch.Columns.Add(c08);
            DataTableBatch.Columns.Add(c09);
            DataTableBatch.Columns.Add(c10);
            DataTableBatch.Columns.Add(c11);
            DataTableBatch.Columns.Add(c12);
            DataTableBatch.Columns.Add(c13);

            iMeters = 0;

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    iMeters++;

                    DataRow DataRow = DataTableBatch.NewRow();

                    DataRow["Pos"] = iPos.ToString("D2");
                    DataRow["Status"] = TestBoard[iPos].Status;
                    DataRow["Meter Type"] = TestBoard[iPos].MeterType;
                    DataRow["MSN"]    = TestBoard[iPos].MSN;
                    DataRow["Owner No"]  = TestBoard[iPos].OwnerNo;
                    DataRow["Contract No"] = TestBoard[iPos].ContractNo;
                    DataRow["Client"] = TestBoard[iPos].Client;
                    DataRow["Client No"] = TestBoard[iPos].ClientNo;
                    DataRow["Firmware"] = TestBoard[iPos].Firmware;
                    DataRow["MAC"]      = TestBoard[iPos].MAC;
                    DataRow["Datetime"] = TestBoard[iPos].DateTime;
                    DataRow["Program"]  = TestBoard[iPos].Program;
                    DataRow["Ripple"]   = TestBoard[iPos].Ripple;

                    DataTableBatch.Rows.Add(DataRow);
                }
            }

            DGV_Batch.DataSource = DataTableBatch;

            Padding newPadding = new Padding(15, 1, 15, 1);
            DGV_Batch.RowTemplate.DefaultCellStyle.Padding = newPadding;

            /*
            int iRow = iMeters-1;

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    if (TestBoard[iPos].FailedFW > 0)
                    {
                        DGV_Batch.Rows[iRow].Cells[7].Style.BackColor = Color.Red;
                    }

                    if (TestBoard[iPos].FailedMSN > 0)
                    {
                        DGV_Batch.Rows[iRow].Cells[2].Style.BackColor = Color.Red;
                    }

                    if (TestBoard[iPos].FailedMAC > 0)
                    {
                        DGV_Batch.Rows[iRow].Cells[8].Style.BackColor = Color.Red;
                    }

                    if (TestBoard[iPos].FailedClock > 0)
                    {
                        DGV_Batch.Rows[iRow].Cells[9].Style.BackColor = Color.Red;
                    }

                    iRow--;
                }
            }
            */

            int m = 0;

            // Hide/show results columns
            foreach (DataGridViewColumn col in DGV_Results.Columns)
            {
                if ((m >= 3) && (m < (Bench.numPosition + 3)))
                {
                    if (TestBoard[m - 2].Active)
                    {
                        col.Visible = true;
                    }
                    else
                    {
                        col.Visible = false;
                    }
                }

                m++;
            }

            UpdateTestBoard();
        }

        /*
        private void CheckMSNs()
        {
            int iRow = 0;
            bool bFailed = false;

            LogMessage(MSG.DEBUG, "CheckMSNs()");

            if (DGV_Batch.Rows.Count <= 0)
            {
                return;
            }

            string[] series = { "U12", "U13", "U33", "U34", "U35" };

            int j;

            for (j = 0; j < series.Length; j++)
            {
                if (Regex.IsMatch(comboBoxMeterType.Text, series[j]))
                {
                    break;
                }
            }

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    bFailed = false;

                    if (TestBoard[iPos].MSN == "MSN?")
                    {
                        iRow++;
                        TestBoard[iPos].Passed++;
                        continue;
                    }

                    if (TestBoard[iPos].MSN.Length != 16)
                    {
                        bFailed = true;
                    }

                    if (j < series.Length)
                    {
                        if (!Regex.IsMatch(TestBoard[iPos].MSN, series[j]))
                        {
                            bFailed = true;
                        }
                    }
                    else
                    {
                        bFailed = true;
                    }

                    if (bFailed)
                    {
                        TestBoard[iPos].Failed++;
                        TestBoard[iPos].FailedMSN++;
                        LogMessage(MSG.FAILURE, "Meter position " + iPos + " invalid MSN");
                    }
                    else
                    {
                        TestBoard[iPos].Passed++;
                    }

                    iRow++;
                }
            }
        }

        private void CheckMacAddresses()
        {
            int iRow = 0;
            bool bFailed = false;
            double dValue = 0;

            LogMessage(MSG.DEBUG, "CheckMacAddresses()");

            if (DGV_Batch.Rows.Count <= 0)
            {
                return;
            }

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    bFailed = false;

                    if (TestBoard[iPos].MAC == "MAC?")
                    {
                        iRow++;
                        TestBoard[iPos].Passed++;
                        continue;
                    }

                    if (TestBoard[iPos].MAC.Length != 16)
                    {
                        bFailed = true;
                    }

                    try
                    {
                        dValue = Convert.ToUInt64(TestBoard[iPos].MAC, 16);
                    }
                    catch (Exception ex)
                    {
                        LogMessage(MSG.FAILURE, ex.Message);
                        bFailed = true;
                    }

                    if (textBoxCommsModule.Text != "None")
                    {
                        if (dValue == 0xFFFFFFFFFFFFFFFF)
                        {
                            bFailed = true;
                        }
                    }

                    if (bFailed)
                    {
                        TestBoard[iPos].Failed++;
                        TestBoard[iPos].FailedMAC++;
                        LogMessage(MSG.FAILURE, "Meter position " + iPos + " invalid MAC address");
                    }
                    else
                    {
                        TestBoard[iPos].Passed++;
                    }

                    iRow++;
                }
            }
        }

        private void CheckMeterClock(int iPos)
        {
            bool bFailed = false;
            double dTimeDiff = 0;
            int iTimeError = 10;

            LogMessage(MSG.DEBUG, "CheckMeterClock()");

            if (ToolStripMenuItem_Allow60s.Checked)
            {
                iTimeError = 60;
            }

            string[] arySegments = TestBoard[iPos].DateTime.Split(' ');

            if (arySegments.Length != 4)
            {
                bFailed = true;
            }
            else
            {
                arySegments[3] = arySegments[3].Trim(new Char[] { ' ', '[', ']' });

                TimeSpan result;

                if (TimeSpan.TryParse(arySegments[3], out result))
                {
                    dTimeDiff = result.TotalSeconds;
                    TestBoard[iPos].Values[iStep] = dTimeDiff;
                }
                else
                {
                    bFailed = true;
                }
            }

            if (bFailed || (Math.Abs(dTimeDiff) > iTimeError))
            {
                TestBoard[iPos].Failed++;
                TestBoard[iPos].FailedClock++;
                LogMessage(MSG.FAILURE, "Meter position " + iPos + " time check failed");
            }
            else
            {
                TestBoard[iPos].Passed++;
            }
        }
        */

        private void UpdateTestBoard()
        {
            LogMessage(MSG.DEBUG, "UpdateTestBoard()");

            for (int iPos = 0; iPos < Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos+1].Active)
                {
                    bool bFailed = false;

                    TestBoard[iPos + 1].Failed = 0;

                    for (int j = 0; j < iMaxSteps; j++)
                    {
                        if (TestBoard[iPos + 1].ErrorR[j] > 0)
                        {
                            bFailed = true;
                        }
                    }

                    if (bFailed)
                    {
                        buttons[iPos].BackgroundImage = global::EzyCal.Properties.Resources.MeterFail;
                        buttons[iPos].ForeColor = System.Drawing.Color.Black;
                        TestBoard[iPos + 1].Failed = 1;
                    }
                    else if ((!bFailed) && (TestBoard[iPos + 1].Saved == true))
                    {
                        buttons[iPos].BackgroundImage = global::EzyCal.Properties.Resources.MeterPass;
                        buttons[iPos].ForeColor = System.Drawing.Color.Black;
                    }
                    else
                    {
                        buttons[iPos].BackgroundImage = global::EzyCal.Properties.Resources.MeterStart;
                        buttons[iPos].ForeColor = System.Drawing.Color.White;
                    }
                }
                else
                {
                    buttons[iPos].BackColor = System.Drawing.Color.Transparent;
                    buttons[iPos].FlatStyle = FlatStyle.Popup;
                    buttons[iPos].ForeColor = System.Drawing.Color.Black;
                    buttons[iPos].BackgroundImage = null;
                }
            }
        }

        // CHED AMI Property Numbers
        private string ChedCheckSum(string strOwnerNo)
        {
            int iSum = 0;
            int iCheckSum = 0;
            int iLength = strOwnerNo.Length;
            byte[] bNewOwnerNo = new byte[iLength + 1];

            for (int n = 0; n < iLength; n++)
            {
                iSum += strOwnerNo[n] * (n % 2 == 0 ? 2 : 1);
            }

            iCheckSum = ((iSum / 10) * 10 + 10) - iSum;

            if (iCheckSum == 10)
            {
                iCheckSum = 0;
            }

            bNewOwnerNo[0] = (byte)strOwnerNo[0];
            bNewOwnerNo[1] = (byte)(48 + iCheckSum); // ASCII

            for (int n = 2; n < iLength + 1; n++)
	        {
                bNewOwnerNo[n] = (byte)strOwnerNo[n - 1];
            }

            return Encoding.ASCII.GetString(bNewOwnerNo);
        }

        private void BatchAdd(int iFirst, int iLast)
        {
            int i;
            string strNumValue;
            string strCharValue;
            string[] strLines = null;

            LogMessage(MSG.DEBUG, "BatchAdd(" + iFirst + ", " + iLast + ")");

            if (string.IsNullOrEmpty(textBoxContractNo.Text))
            {
                LogMessage(MSG.ERROR, "Please enter Contract No.");
                return;
            }

            if (!String.Equals(textBoxContractNo.Text, "TBA", StringComparison.OrdinalIgnoreCase))
            {
                if (textBoxContractNo.Text.Length != 6)
                {
                    LogMessage(MSG.ERROR, "Invalid Contract No. - Length");
                    return;
                }

                if (!Regex.IsMatch(textBoxContractNo.Text, @"^[0-9]+$"))
                {
                    LogMessage(MSG.ERROR, "Invalid Contract No. - Invalid characters");
                    return;
                }
            }

            if (string.IsNullOrEmpty(textBoxOwnersNo.Text))
            {
                LogMessage(MSG.ERROR, "Please enter Owners No.");
                return;
            }

            if (textBoxOwnersNo.Text.Any(Char.IsWhiteSpace))
            {
                LogMessage(MSG.ERROR, "Invalid Owners No. - White space");
                return;
            }

            int iOwnersNoLength = textBoxOwnersNo.Text.Length;
            int iCheckLength = iOwnersNoLength;

            if (checkBoxChed.Checked == true)
            {
                iCheckLength += 1;
            }

            if ((iOwnersNoLength < 4) || (iOwnersNoLength > 19))
            {
                LogMessage(MSG.ERROR, "Invalid Owners No. - Length");
                return;
            }

            if (string.IsNullOrEmpty(comboBoxClient.Text))
            {
                LogMessage(MSG.ERROR, "Please select a Client");
                return;
            }

            if (string.IsNullOrEmpty(comboBoxMeterType.Text))
            {
                LogMessage(MSG.ERROR, "Please select Meter Type");
                return;
            }

            // Determine starting index of numerical value.
            // e.g. ABC123DEF00001 -> Increment 00001 only
            for (i = iOwnersNoLength - 1; i > 0; i--)
            {
                if (!Char.IsDigit(textBoxOwnersNo.Text, i))
                {
                    break;
                }
            }

            if (Regex.IsMatch(textBoxOwnersNo.Text, @"^[0-9]+$"))
            {
                strNumValue = textBoxOwnersNo.Text;
                strCharValue = "";
            }
            else
            {
                strNumValue = textBoxOwnersNo.Text.Substring(i + 1);
                strCharValue = textBoxOwnersNo.Text.Substring(0, i + 1);
                strCharValue = strCharValue.ToUpper();
            }

            UInt64 i64OwnerNum = ConvertToUInt64(strNumValue);

            if (strNumValue.Length < 2)
            {
                LogMessage(MSG.ERROR, "Invalid Owners No. - Number length");
                return;
            }

            if (checkBoxSaved.Checked == true)
            {
                if (File.Exists(Settings.strRootPath + @"\Bypassed.txt") == true)
                {
                    strLines = File.ReadAllLines(Settings.strRootPath + @"\Bypassed.txt");
                    Array.Sort(strLines);
                }
            }

            int iSavedCount = 1;

            for (int iPos = iFirst; iPos <= iLast; iPos++)
            {
                TestBoard[iPos].Active = true;
                TestBoard[iPos].ContractNo = textBoxContractNo.Text.ToUpper();
                TestBoard[iPos].MeterType = comboBoxMeterType.Text;

                if ((checkBoxSaved.Checked == true) && (strLines != null) && (iSavedCount <= strLines.Length))
                {
                    TestBoard[iPos].OwnerNo = strLines[iSavedCount - 1];
                }
                else
                {
                    TestBoard[iPos].OwnerNo = strCharValue + (i64OwnerNum++).ToString("D" + (iOwnersNoLength - strCharValue.Length));

                    if (checkBoxChed.Checked == true)
                    {
                        TestBoard[iPos].OwnerNo = ChedCheckSum(TestBoard[iPos].OwnerNo);
                    }
                }

                if (TestBoard[iPos].OwnerNo.Length != iCheckLength)
                {
                    LogMessage(MSG.ERROR, "Invalid Owners No. - " + TestBoard[iPos].OwnerNo);
                    return;
                }

                if (comboBoxClient.Text.Length >= 3)
                {
                    TestBoard[iPos].Client = comboBoxClient.Text;
                    TestBoard[iPos].ClientNo = iClientNo.ToString();
                }

                if (comboBoxFirmware.Text.Length >= 10)
                {
                    TestBoard[iPos].Firmware = comboBoxFirmware.Text;
                }

                TestBoard[iPos].MSN = "MSN?";
                TestBoard[iPos].MAC = "MAC?";
                TestBoard[iPos].DateTime = "DateTime?";

                //TestBoard[iPos].MSN = "U34NFC491" + (iPos).ToString("D6") + "L";
                //TestBoard[iPos].MAC = Random.Next(0, 16777215).ToString("X6") + Random.Next(0, 16777215).ToString("X6");
                //TestBoard[iPos].MAC = Random.Next(0, 16777215).ToString("X6") + (iPos).ToString("D6");

                if (comboBoxCustomerProgram.Text.Length >= 10)
                {
                    TestBoard[iPos].Program = comboBoxCustomerProgram.Text;
                }

                if (comboBoxRippleProgram.Text.Length >= 10)
                {
                    TestBoard[iPos].Ripple = comboBoxRippleProgram.Text;
                }

                iSavedCount++;
            }

            // Check for duplicate Owner numbers
            for (int iPos1 = 1; iPos1 <= Bench.numPosition; iPos1++)
            {
                if (!TestBoard[iPos1].Active)
                {
                    continue;
                }

                for (int iPos2 = 1; iPos2 <= Bench.numPosition; iPos2++)
                {
                    if (!TestBoard[iPos2].Active)
                    {
                        continue;
                    }

                    if (iPos1 != iPos2)
                    {
                        if (TestBoard[iPos1].OwnerNo == TestBoard[iPos2].OwnerNo)
                        {
                            LogMessage(MSG.ERROR, "Duplicate Owners No. - " + TestBoard[iPos1].OwnerNo);
                            return;
                        }
                    }
                }
            }

            UpdateGridViewBatch();
        }

        private static byte[] GetHash(string strInput)
        {
            HashAlgorithm algorithm = SHA1.Create();
            return algorithm.ComputeHash(Encoding.UTF8.GetBytes(strInput));
        }

        private static string GetHashString(string strInput)
        {
            StringBuilder sb = new StringBuilder();

            foreach (byte b in GetHash(strInput))
            {
                sb.Append(b.ToString("X2"));
            }

            return sb.ToString();
        }

        private void button_Fill_Click(object sender, EventArgs e)
        {
            BatchClear(1, Bench.numPosition);
            BatchAdd(1, Bench.numPosition);
        }

        private void Button_Clear_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_Clear_Click()");

            BatchClear(1, Bench.numPosition);
        }

        private void BatchClear(int iFirst, int iLast)
        {
            LogMessage(MSG.DEBUG, "BatchClear(" + iFirst + ", " + iLast + ")");

            for (int iPos = iFirst; iPos <= iLast; iPos++)
            {
                TestBoard[iPos].Active = false;
                TestBoard[iPos].Saved = false;
                TestBoard[iPos].Failed = 0;
                TestBoard[iPos].Status = "Status?";
                TestBoard[iPos].MeterType = "MeterType?";
                TestBoard[iPos].MSN = "MSN?";
                TestBoard[iPos].OwnerNo = "OwnerNo?";
                TestBoard[iPos].ContractNo = "ContractNo?";
                TestBoard[iPos].Client = "Client?";
                TestBoard[iPos].ClientNo = "ClientNo?";
                TestBoard[iPos].Firmware = "Firmware?";
                TestBoard[iPos].MAC = "MAC?";
                TestBoard[iPos].DateTime = "DateTime?";
                TestBoard[iPos].Program = "Program?";
                TestBoard[iPos].Ripple = "Ripple?";

                for (int j = 0; j < iMaxSteps; j++)
                {
                    TestBoard[iPos].ErrorV[j] = dDEFAULT;
                    TestBoard[iPos].ErrorR[j] = 0;
                }
            }

            UpdateGridViewBatch();
        }

        private void ClearMeterResults()
        {
            LogMessage(MSG.DEBUG, "ClearMeterResults()");

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                TestBoard[iPos].Saved = false;
                TestBoard[iPos].Failed = 0;
                TestBoard[iPos].Status = "Status?";

                for (int j = 0; j < iMaxSteps; j++)
                {
                    TestBoard[iPos].ErrorV[j] = dDEFAULT;
                    TestBoard[iPos].ErrorR[j] = 0;
                }
            }
        }

        private void ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem_Min.CheckState = CheckState.Unchecked;
            ToolStripMenuItem_Nor.CheckState = CheckState.Unchecked;

            ((ToolStripMenuItem)sender).CheckState = CheckState.Checked;
        }

        private void ToolStripMenuItem_RunningMode_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem_Continuous.CheckState = CheckState.Unchecked;
            ToolStripMenuItem_SingleStep.CheckState = CheckState.Unchecked;

            ((ToolStripMenuItem)sender).CheckState = CheckState.Checked;
        }

        private void ToolStripMenuItem_MeterClock_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem_Allow10s.CheckState = CheckState.Unchecked;
            ToolStripMenuItem_Allow60s.CheckState = CheckState.Unchecked;
            ToolStripMenuItem_DontCheck.CheckState = CheckState.Unchecked;

            ((ToolStripMenuItem)sender).CheckState = CheckState.Checked;
        }

        private void ButtonsEnabled(bool state)
        {
            LogMessage(MSG.DEBUG, "ButtonsEnabled(" + state + ")");

            buttonStart.Enabled = state;
            buttonSave.Enabled = state;
            buttonSaveCsv.Enabled = state;
            buttonClearAll.Enabled = state;
            buttonClearLogs.Enabled = state;
            buttonClearList.Enabled = state;

            buttonAdd.Enabled = state;
            buttonRemove.Enabled = state;
            buttonFill.Enabled = state;
            buttonClear.Enabled = state;
            buttonDelete.Enabled = state;

            PowerControls(state);
        }

        private int SleepDoEvents(DoWorkEventArgs e, bool bCancel, int iMs)
        {
            LogMessage(MSG.INFO, "Wait " + iMs + " ms ...");

            iMs = iMs / Settings.iSleepDivide;

            while (iMs > 0)
            {
                Thread.Sleep(100);
                iMs -= 100;

                if (BgWorker.CancellationPending && bCancel)
                {
                    e.Cancel = true;
                    return 1;
                }
            }

            return 0;
        }

        private void KillProcess(string strProcessName)
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName(strProcessName))
                {
                    LogMessage(MSG.INFO, "Process.Kill(" + strProcessName + ")");
                    p.Kill();
                }
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
            }
        }

        private void StartProcess(string strProcessName)
        {
            ProcessStartInfo p = new ProcessStartInfo(Settings.strRootPath + @"\" + strProcessName);
            p.WindowStyle = ProcessWindowStyle.Hidden;

            // Debug mode to see shell window
            if (ToolStripMenuItem_Min.Checked)
            {
                p.WindowStyle = ProcessWindowStyle.Normal;
            }

            if (Settings.bRunCtrComm)
            {
                try
                {
                    LogMessage(MSG.INFO, "Process.Start(" + strProcessName + ")");
                    Process.Start(p);
                }
                catch (Exception ex)
                {
                    LogMessage(MSG.ERROR, ex.ToString());
                    return;
                }
            }
        }

        private bool CreateExportFile(int iStep)
        {
            LogMessage(MSG.DEBUG, "CreateExportFile()");

            string strFname = "Exp" + StepValues.PStepNo.ToString("D2") + ".txt";
            string strText = "[Name] = " + StepValues.Name + Environment.NewLine;

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    strText += "[Pos] = " + iPos + Environment.NewLine;
                    strText += "[OwnerNo] = \"" + TestBoard[iPos].OwnerNo + "\"" + Environment.NewLine;
                    strText += "[ManufacturerNo] = \"" + TestBoard[iPos].MSN + "\"" + Environment.NewLine;
                    strText += "[Total] = \"" + TestBoard[iPos].Status + "\"" + Environment.NewLine;
                    strText += "[Result.Error] = " + TestBoard[iPos].ErrorV[iStep] + "%" + Environment.NewLine;
                }
            }

            if (WriteToFile(strFname, strText))
            {
                return true;
            }

            LogMessage(MSG.INFO, "Export File : " + Settings.strRootPath + @"\" + strFname);

            return false;
        }

        private bool GetProcedure(string strProcedure)
        {
            DataSet DataSetPStep;

            LogMessage(MSG.DEBUG, "GetProcedure(" + strProcedure + ")");

            if (strProcedure != "")
            {
                DataSetPStep = GetTestProcedure(strProcedure);

                if (DataSetPStep != null)
                {
                    DGV_Procedure.DataSource = DataSetPStep.Tables[0];
                    return false;
                }
            }
            else
            {
                LogMessage(MSG.ERROR, "No Procedure defined");
            }

            return true;
        }

        private void CreateResultsColumns()
        {
            LogMessage(MSG.DEBUG, "CreateResultsColumns()");

            DGV_Results.Columns.Clear();

            // Test step name column
            DGV_Results.ColumnCount = 1;
            DGV_Results.Columns[0].Name = "Name";
            DGV_Results.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            DGV_Results.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            // Test step number column
            DGV_Results.ColumnCount = 2;
            DGV_Results.Columns[1].Name = "Step";
            DGV_Results.Columns[1].Width = 50;
            DGV_Results.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Progress bar column
            DataGridViewProgressColumn column = new DataGridViewProgressColumn();
            column.HeaderText = "Average Error";
            DGV_Results.Columns.Add(column);
            DGV_Results.Columns[2].Width = 150;

            DGV_Results.ColumnCount = Bench.numPosition + 3;

            int iPos = 1;

            // Meter position columns
            for (int i = 3; i < Bench.numPosition + 3; i++, iPos++)
            {
                DGV_Results.Columns[i].Name = iPos.ToString("D");
                DGV_Results.Columns[i].DefaultCellStyle.Format = "N2";
                DGV_Results.Columns[i].Width = 70;
            }

            // Columns for base values, meter type & procedure
            DataGridViewTextBoxColumn c1 = new DataGridViewTextBoxColumn();
            c1.HeaderText = "Ub";
            c1.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            c1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //c1.Visible = false;
            DGV_Results.Columns.Add(c1);

            DataGridViewTextBoxColumn c2 = new DataGridViewTextBoxColumn();
            c2.HeaderText = "Ib";
            c2.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            c2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //c2.Visible = false;
            DGV_Results.Columns.Add(c2);

            DataGridViewTextBoxColumn c3 = new DataGridViewTextBoxColumn();
            c3.HeaderText = "Im";
            c3.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            c3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //c3.Visible = false;
            DGV_Results.Columns.Add(c3);

            DataGridViewTextBoxColumn c4 = new DataGridViewTextBoxColumn();
            c4.HeaderText = "MeterType";
            c4.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            c4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //c4.Visible = false;
            DGV_Results.Columns.Add(c4);

            DataGridViewTextBoxColumn c5 = new DataGridViewTextBoxColumn();
            c5.HeaderText = "Run";
            c5.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            c5.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            //c5.Visible = false;
            DGV_Results.Columns.Add(c5);

            int m = 0;

            foreach (DataGridViewColumn col in DGV_Results.Columns)
            {
                if ((m >= 3) && (m < (Bench.numPosition + 3)))
                {
                    if (TestBoard[m - 2].Active)
                    {
                        col.Visible = true;
                    }
                    else
                    {
                        col.Visible = false;
                    }
                }

                m++;
            }
        }

        private void CreateResultsRows()
        {
            LogMessage(MSG.DEBUG, "CreateResultsRows()");

            int iResultsRow = 0;

            // Add a row for each storing test step
            for (int i = 0; i < DGV_ProcedureRun.Rows.Count; i++)
            {
                if ((int)ConvertToDouble(DGV_ProcedureRun.Rows[i].Cells[20].Value.ToString()) > 1)
                {
                    // Populate test step name from DGV_ProcedureRun table
                    string strTestName = DGV_ProcedureRun.Rows[i].Cells[2].Value.ToString();
                    object[] oRow = new object[] { strTestName };
                    DGV_Results.Rows.Add(oRow);

                    // Populate test step number
                    DGV_Results.Rows[iResultsRow].Cells[1].Value = DGV_ProcedureRun.Rows[i].Cells[1].Value.ToString();

                    // Save base values, meter type & procedure for eMRS
                    DGV_Results.Rows[iResultsRow].Cells[Bench.numPosition + 3].Value = DGV_ProcedureRun.Rows[i].Cells[30].Value.ToString();
                    DGV_Results.Rows[iResultsRow].Cells[Bench.numPosition + 4].Value = DGV_ProcedureRun.Rows[i].Cells[31].Value.ToString();
                    DGV_Results.Rows[iResultsRow].Cells[Bench.numPosition + 5].Value = DGV_ProcedureRun.Rows[i].Cells[32].Value.ToString();
                    DGV_Results.Rows[iResultsRow].Cells[Bench.numPosition + 6].Value = DGV_ProcedureRun.Rows[i].Cells[33].Value.ToString();
                    DGV_Results.Rows[iResultsRow].Cells[Bench.numPosition + 7].Value = DGV_ProcedureRun.Rows[i].Cells[34].Value.ToString();

                    iResultsRow++;
                }
            }
        }

        private void UpdateResults(int iStep, int iResultsRow)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<int, int>(UpdateResults), new object[] { iStep, iResultsRow });
                return;
            }

            LogMessage(MSG.DEBUG, "UpdateResults()");

            // Create results DataGrid if needed
            if (DGV_Results.Columns.Count == 0)
            {
                CreateResultsColumns();
                CreateResultsRows();
            }

            int iPos = 1;
            int iCol = 3;
            double dVal = 0;
            int iTotalError = 0;

            // Don't rely on StepValues, worker thread may be one step ahead of GUI
            double dUlimit = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[17].Value.ToString());
            double dLlimit = ConvertToDouble(DGV_ProcedureRun.Rows[iStep].Cells[18].Value.ToString());

            for (iCol = 3; iCol < Bench.numPosition + 3; iCol++, iPos++)
            {
                // Next meter position
                if (!TestBoard[iPos].Active)
                {
                    continue;
                }

                dVal = TestBoard[iPos].ErrorV[iStep];

                if (dVal == iPASS)
                {
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Style.BackColor = Color.White;
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Value = "Pass";
                    TestBoard[iPos].ErrorR[iStep] = 0;
                }
                else if (dVal == iFAIL)
                {
                    LogMessage(MSG.DEBUG, "Position: " + iPos.ToString("D2") + ", Step: " + (iStep + 1).ToString("D2") + ", Value: " + "Fail");
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Style.BackColor = Color.Red;
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Value = "Fail";
                    TestBoard[iPos].ErrorR[iStep] = 1;
                }
                else if ((dVal >= dLlimit) && (dVal <= dUlimit))
                {
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Style.BackColor = Color.White;
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Value = dVal.ToString("F2") + "%";
                    TestBoard[iPos].ErrorR[iStep] = 0;
                    iTotalError += Math.Abs((int)(dVal * 100));
                }
                else
                {
                    LogMessage(MSG.DEBUG, "Position: " + iPos.ToString("D2") + ", Step: " + (iStep + 1).ToString("D2") + ", Value: " + dVal);
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Style.BackColor = Color.Red;
                    DGV_Results.Rows[iResultsRow].Cells[iCol].Value = ">>> " + dVal.ToString("F2") + "%";
                    TestBoard[iPos].ErrorR[iStep] = 1;
                    iTotalError += Math.Abs((int)(dVal * 100));
                }
            }

            UpdateTestBoard();

            // Average of errors
            DGV_Results.Rows[iResultsRow].Cells[2].Value = iTotalError / iMeters;
        }

        private DataSet GetTestProcedure(string strProcedure)
        {
            DataSet DataSetTP;
            DataSet DataSetPStep;

            LogMessage(MSG.DEBUG, "GetTestProcedure(" + strProcedure + ")");

            // Get ProcedureID
            if (Settings.strProcedurePC.Length == 10)
            {
                DataSetTP = GetMSAccessData(@"SELECT * FROM TestProcedure WHERE Name = " + "\"" + strProcedure + "\"" + @" ORDER BY Revision DESC", @"TestProcedure", @"\\" + Settings.strProcedurePC + @"\msc$\remo.ixf");
            }
            else
            {
                DataSetTP = GetMSAccessData(@"SELECT * FROM TestProcedure WHERE Name = " + "\"" + strProcedure + "\"" + @" ORDER BY Revision DESC", @"TestProcedure", Settings.strRootPath + @"\db\remo.ixf");
            }

            if (DataSetTP == null)
            { 
                return null;
            }

            if (DataSetTP.Tables["TestProcedure"].Rows.Count < 1)
            {
                LogMessage(MSG.ERROR, "Test procedure \"" + strProcedure + "\"" + " NOT found in database");
                return null;
            }

            DataRow dr = DataSetTP.Tables[0].Rows[0];
            string ProcedureID = dr["ProcedureID"].ToString();

            // Get procedure
            if (Settings.strProcedurePC.Length == 10)
            {
                DataSetPStep = GetMSAccessData("SELECT * FROM PStep WHERE ProcedureID = " + ProcedureID + " ORDER BY PStepNo", "PStep", @"\\" + Settings.strProcedurePC + @"\msc$\remo.ixf");
            }
            else
            {
                DataSetPStep = GetMSAccessData("SELECT * FROM PStep WHERE ProcedureID = " + ProcedureID + " ORDER BY PStepNo", "PStep", Settings.strRootPath + @"\db\remo.ixf");
            }

            return DataSetPStep;
        }

        private DataSet GetMSAccessData(string strAccessSelect, string strTable, string strDatabase)
        {
            LogMessage(MSG.DEBUG, "GetMSAccessData(" + strAccessSelect + ", " + strTable + ", " + strDatabase + ")");

            // Set Access connection and select strings
            string strAccessConn = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " + strDatabase;

            LogMessage(MSG.DEBUG, strAccessConn);

            // Create the dataset
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;

            try
            {
                myAccessConn = new OleDbConnection(strAccessConn);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, "Failed to create Database connection");
                LogMessage(MSG.ERROR, ex.Message);
                return null;
            }

            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(myDataSet, strTable);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, "Failed to retrieve data from the Database");
                LogMessage(MSG.ERROR, ex.Message);
                return null;
            }
            finally
            {
                myAccessConn.Close();
            }

            // A dataset can contain multiple tables, so let's get them all into an array.
            DataTableCollection dta = myDataSet.Tables;

            foreach (DataTable dt in dta)
            {
                LogMessage(MSG.DEBUG, "Found table : " + dt.TableName);
            }

            // The next two lines show two different ways you can get the count of tables in a dataset.
            //LogMessage(MSG.INFO, myDataSet.Tables.Count + " tables in data set");
            //LogMessage(MSG.INFO, dta.Count + " tables in data set");

            // The next several lines show how to get information on a specific table by name from the dataset.
            LogMessage(MSG.DEBUG, myDataSet.Tables[strTable].Rows.Count + " rows in " + strTable + " table");

            // The column info is automatically fetched from the database, so we can read it here.
            LogMessage(MSG.DEBUG, myDataSet.Tables[strTable].Columns.Count + " columns in " + strTable + " table");

            //DataColumnCollection drc = myDataSet.Tables[strTable].Columns;

            //int i = 0;

            //foreach (DataColumn dc in drc)
            //{
            // Print the column subscript, then the column's name and its data type.
            //LogMessage(MSG.INFO, "Column name[" + i++ + "]" + " is " + dc.ColumnName + " of type " + dc.DataType);
            //}

            //DataRowCollection dra = myDataSet.Tables[strTable].Rows;

            //foreach (DataRow dr in dra)
            //{
            // Print the ProcedureID as a subscript, then the Name.
            //LogMessage(MSG.INFO, "Column 1 = " + dr[0] + " Column 2 = " + dr[1]);
            //}

            return myDataSet;
        }

        private void Button_Add_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_Add_Click()");
            BatchModify(1);
        }

        private void Button_Remove_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_Remove_Click()");
            BatchModify(0);
        }

        private void BatchModify(int iAdd)
        {
            LogMessage(MSG.DEBUG, "BatchModify(" + iAdd + ")");

            string[] strWords = textBoxRange.Text.Split('.');

            if (string.IsNullOrEmpty(textBoxRange.Text))
            {
                LogMessage(MSG.ERROR, "No meter positions defined");
                return;
            }

            if (strWords.Length == 1)
            {
                int iPos = (int)ConvertToDouble(Regex.Replace(textBoxRange.Text, "[^0-9.]", ""));

                if ((iPos >= 1) && (iPos <= Bench.numPosition))
                {
                    if (iAdd == 1)
                    {
                        BatchAdd(iPos, iPos);
                    }
                    else
                    {
                        BatchClear(iPos, iPos);
                    }
                }

                return;
            }

            if (strWords.Length != 3)
            {
                return;
            }

            int iPosStart = (int)ConvertToDouble(Regex.Replace(strWords[0], "[^0-9.]", ""));
            int iPosEnd = (int)ConvertToDouble(Regex.Replace(strWords[2], "[^0-9.]", ""));

            if ((iPosStart >= 1) && (iPosEnd <= Bench.numPosition))
            {
                if (iAdd == 1)
                {
                    BatchAdd(iPosStart, iPosEnd);
                }
                else
                {
                    BatchClear(iPosStart, iPosEnd);
                }
            }
        }

        private void Button_All_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_All_Click()");

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                TestBoard[iPos].Active = true;
                TestBoard[iPos].Saved = false;
                TestBoard[iPos].Failed = 0;
                TestBoard[iPos].Status = "Status?";

                for (int j = 0; j < iMaxSteps; j++)
                {
                    TestBoard[iPos].ErrorV[j] = dDEFAULT;
                    TestBoard[iPos].ErrorR[j] = 0;
                }
            }

            BatchAdd(1, Bench.numPosition);
            UpdateGridViewBatch();
        }

        public void LogMessage(MSG type, string strMsg)
        {
            // To update GUI from worker thread
            if (InvokeRequired)
            {
                this.Invoke(new Action<MSG, string>(LogMessage), new object[] { type, strMsg });
                return;
            }

            string strMsgString = "\r\n" + DateTime.Now.ToString("HH:mm:ss.ff") + " - " + strMsg;

            switch (type)
            {
                case MSG.INFO:
                    if (Settings.bLogInfo)
                    {
                        textBoxMesg.AppendText(strMsgString);
                    }
                    Counters.iMsgInfo++;
                    break;
                case MSG.WARNING:
                    textBoxMesg.AppendText(strMsgString);
                    Counters.iMsgWarning++;
                    break;
                case MSG.ERROR:
                    textBoxMesg.AppendText(strMsgString);
                    Counters.iMsgErrors++;
                    break;
                case MSG.FAILURE:
                    textBoxMesg.AppendText(strMsgString);
                    Counters.iMsgFailure++;
                    break;
                case MSG.DEBUG:
                    if (Settings.bLogDebug)
                    {
                        textBoxMesg.AppendText(strMsgString);
                    }
                    Counters.iMsgDebug++;
                    break;
            }

            Application.DoEvents();
        }

        private void Button_ClearLogs_Click(object sender, EventArgs e)
        {
            if (BgWorker.IsBusy)
            {
                LogMessage(MSG.INFO, "Program is busy");
                return;
            }

            ClearMessageLogs();

            progressBar.Value = 0;

            textBoxContractNo.Clear();
            textBoxOwnersNo.Clear();
            comboBoxClient.Items.Clear();
            comboBoxMeterType.Items.Clear();
            comboBoxFirmware.Items.Clear();
            comboBoxCustomerProgram.Items.Clear();
            comboBoxRippleProgram.Items.Clear();

            listBoxProcedures.Items.Clear();
            listBox01to24.Items.Clear();
            listBox25to48.Items.Clear();

            DGV_ProcedureList.Rows.Clear();
            iProcedureListRow = 0;
            DGV_Batch.Columns.Clear();
            DGV_ProcedureRun.Columns.Clear();
            iProcedureRunRow = 0;
            DGV_Results.Columns.Clear();

            BatchClear(0, Bench.numPosition);
            bResultsSavedAccess = false;
            bResultsSavedSql = false;

            LogMessage(MSG.INFO, "Cleared ALL data");
        }

        private void ClearMessageLogs()
        {
            textBoxMesg.Clear();

            Counters.iMsgInfo = 0;
            Counters.iMsgWarning = 0;
            Counters.iMsgErrors = 0;
            Counters.iMsgFailure = 0;
            Counters.iMsgDebug = 0;

            Counters.iSetVoltage = 0;
            Counters.iSetVoltageErrors = 0;
            Counters.iSetCurrent = 0;
            Counters.iSetCurrentErrors = 0;
            Counters.iReadRefMeter = 0;
            Counters.iReadRefMeterErrors = 0;
            Counters.iSetErrorCounter = 0;
            Counters.iSetErrorCounterErrors = 0;
            Counters.iReadErrorCounter = 0;
            Counters.iReadErrorCounterErrors = 0;

            LogMessage(MSG.INFO, "Cleared message log");
        }

        private void Button_Save_Disk_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_Save_Click()");

            if (BgWorker.IsBusy)
            {
                LogMessage(MSG.INFO, "Program is busy");
                return;
            }

            ButtonsEnabled(false);
            SaveMeterResultsDisk();
            ButtonsEnabled(true);
        }

        private void Button_Save_Csv_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "Button_Save_Csv_Click()");

            if (BgWorker.IsBusy)
            {
                LogMessage(MSG.INFO, "Program is busy");
                return;
            }

            ButtonsEnabled(false);
            SaveMeterResultsCsv();
            ButtonsEnabled(true);
        }

        private void SaveMeterResultsDisk()
        {
            int RunID = 0;
            String strQuery = null;
            string strUb = "240";
            string strIb = "10";
            string strIm = "100";
            string strMeterType = "SP EM1000";
            string strRunName = "None";

            LogMessage(MSG.DEBUG, "SaveMeterResultsDisk()");

            if (DGV_Results.Rows.Count < 1)
            {
                LogMessage(MSG.INFO, "No results to save!");
                return;
            }

            if (!Settings.bSaveAccess && !Settings.bSaveSql)
            {
                LogMessage(MSG.WARNING, "Access or SQL storage options NOT set");
                return;
            }

            if (Settings.bSaveAccess && bResultsSavedAccess)
            {
                LogMessage(MSG.INFO, "Results already saved to MS Access");
            }

            if (Settings.bSaveAccess && !bResultsSavedAccess)
            {
                LogMessage(MSG.INFO, "Saving results to MS Access...");

                for (int iPos = 0; iPos <= Bench.numPosition; iPos++)
                {
                    TestBoard[iPos].Saved = false;
                }

                OleDbConnection OleDbConnection = new OleDbConnection();
                OleDbConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " + Settings.strRootPath + @"\db\AemCalData.mdb";

                strRunName = "None";

                // Save results. Run entry in database per procedure
                for (int iRow = 0; iRow < DGV_Results.Rows.Count; iRow++)
                {
                    // Update base values, meter type & run name for each procedure
                    if (strRunName != DGV_Results.Rows[iRow].Cells[Bench.numPosition + 7].Value.ToString())
                    {
                        strUb = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 3].Value.ToString();
                        strIb = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 4].Value.ToString();
                        strIm = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 5].Value.ToString();
                        strMeterType = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 6].Value.ToString();
                        strRunName = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 7].Value.ToString();

                        LogMessage(MSG.INFO, "Run Name - " + strRunName);

                        DateTime TimeRun = DateTime.Now;
                        int SupervisorID = 3;
                        int OperatorID = 2;
                        int MinTemp = 20;
                        int MaxTemp = 30;
                        int MinRH = 40;
                        int MaxRH = 80;
                        int Status = 3;

                        try
                        {
                            OleDbConnection.Open();
                            strQuery = "INSERT INTO Run (Name, TimeRun, SupervisorID, OperatorID, MinTemp, MaxTemp, MinRH, MaxRH, Status) VALUES('" + strRunName + "','" + TimeRun + "','" + SupervisorID + "','" + OperatorID + "','" + MinTemp + "','" + MaxTemp + "','" + MinRH + "','" + MaxRH + "','" + Status + "')";
                            LogMessage(MSG.DEBUG, strQuery);
                            OleDbCommand OleDbCommand = new OleDbCommand(strQuery, OleDbConnection);
                            OleDbCommand.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                        }
                        finally
                        {
                            OleDbConnection.Close();
                        }

                        int iLineType = Electricals.lineType;
                        int iConnectMode = Electricals.connectMode;
                        int iPrincipal = Electricals.principle;
                        string strChContent = "1,1,1000.000000";

                        try
                        {
                            OleDbConnection.Open();
                            OleDbCommand OleDbCommandRunID = new OleDbCommand("SELECT max(RunID) from Run", OleDbConnection);
                            RunID = (Int32)OleDbCommandRunID.ExecuteScalar();
                            strQuery = "INSERT INTO RMeterData (RunID, MeterName, LineType, ConnectMode, Principal, Ub, Ib, Imax, ChContent) VALUES('" + RunID + "','" + strMeterType + "','" + iLineType + "','" + iConnectMode + "','" + iPrincipal + "','" + ConvertToDouble(strUb) + "','" + ConvertToDouble(strIb) + "','" + ConvertToDouble(strIm) + "','" + strChContent + "')";
                            LogMessage(MSG.DEBUG, strQuery);
                            OleDbCommand OleDbCommand = new OleDbCommand(strQuery, OleDbConnection);
                            OleDbCommand.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                        }
                        finally
                        {
                            OleDbConnection.Close();
                        }

                        try
                        {
                            OleDbConnection.Open();

                            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                            {
                                if (TestBoard[iPos].Active)
                                {
                                    Status = TestBoard[iPos].Failed == 0 ? 1 : 2;
                                    String OwnerNo = TestBoard[iPos].OwnerNo;
                                    String MSN = TestBoard[iPos].MSN;
                                    int YearOfManufacture = Convert.ToInt32(DateTime.Now.ToString("yyyy"));
                                    String LastApproval = "";
                                    String ContractNo = TestBoard[iPos].ContractNo;
                                    String ClientName = TestBoard[iPos].Client;
                                    String ClientNo = TestBoard[iPos].ClientNo;

                                    strQuery = "INSERT INTO RMeter (RunID, PositionNo, Status, MeterName, OwnerNo, MSN, YearOfManufacture, LastApproval, ContractNo, ClientName, ClientNo) VALUES('" + RunID + "','" + iPos + "','" + Status + "','" + strMeterType + "','" + OwnerNo + "','" + MSN + "','" + YearOfManufacture + "','" + LastApproval + "','" + ContractNo + "','" + ClientName + "','" + ClientNo + "')";

                                    LogMessage(MSG.DEBUG, strQuery);

                                    OleDbCommand OleDbCommand = new OleDbCommand(strQuery, OleDbConnection);
                                    OleDbCommand.ExecuteNonQuery();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                        }
                        finally
                        {
                            OleDbConnection.Close();
                        }

                        try
                        {
                            OleDbConnection.Open();

                            for (int iRowStart = 0; (iRowStart + iRow) < DGV_Results.Rows.Count; iRowStart++)
                            {
                                if (strRunName == DGV_Results.Rows[iRowStart + iRow].Cells[Bench.numPosition + 7].Value.ToString())
                                {
                                    int iPos = 1;
                                    string strStepNo = DGV_Results.Rows[iRowStart + iRow].Cells[1].Value.ToString();

                                    for (int i = 3; i < (Bench.numPosition + 3); i++, iPos++)
                                    {
                                        if (!TestBoard[iPos].Active)
                                        {
                                            continue;
                                        }

                                        string strRValue;

                                        try
                                        {
                                            strRValue = DGV_Results.Rows[iRowStart + iRow].Cells[i].Value.ToString();
                                        }
                                        catch
                                        {
                                            strRValue = "";
                                        }

                                        strQuery = "INSERT INTO RResult (RunID, StepNo, PositionNo, RValue) VALUES('" + RunID + "','" + strStepNo + "','" + iPos + "','" + strRValue + "')";
                                        LogMessage(MSG.DEBUG, strQuery);
                                        OleDbCommand OleDbCommand = new OleDbCommand(strQuery, OleDbConnection);
                                        OleDbCommand.ExecuteNonQuery();
                                    }
                                }
                            }

                            bResultsSavedAccess = true;

                            // If we made it here results are in db
                            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                            {
                                if (TestBoard[iPos].Active)
                                {
                                    TestBoard[iPos].Saved = true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                        }
                        finally
                        {
                            OleDbConnection.Close();
                        }
                    }
                }

                LogMessage(MSG.INFO, "Finished");
            }

            if (Settings.bSaveSql && bResultsSavedSql)
            {
                LogMessage(MSG.INFO, "Results already saved to SQL");
            }

            if (Settings.bSaveSql && !bResultsSavedSql)
            {
                LogMessage(MSG.INFO, "Saving results to SQL...");

                for (int iPos = 0; iPos <= Bench.numPosition; iPos++)
                {
                    TestBoard[iPos].Saved = false;
                }

                strRunName = "None";

                // Save results. Run entry in database per procedure
                for (int iRow = 0; iRow < DGV_Results.Rows.Count; iRow++)
                {
                    // Update base values, meter type & run name for each procedure
                    if (strRunName != DGV_Results.Rows[iRow].Cells[Bench.numPosition + 7].Value.ToString())
                    {
                        strUb = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 3].Value.ToString();
                        strIb = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 4].Value.ToString();
                        strIm = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 5].Value.ToString();
                        strMeterType = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 6].Value.ToString();
                        strRunName = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 7].Value.ToString();

                        LogMessage(MSG.INFO, "Run Name - " + strRunName);

                        SqlConnection SqlConn = new SqlConnection(strSqlConnection);

                        string strTimeRun = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        int SupervisorID = 3;
                        int OperatorID = 2;
                        int MinTemp = 20;
                        int MaxTemp = 30;
                        int MinRH = 40;
                        int MaxRH = 80;
                        int Status = 3;

                        try
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Open()");
                            SqlConn.Open();
                            strQuery = "INSERT INTO Run (Name, TimeRun, SupervisorID, OperatorID, MinTemp, MaxTemp, MinRH, MaxRH, Status, StationID) VALUES('" + strRunName + "','" + strTimeRun + "','" + SupervisorID + "','" + OperatorID + "','" + MinTemp + "','" + MaxTemp + "','" + MinRH + "','" + MaxRH + "','" + Status + "','" + System.Environment.MachineName + "')";
                            LogMessage(MSG.DEBUG, strQuery);
                            SqlCommand SqlCommand = new SqlCommand(strQuery, SqlConn);
                            LogMessage(MSG.DEBUG, "SqlCommand.ExecuteNonQuery()");
                            SqlCommand.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, ex.ToString());
                            return;
                        }
                        finally
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Close()");
                            SqlConn.Close();
                        }

                        int iLineType = Electricals.lineType;
                        int iConnectMode = Electricals.connectMode;
                        int iPrincipal = Electricals.principle;
                        string strChContent = "1,1,1000.000000";

                        try
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Open()");
                            SqlConn.Open();
                            SqlCommand cmd = new SqlCommand("SELECT max(RunID) from Run", SqlConn);
                            RunID = (Int32)cmd.ExecuteScalar();
                            strQuery = "INSERT INTO RMeterData (RunID, MeterName, LineType, ConnectMode, Principal, Ub, Ib, Imax, ChContent) VALUES('" + RunID + "','" + strMeterType + "','" + iLineType + "','" + iConnectMode + "','" + iPrincipal + "','" + ConvertToDouble(strUb) + "','" + ConvertToDouble(strIb) + "','" + ConvertToDouble(strIm) + "','" + strChContent + "')";
                            LogMessage(MSG.DEBUG, strQuery);
                            SqlCommand SqlCommand = new SqlCommand(strQuery, SqlConn);
                            LogMessage(MSG.DEBUG, "SqlCommand.ExecuteNonQuery()");
                            SqlCommand.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                            return;
                        }
                        finally
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Close()");
                            SqlConn.Close();
                        }

                        try
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Open()");
                            SqlConn.Open();

                            // Create a DataTable to store table data
                            DataTable DataTableRMeter = new DataTable();

                            // Create the Columns for the DataTable
                            DataColumn c01 = new DataColumn("RunID");
                            DataColumn c02 = new DataColumn("PositionNo");
                            DataColumn c03 = new DataColumn("Status");
                            DataColumn c04 = new DataColumn("MeterName");
                            DataColumn c05 = new DataColumn("OwnerNo");
                            DataColumn c06 = new DataColumn("MSN");
                            DataColumn c07 = new DataColumn("YearOfManufacture");
                            DataColumn c08 = new DataColumn("LastApproval");
                            DataColumn c09 = new DataColumn("ContractNo");
                            DataColumn c10 = new DataColumn("ClientName");
                            DataColumn c11 = new DataColumn("ClientNo");

                            // Add Columns to the Datatable
                            DataTableRMeter.Columns.Add(c01);
                            DataTableRMeter.Columns.Add(c02);
                            DataTableRMeter.Columns.Add(c03);
                            DataTableRMeter.Columns.Add(c04);
                            DataTableRMeter.Columns.Add(c05);
                            DataTableRMeter.Columns.Add(c06);
                            DataTableRMeter.Columns.Add(c07);
                            DataTableRMeter.Columns.Add(c08);
                            DataTableRMeter.Columns.Add(c09);
                            DataTableRMeter.Columns.Add(c10);
                            DataTableRMeter.Columns.Add(c11);

                            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                            {
                                if (TestBoard[iPos].Active)
                                {
                                    DataRow DataRow = DataTableRMeter.NewRow();

                                    DataRow["RunID"] = RunID;
                                    DataRow["PositionNo"] = iPos;
                                    DataRow["Status"] = TestBoard[iPos].Failed == 0 ? 1 : 2;
                                    DataRow["MeterName"] = strMeterType;
                                    DataRow["OwnerNo"] = TestBoard[iPos].OwnerNo;
                                    DataRow["MSN"] = TestBoard[iPos].MSN;
                                    DataRow["YearOfManufacture"] = Convert.ToInt32(DateTime.Now.ToString("yyyy"));
                                    DataRow["LastApproval"] = "";
                                    DataRow["ContractNo"] = TestBoard[iPos].ContractNo;
                                    DataRow["ClientName"] = TestBoard[iPos].Client;
                                    DataRow["ClientNo"] = TestBoard[iPos].ClientNo;

                                    DataTableRMeter.Rows.Add(DataRow);
                                }
                            }

                            using (SqlBulkCopy SqlBulkCopy = new SqlBulkCopy(SqlConn))
                            {
                                SqlBulkCopy.DestinationTableName = "dbo.RMeter";

                                try
                                {
                                    LogMessage(MSG.INFO, "WriteToServer() - " + DataTableRMeter.Rows.Count + " records...");
                                    SqlBulkCopy.WriteToServer(DataTableRMeter);
                                }
                                catch (Exception ex)
                                {
                                    LogMessage(MSG.ERROR, ex.ToString());
                                    return;
                                }
                            }

                            DataTableRMeter.Clear();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                            return;
                        }
                        finally
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Close()");
                            SqlConn.Close();
                        }

                        try
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Open()");
                            SqlConn.Open();

                            // Create a DataTable to store table data
                            DataTable DataTableRResult = new DataTable();

                            // Create the Columns for the DataTable
                            DataColumn c01 = new DataColumn("RunID");
                            DataColumn c02 = new DataColumn("StepNo");
                            DataColumn c03 = new DataColumn("PositionNo");
                            DataColumn c04 = new DataColumn("RValue");

                            // Add Columns to the Datatable
                            DataTableRResult.Columns.Add(c01);
                            DataTableRResult.Columns.Add(c02);
                            DataTableRResult.Columns.Add(c03);
                            DataTableRResult.Columns.Add(c04);

                            for (int iRowStart = 0; (iRowStart + iRow) < DGV_Results.Rows.Count; iRowStart++)
                            {
                                if (strRunName == DGV_Results.Rows[iRowStart + iRow].Cells[Bench.numPosition + 7].Value.ToString())
                                {
                                    int iPos = 1;
                                    string strStepNo = DGV_Results.Rows[iRowStart + iRow].Cells[1].Value.ToString();

                                    for (int i = 3; i < (Bench.numPosition + 3); i++, iPos++)
                                    {
                                        if (!TestBoard[iPos].Active)
                                        {
                                            continue;
                                        }

                                        DataRow DataRow = DataTableRResult.NewRow();

                                        DataRow["RunID"] = RunID;
                                        DataRow["StepNo"] = strStepNo;
                                        DataRow["PositionNo"] = iPos;

                                        try
                                        {
                                            DataRow["RValue"] = DGV_Results.Rows[iRowStart + iRow].Cells[i].Value.ToString();
                                        }
                                        catch
                                        {
                                            DataRow["RValue"] = "";
                                        }

                                        DataTableRResult.Rows.Add(DataRow);
                                    }
                                }
                            }

                            using (SqlBulkCopy SqlBulkCopy = new SqlBulkCopy(SqlConn))
                            {
                                SqlBulkCopy.DestinationTableName = "dbo.RResult";

                                try
                                {
                                    LogMessage(MSG.INFO, "WriteToServer() - " + DataTableRResult.Rows.Count + " records...");
                                    SqlBulkCopy.WriteToServer(DataTableRResult);

                                    bResultsSavedSql = true;

                                    // If we made it here results are in db
                                    for (int iPos=1; iPos<=Bench.numPosition; iPos++)
                                    {
                                        if (TestBoard[iPos].Active)
                                        {
                                            TestBoard[iPos].Saved = true;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage(MSG.ERROR, ex.ToString());
                                    return;
                                }
                            }

                            DataTableRResult.Clear();
                        }
                        catch (Exception ex)
                        {
                            LogMessage(MSG.ERROR, "Failed - " + strQuery);
                            LogMessage(MSG.ERROR, ex.Message);
                            return;
                        }
                        finally
                        {
                            LogMessage(MSG.DEBUG, "SqlConnection.Close()");
                            SqlConn.Close();
                        }
                    }
                }

                LogMessage(MSG.INFO, "Finished");
            }

            UpdateTestBoard();
        }

        private void SaveMeterResultsCsv()
        {
            string strRunName = "None";

            LogMessage(MSG.DEBUG, "SaveMeterResultsCsv()");

            if (DGV_Results.Rows.Count < 1)
            {
                LogMessage(MSG.INFO, "No results to save!");
                return;
            }

            if (!Settings.bSaveCsv)
            {
                LogMessage(MSG.WARNING, "CSV storage options NOT set");
                return;
            }

            LogMessage(MSG.INFO, "Saving results to CSV...");

            // Save results. CSV file per procedure
            for (int iRow = 0; iRow < DGV_Results.Rows.Count; iRow++)
            {
                // Update base values, meter type & run name for each procedure
                if (strRunName != DGV_Results.Rows[iRow].Cells[Bench.numPosition + 7].Value.ToString())
                {
                    strRunName = DGV_Results.Rows[iRow].Cells[Bench.numPosition + 7].Value.ToString();

                    LogMessage(MSG.INFO, "Run Name - " + strRunName);

                    StringBuilder sb = new StringBuilder();
                    string strDirectory = @"C:\Temp\";
                    string strResultsFname = strDirectory + strRunName + ".csv";

                    if (!Directory.Exists(strDirectory))
                    {
                        Directory.CreateDirectory(strDirectory);
                        LogMessage(MSG.INFO, "Directory.CreateDirectory(" + strDirectory + ")");
                    }

                    StreamWriter writer = new StreamWriter(strResultsFname);

                    sb.Clear();
                    sb.Append(",Position,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(iPos + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Meter Status,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            string strFailed = TestBoard[iPos].Failed > 0 ? "-" : "+";
                            sb.Append(strFailed + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Meter Type,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(TestBoard[iPos].MeterType + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Meter Serial No.,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(TestBoard[iPos].MSN + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Owner No.,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(TestBoard[iPos].OwnerNo + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Year of Manufacture,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(DateTime.Now.ToString("yyyy") + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Contract No.,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(TestBoard[iPos].ContractNo + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Client,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(TestBoard[iPos].Client + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    sb.Clear();
                    sb.Append(",Client No.,");

                    for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
                    {
                        if (TestBoard[iPos].Active)
                        {
                            sb.Append(TestBoard[iPos].ClientNo + ",");
                        }
                    }

                    LogMessage(MSG.INFO, sb.ToString());
                    sb.Append(Environment.NewLine);
                    writer.Write(sb.ToString());

                    for (int iRowStart = 0; (iRowStart + iRow) < DGV_Results.Rows.Count; iRowStart++)
                    {
                        if (strRunName == DGV_Results.Rows[iRowStart + iRow].Cells[Bench.numPosition + 7].Value.ToString())
                        {
                            sb.Clear();
                            sb.Append(",");

                            for (int i = 0; i <= Bench.numPosition + 2; i++)
                            {
                                // Skip columns
                                if ((i == 1) || (i == 2) || ((i >= 3) && (!TestBoard[i - 2].Active)))
                                {
                                    continue;
                                }

                                try
                                {
                                    string strRValue = DGV_Results.Rows[iRowStart + iRow].Cells[i].Value.ToString();
                                    sb.Append(strRValue + ",");
                                }
                                catch
                                {
                                    sb.Append(",");
                                }
                            }

                            LogMessage(MSG.INFO, sb.ToString());
                            sb.Append(Environment.NewLine);
                            writer.Write(sb.ToString());
                        }
                    }

                    writer.Close();

                    if (File.Exists(strResultsFname))
                    {
                        LogMessage(MSG.INFO, "File created - " + strResultsFname);
                    }
                    else
                    {
                        LogMessage(MSG.ERROR, "ERROR creating - " + strResultsFname);
                    }
                }
            }

            LogMessage(MSG.INFO, "Finished");
        }

        private void ReadMeterClock(int iStep)
        {
            LogMessage(MSG.DEBUG, "ReadMeterClock()");

            for (int iPos = 1; iPos <= Bench.numPosition; iPos++)
            {
                if (TestBoard[iPos].Active)
                {
                    string strFname = "MeterTime." + iPos.ToString("D2");

                    if (ToolStripMenuItem_DontCheck.Checked)
                    {
                        TestBoard[iPos].ErrorV[iStep] = 0;
                        continue;
                    }

                    if (File.Exists(strFname) == true)
                    {
                        string[] strLines = File.ReadAllLines(strFname);

                        foreach (string strLine in strLines)
                        {
                            TestBoard[iPos].DateTime = strLine;
                        }

                        File.Delete(strFname);

                        //CheckMeterClock(iPos);
                    }
                    else
                    {
                        LogMessage(MSG.DEBUG, "File NOT found: " + strFname);
                    }
                }
            }
        }

        private bool ReadIniFile()
        {
            string strVal;

            // Loads .ini file from executable directory
            var IniFile = new IniFile();

            strVal = IniFile.Read("Log.Debug", "Debug");

            if (ConvertToDouble(strVal) == 1)
            {
                Settings.bLogDebug = true;
            }

            LogMessage(MSG.DEBUG, "Log.Debug : " + Settings.bLogDebug);

            strVal = IniFile.Read("Log.Info", "Debug");

            if (ConvertToDouble(strVal) == 0)
            {
                Settings.bLogInfo = false;
            }

            LogMessage(MSG.DEBUG, "Log.Info : " + Settings.bLogInfo);

            strVal = IniFile.Read("Sleep.Divide", "Debug");

            Settings.iSleepDivide = (int)ConvertToDouble(strVal);

            if (Settings.iSleepDivide != 1)
            {
                LogMessage(MSG.WARNING, "Sleep.Divide : " + Settings.iSleepDivide);
            }
            else
            {
                LogMessage(MSG.DEBUG, "Sleep.Divide : " + Settings.iSleepDivide);
            }

            strVal = IniFile.Read("Errors.Generate", "Debug");

            if (ConvertToDouble(strVal) == 1)
            {
                Settings.bErrorsGenerate = true;
            }

            LogMessage(MSG.DEBUG, "Errors.Generate : " + Settings.bErrorsGenerate);

            strVal = IniFile.Read("Errors.Type", "Debug");

            if (ConvertToDouble(strVal) == 1)
            {
                Settings.bErrorsType = true;
            }

            LogMessage(MSG.DEBUG, "Errors.Type : " + Settings.bErrorsType);

            strVal = IniFile.Read("Errors.Range", "Debug");

            Settings.dErrorsRange = ConvertToDouble(strVal);

            LogMessage(MSG.DEBUG, "Errors.Range : " + Settings.dErrorsRange);

            strVal = IniFile.Read("Run.Cmds", "Program");

            if (ConvertToDouble(strVal) == 0)
            {
                Settings.bRunCmds = false;
            }

            LogMessage(MSG.DEBUG, "Run.Cmds : " + Settings.bRunCmds);

            strVal = IniFile.Read("Run.CtrComm", "Program");

            if (ConvertToDouble(strVal) == 0)
            {
                Settings.bRunCtrComm = false;
            }

            LogMessage(MSG.DEBUG, "Run.CtrComm : " + Settings.bRunCtrComm);

            strVal = IniFile.Read("Files.Exist", "Program");

            if (ConvertToDouble(strVal) == 0)
            {
                Settings.bFilesExist = false;
            }

            LogMessage(MSG.DEBUG, "Files.Exist : " + Settings.bFilesExist);

            strVal = IniFile.Read("Process.ExitCode", "Program");

            if (ConvertToDouble(strVal) == 0)
            {
                Settings.bProcessExitCode = false;
            }

            LogMessage(MSG.DEBUG, "Process.ExitCode : " + Settings.bProcessExitCode);

            strVal = IniFile.Read("Phase", "Board");

            Bench.numPhase = (int)ConvertToDouble(strVal);

            LogMessage(MSG.DEBUG, "Phase : " + Bench.numPhase);

            strVal = IniFile.Read("Positions", "Board");

            int iPos = (int)ConvertToDouble(strVal);

            if ((iPos == 20) || (iPos == 24)||(iPos == 32)||(iPos == 48))
            {
                Bench.numPosition = iPos;
                LogMessage(MSG.DEBUG, "Positions : " + Bench.numPosition);
            }
            else
            {
                Bench.numPosition = 20;
                LogMessage(MSG.ERROR, "Invalid positions : " + iPos);
            }

            strVal = IniFile.Read("Store.Access", "Results");

            if (ConvertToDouble(strVal) == 1)
            {
                Settings.bSaveAccess = true;
            }

            LogMessage(MSG.DEBUG, "Store.Access : " + Settings.bSaveAccess);

            strVal = IniFile.Read("Store.Sql", "Results");

            if (ConvertToDouble(strVal) == 1)
            {
                Settings.bSaveSql = true;
            }

            LogMessage(MSG.DEBUG, "Store.Sql : " + Settings.bSaveSql);

            strVal = IniFile.Read("Store.Csv", "Results");

            if (ConvertToDouble(strVal) == 1)
            {
                Settings.bSaveCsv = true;
            }

            LogMessage(MSG.DEBUG, "Store.Csv : " + Settings.bSaveCsv);

            strVal = IniFile.Read("Root.Path", "PC");

            if (Directory.Exists(strVal))
            {
                Settings.strRootPath = strVal;
                LogMessage(MSG.DEBUG, "Root.Path : " + Settings.strRootPath);
            }
            else
            {
                LogMessage(MSG.ERROR, "Invalid path : " + strVal);
            }

            strVal = IniFile.Read("Procedure.PC", "PC");

            if (strVal.Length == 10)
            {
                Settings.strProcedurePC = strVal;
                LogMessage(MSG.DEBUG, "Procedure.PC : " + Settings.strProcedurePC);
            }
            else
            {
                LogMessage(MSG.ERROR, "Invalid PC name : " + strVal);
            }

            strVal = IniFile.Read("SQL.Connection", "SQL");

            if (strVal.Length > 50)
            {
                strSqlConnection = strVal;
                LogMessage(MSG.DEBUG, "SQL.Connection : " + strSqlConnection);
            }
            else
            {
                LogMessage(MSG.ERROR, "Invalid SQL Connection : " + strVal);
            }

            return false;
        }

        private bool ReadImportFile(int iStep, int iImport2)
        {
            int iPos = 1;
            string strFname = Settings.strRootPath + @"\Imp.res";
            string[] strLines;
            int iTagsFound = 0;

            // Generate errors if not executing commands
            if (!Settings.bRunCmds)
            {
                for (iPos = 1; iPos <= Bench.numPosition; iPos++)
                {
                    if (TestBoard[iPos].Active)
                    {
                        GenerateErrors(iStep, iPos);
                    }
                }

                return false;
            }

            if (File.Exists(strFname) == false)
            {
                LogMessage(MSG.WARNING, "File NOT found : " + strFname);
                return true;
            }
            else
            {
                LogMessage(MSG.INFO, "Import File : " + strFname);
            }

            try
            {
                strLines = File.ReadAllLines(strFname);
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
                return true;
            }

            foreach (string strLine in strLines)
            {
                string[] strLineSplit = strLine.Split('=');

                if (strLineSplit.Length == 2)
                {
                    string strTag = strLineSplit[0].Trim();
                    string strVal = strLineSplit[1].Trim();

                    if (String.Equals(strTag, "[Name]", StringComparison.OrdinalIgnoreCase))
                    {
                        // Do we need this?
                        continue;
                    }

                    if (String.Equals(strTag, "[Pos]", StringComparison.OrdinalIgnoreCase))
                    {
                        iPos = (int)ConvertToDouble(strVal);

                        if (iPos > Bench.numPosition)
                        {
                            LogMessage(MSG.ERROR, "File: " + strFname + ", invalid position: " + iPos);
                        }

                        continue;
                    }

                    if (String.Equals(strTag, "[SSNMACAddress]", StringComparison.OrdinalIgnoreCase))
                    {
                        iTagsFound++;
                        LogMessage(MSG.DEBUG, "[Pos] = " + iPos + " [SSNMACAddress] = " + strVal);
                        TestBoard[iPos].MAC = strVal.Trim('"');
                        continue;
                    }

                    if (String.Equals(strTag, "[OwnerNo]", StringComparison.OrdinalIgnoreCase))
                    {
                        iTagsFound++;
                        LogMessage(MSG.DEBUG, "[Pos] = " + iPos + " [OwnerNo] = " + strVal);
                        TestBoard[iPos].OwnerNo = strVal.Trim('"');
                        continue;
                    }

                    if (String.Equals(strTag, "[ManufacturerNo]", StringComparison.OrdinalIgnoreCase))
                    {
                        iTagsFound++;
                        LogMessage(MSG.DEBUG, "[Pos] = " + iPos + " [ManufacturerNo] = " + strVal);
                        TestBoard[iPos].MSN = strVal.Trim('"');
                        continue;
                    }

                    if (String.Equals(strTag, "[FirmwareVersion]", StringComparison.OrdinalIgnoreCase))
                    {
                        iTagsFound++;
                        LogMessage(MSG.DEBUG, "[Pos] = " + iPos + " [FirmwareVersion] = " + strVal);
                        TestBoard[iPos].Firmware = strVal.Trim('"');
                        continue;
                    }

                    if (String.Equals(strTag, "[Total]", StringComparison.OrdinalIgnoreCase))
                    {
                        iTagsFound++;
                        LogMessage(MSG.DEBUG, "[Pos] = " + iPos + " [Total] = " + strVal);

                        if (TestBoard[iPos].Status != "-")
                        {
                            TestBoard[iPos].Status = strVal.Trim('"');
                        }

                        continue;
                    }

                    if (String.Equals(strTag, "[Result.Error]", StringComparison.OrdinalIgnoreCase))
                    {
                        if ((StepValues.Storing != 1) && (iImport2 == 1))
                        {
                            LogMessage(MSG.DEBUG, "[Pos] = " + iPos + " [Result.Error] = " + strVal);

                            if (String.Equals(strVal, "PASS", StringComparison.OrdinalIgnoreCase))
                            {
                                TestBoard[iPos].ErrorV[iStep] = iPASS;
                            }
                            else if (String.Equals(strVal, "FAIL", StringComparison.OrdinalIgnoreCase))
                            {
                                TestBoard[iPos].ErrorV[iStep] = iFAIL;
                            }
                            else
                            {
                                TestBoard[iPos].ErrorV[iStep] = ConvertToDouble(strVal.Replace("%", ""));
                            }
                        }

                        continue;
                    }

                    LogMessage(MSG.WARNING, "Tag NOT found: " + strTag);
                }
            }

            // Only update if needed
            if (iTagsFound > 0)
            {
                UpdateGridViewBatch();
            }

            return false;
        }

        private void ComboBox_Customer_Dropdown(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_Customer_Dropdown()");

            DataSet DataSetCustomer = GetMSAccessData("SELECT DISTINCT ClientNo, Name FROM Client ORDER BY Name", "Client", Settings.strRootPath + @"\db\AemCalData.mdb");

            if (DataSetCustomer != null)
            {
                comboBoxClient.Items.Clear();

                for (int r = 0; r < DataSetCustomer.Tables[0].Rows.Count; r++)
                {
                    DataRow dr = DataSetCustomer.Tables[0].Rows[r];
                    string strClient = dr["Name"].ToString();
                    comboBoxClient.Items.Add(strClient);
                }
            }
        }

        private void ComboBox_Customer_SelectedIndexChanged(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_Customer_SelectedIndexChanged()");

            DataSet DataSetCustomer = GetMSAccessData("SELECT DISTINCT ClientNo, Name FROM Client ORDER BY Name", "Client", Settings.strRootPath + @"\db\AemCalData.mdb");

            if (DataSetCustomer != null)
            {
                for (int r = 0; r < DataSetCustomer.Tables[0].Rows.Count; r++)
                {
                    DataRow dr = DataSetCustomer.Tables[0].Rows[r];
                    string strClient = dr["Name"].ToString();

                    if (strClient == comboBoxClient.Text)
                    {
                        iClientNo = (int)ConvertToDouble(dr["ClientNo"].ToString());
                        LogMessage(MSG.DEBUG, "iClientNo = " + iClientNo);
                    }
                }
            }
        }

        private void ComboBox_MeterType_Dropdown(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_MeterType_Dropdown()");

            DataSet DataSetMeterType = GetMSAccessData("SELECT DISTINCT Name, Ub, Ib, Imax FROM MeterType ORDER BY Name", "MeterType", Settings.strRootPath + @"\db\AemCalData.mdb");

            if (DataSetMeterType != null)
            {
                comboBoxMeterType.Items.Clear();

                for (int r = 0; r < DataSetMeterType.Tables[0].Rows.Count; r++)
                {
                    DataRow dr = DataSetMeterType.Tables[0].Rows[r];
                    string meterType = dr["Name"].ToString();
                    comboBoxMeterType.Items.Add(meterType);
                }
            }
        }

        private void ComboBox_MeterType_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataSet DataSetTP;

            LogMessage(MSG.DEBUG, "ComboBox_MeterType_SelectedIndexChanged()");

            DataSet DataSetMeterType = GetMSAccessData("SELECT DISTINCT Name, LineType, ConnectMode, Principal, Ub, Ib, Imax FROM MeterType ORDER BY Name", "MeterType", Settings.strRootPath + @"\db\AemCalData.mdb");

            if (DataSetMeterType != null)
            {
                for (int r = 0; r < DataSetMeterType.Tables[0].Rows.Count; r++)
                {
                    DataRow dr = DataSetMeterType.Tables[0].Rows[r];
                    string meterType = dr["Name"].ToString();

                    if (meterType == comboBoxMeterType.Text)
                    {
                        Electricals.iUBase = (int)ConvertToDouble(dr["Ub"].ToString());
                        Electricals.iIBase = (int)ConvertToDouble(dr["Ib"].ToString());
                        Electricals.iIMax = (int)ConvertToDouble(dr["Imax"].ToString());
                        Electricals.lineType = (int)ConvertToDouble(dr["LineType"].ToString());
                        Electricals.connectMode = (int)ConvertToDouble(dr["ConnectMode"].ToString());
                        Electricals.principle = (int)ConvertToDouble(dr["Principal"].ToString());
                    }
                }
            }

            listBoxProcedures.Items.Clear();

            // Get list of procedures
            if (Settings.strProcedurePC.Length == 10)
            {
                DataSetTP = GetMSAccessData("SELECT DISTINCT Name FROM TestProcedure ORDER BY Name", "TestProcedure", @"\\" + Settings.strProcedurePC + @"\msc$\remo.ixf");
            }
            else
            {
                DataSetTP = GetMSAccessData("SELECT DISTINCT Name FROM TestProcedure ORDER BY Name", "TestProcedure", Settings.strRootPath + @"\db\remo.ixf");
            }

            if (DataSetTP != null)
            {
                string[] strMeterSeries1  = { "EM10", "EM12", "EM50", "EM51", "EM53", "EM54", "U12", "U13", "U330", "U335", "U34" };
                string[] strMeterSeries2  = { "E100", "E12",  "E50",  "E51",  "E53",  "E54",  "U12", "U13", "U330", "U335", "U34" };
                string[] strMeterPlatform = { "EM",   "EM",   "EM",   "EM",   "EM",   "EM",   "U",   "U",   "U",    "U",    "U" };

                int i = 0;

                if (checkBoxViewAll.Checked == false)
                {
                    for (i = 0; i < strMeterSeries1.Length; i++)
                    {
                        if (Regex.IsMatch(comboBoxMeterType.Text, strMeterSeries1[i]))
                        {
                            break;
                        }
                    }

                    if (i >= strMeterSeries1.Length)
                    {
                        LogMessage(MSG.WARNING, "No meter series found from meter type selected!");
                        return;
                    }
                }

                for (int r = 0; r < DataSetTP.Tables[0].Rows.Count; r++)
                {
                    DataRow dr = DataSetTP.Tables[0].Rows[r];
                    string strProcName = dr["Name"].ToString();

                    if (checkBoxViewAll.Checked == true)
                    {
                        listBoxProcedures.Items.Add(strProcName);
                        continue;
                    }

                    if ((Regex.IsMatch(strProcName, strMeterSeries1[i])) || (Regex.IsMatch(strProcName, strMeterSeries2[i])))
                    {
                        listBoxProcedures.Items.Add(strProcName);
                        continue;
                    }

                    if (Regex.IsMatch(strProcName, "Customer Program"))
                    {
                        if (Regex.IsMatch(strProcName, strMeterPlatform[i]))
                        {
                            listBoxProcedures.Items.Add(strProcName);
                            continue;
                        }
                    }

                    if (strMeterPlatform[i] == "U")
                    {
                        if (Regex.IsMatch(strProcName, "SSN"))
                        {
                            listBoxProcedures.Items.Add(strProcName);
                            continue;
                        }
                    }
                }
            }
        }

        private void ComboBox_Firmware_Dropdown(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_Firmware_Dropdown()");

            if (string.IsNullOrEmpty(comboBoxMeterType.Text))
            {
                LogMessage(MSG.WARNING, "Select Meter Type");
                return;
            }

            comboBoxFirmware.Items.Clear();

            try
            {
                string[] strFileList = Directory.GetFiles(Settings.strRootPath);
                string[] strRegexPattern = { @"^S00237-05.05.", @"^S00244-05.05.", @"^S00237-05.05.", @"^S00237-05.05.", @"^S00261-05.05." };
                string[] strSeries = { "U12", "U13", "U330", "U335", "U340" };

                int i;

                for (i = 0; i < strSeries.Length; i++)
                {
                    if (Regex.IsMatch(comboBoxMeterType.Text, strSeries[i]))
                    {
                        break;
                    }
                }

                if (i >= strSeries.Length)
                {
                    LogMessage(MSG.WARNING, "No firmware for meter type : " + comboBoxMeterType.Text);
                    return;
                }

                foreach (string strFileName in strFileList)
                {
                    string fName = Path.GetFileName(strFileName);

                    if (Regex.IsMatch(fName, strRegexPattern[i]))
                    {
                        comboBoxFirmware.Items.Add(fName);
                    }
                }
            }
            catch
            {
                LogMessage(MSG.ERROR, "Invalid path : " + Settings.strRootPath);
            }
        }

        private void ComboBox_Firmware_SelectedIndexChanged(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_Firmware_SelectedIndexChanged()");

            string[] segmentsSeries = comboBoxMeterType.Text.Split(' ');
            string[] segmentsFW = comboBoxFirmware.Text.Split('.');
            string[] match = { "U12", "U13", "U330", "U340" };
            string[] series = { "U1200", "U1300", "U3300", "U3400" };

            int i;

            for (i = 0; i < match.Length; i++)
            {
                if (Regex.IsMatch(comboBoxMeterType.Text, match[i]))
                {
                    break;
                }
            }

            if (segmentsFW.Length == 4)
            {
                string strText = "";

                strText += "@echo off" + Environment.NewLine;
                strText += "set fname=" + comboBoxFirmware.Text + Environment.NewLine;
                strText += "set rev=05.05." + segmentsFW[2] + Environment.NewLine;

                if (i < series.Length)
                {
                    strText += "set config=" + series[i] + "_defaultconfig.xml" + Environment.NewLine;
                }

                if (WriteToFile("args.bat", strText))
                {
                    return;
                }
            }
        }

        private void ComboBox_CustomerProgram_Dropdown(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_CustomerProgram_Dropdown()");

            if (string.IsNullOrEmpty(comboBoxClient.Text))
            {
                LogMessage(MSG.ERROR, "Select Customer");
                return;
            }

            if (string.IsNullOrEmpty(comboBoxMeterType.Text))
            {
                LogMessage(MSG.WARNING, "Select Meter Type");
                return;
            }

            comboBoxCustomerProgram.Items.Clear();

            try
            {
                string[] directoryList = Directory.GetDirectories(Settings.strRootPath + @"\CfgRoot\Customers\" + (string)comboBoxClient.SelectedItem);
                string[] series = { "EM10", "EM12", "U12", "U13", "U330", "U335", "U340" };

                int i;

                for (i = 0; i < series.Length; i++)
                {
                    if (Regex.IsMatch(comboBoxMeterType.Text, series[i]))
                    {
                        break;
                    }
                }

                if (i >= series.Length)
                {
                    LogMessage(MSG.WARNING, "No Customer Programs for meter type : " + comboBoxMeterType.Text);
                    return;
                }

                foreach (string directoryName in directoryList)
                {
                    string dirName = Path.GetFileName(directoryName);

                    if (Regex.IsMatch(dirName, series[i]) && !Regex.IsMatch(dirName, "^RP"))
                    {
                        comboBoxCustomerProgram.Items.Add(dirName);
                    }
                }
            }
            catch
            {
                LogMessage(MSG.ERROR, "Invalid path : " + Settings.strRootPath + @"\CfgRoot\Customers\" + (string)comboBoxClient.SelectedItem);
            }
        }

        private void ComboBox_RippleProgram_Dropdown(object sender, EventArgs e)
        {
            LogMessage(MSG.DEBUG, "ComboBox_RippleProgram_Dropdown()");

            if (string.IsNullOrEmpty(comboBoxClient.Text))
            {
                LogMessage(MSG.ERROR, "Select Customer");
                return;
            }

            if (string.IsNullOrEmpty(comboBoxMeterType.Text))
            {
                LogMessage(MSG.WARNING, "Select Meter Type");
                return;
            }

            comboBoxRippleProgram.Items.Clear();

            try
            {
                string[] fileList = Directory.GetFiles(Settings.strRootPath + @"\CfgRoot\Customers\" + (string)comboBoxClient.SelectedItem);
                string[] series = { "EM10", "EM12", "U12", "U13", "U330", "U335", "U340" };

                int i;

                for (i = 0; i < series.Length; i++)
                {
                    if (Regex.IsMatch(comboBoxMeterType.Text, series[i]))
                    {
                        break;
                    }
                }

                if (i >= series.Length)
                {
                    LogMessage(MSG.WARNING, "No Ripple Programs for meter type : " + comboBoxMeterType.Text);
                    return;
                }

                foreach (string fileName in fileList)
                {
                    string fName = Path.GetFileName(fileName);

                    if (Regex.IsMatch(fName, series[i]) && Regex.IsMatch(fName, "^RPG"))
                    {
                        comboBoxRippleProgram.Items.Add(fName);
                    }
                }
            }
            catch
            {
                LogMessage(MSG.ERROR, "Invalid path : " + Settings.strRootPath + @"\CfgRoot\Customers\" + (string)comboBoxClient.SelectedItem);
            }
        }

        private void Button_Step_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.INFO, "*** STEP ***");
            bClickStep = true;
        }

        private void Button_Stop_Click(object sender, EventArgs e)
        {
            LogMessage(MSG.INFO, "*** STOP ***");

            if (BgWorker.IsBusy)
            {
                BgWorker.CancelAsync();
            }
        }

        private void UpdateStats()
        {
            using (var uptime = new PerformanceCounter("System", "System Up Time"))
            {
                TimeSpan.FromSeconds(uptime.NextValue());
                textBoxSystemUpTime.Text = TimeSpan.FromSeconds(uptime.NextValue()).ToString(@"dd\.hh\:mm\:ss");
            }

            textBoxAppRunTime.Text = DateTime.Now.Subtract(AppStartTime).ToString(@"dd\.hh\:mm\:ss");

            textBoxInfoCount.Text = Counters.iMsgInfo.ToString();
            textBoxWarningCount.Text = Counters.iMsgWarning.ToString();
            textBoxErrorCount.Text = Counters.iMsgErrors.ToString();
            textBoxFailureCount.Text = Counters.iMsgFailure.ToString();
            textBoxDebugCount.Text = Counters.iMsgDebug.ToString();
            textBoxSetVoltage.Text = Counters.iSetVoltage.ToString();
            textBoxSetVoltageErrors.Text = Counters.iSetVoltageErrors.ToString();
            textBoxSetCurrent.Text = Counters.iSetCurrent.ToString();
            textBoxSetCurrentErrors.Text = Counters.iSetCurrentErrors.ToString();
            textBoxReadReference.Text = Counters.iReadRefMeter.ToString();
            textBoxReadReferenceErrors.Text = Counters.iReadRefMeterErrors.ToString();
            textBoxSetErrorCounter.Text = Counters.iSetErrorCounter.ToString();
            textBoxSetErrorCounterErrors.Text = Counters.iSetErrorCounterErrors.ToString();
            textBoxReadErrorCounter.Text = Counters.iReadErrorCounter.ToString();
            textBoxReadErrorCounterErrors.Text = Counters.iReadErrorCounterErrors.ToString();
        }

        public DateTime GetWindowsInstallationDateTime(string computerName)
        {
            LogMessage(MSG.DEBUG, "GetWindowsInstallationDateTime(" + computerName + ")");

            Microsoft.Win32.RegistryKey key = Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, computerName);
            key = key.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false);

            if (key != null)
            {
                DateTime startDate = new DateTime(1970, 1, 1, 0, 0, 0);
                Int64 regVal = Convert.ToInt64(key.GetValue("InstallDate").ToString());
                DateTime installDate = startDate.AddSeconds(regVal);

                return installDate;
            }

            return DateTime.MinValue;
        }

        public int GetStationID()
        {
            LogMessage(MSG.DEBUG, "GetStationID()");

            Microsoft.Win32.RegistryKey key = Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, "");
            key = key.OpenSubKey(@"SOFTWARE\EMH\Bench", false);

            if (key != null)
            {
                return Convert.ToInt32(key.GetValue("StationID").ToString());
            }

            return 0;
        }

        private void Button_VAll_Click(object sender, EventArgs e)
        {
            if (button_VAll.Text == "240")
            {
                button_VAll.Text = button_VA.Text = button_VB.Text = button_VC.Text = "Off";
                textBox_VA.Text = textBox_VB.Text = textBox_VC.Text = "240";
            }
            else
            {
                button_VAll.Text = button_VA.Text = button_VB.Text = button_VC.Text = "240";
                textBox_VA.Text = textBox_VB.Text = textBox_VC.Text = "0";
            }
        }

        private void ButtonA_All_Click(object sender, EventArgs e)
        {
            if (button_IAll.Text == "10")
            {
                button_IAll.Text = button_IA.Text = button_IB.Text = button_IC.Text = "Off";
                textBox_IA.Text = textBox_IB.Text = textBox_IC.Text = "10";
            }
            else
            {
                button_IAll.Text = button_IA.Text = button_IB.Text = button_IC.Text = "10";
                textBox_IA.Text = textBox_IB.Text = textBox_IC.Text = "0";
            }
        }

        private void ButtonPhi_All_Click(object sender, EventArgs e)
        {
            if (button_PHIAll.Text == "60")
            {
                button_PHIAll.Text = button_PHIA.Text = button_PHIB.Text = button_PHIC.Text = "0";
                textBox_PHIA.Text = textBox_PHIB.Text = textBox_PHIC.Text = "60";
            }
            else
            {
                button_PHIAll.Text = button_PHIA.Text = button_PHIB.Text = button_PHIC.Text = "60";
                textBox_PHIA.Text = textBox_PHIB.Text = textBox_PHIC.Text = "0";
            }
        }

        private void ButtonV_A_Click(object sender, EventArgs e)
        {
            if (button_VA.Text == "240")
            {
                button_VA.Text = "Off";
                textBox_VA.Text = "240";
            }
            else
            {
                button_VA.Text = "240";
                textBox_VA.Text = "0";
            }
        }

        private void ButtonV_B_Click(object sender, EventArgs e)
        {
            if (button_VB.Text == "240")
            {
                button_VB.Text = "Off";
                textBox_VB.Text = "240";
            }
            else
            {
                button_VB.Text = "240";
                textBox_VB.Text = "0";
            }
        }

        private void ButtonV_C_Click(object sender, EventArgs e)
        {
            if (button_VC.Text == "240")
            {
                button_VC.Text = "Off";
                textBox_VC.Text = "240";
            }
            else
            {
                button_VC.Text = "240";
                textBox_VC.Text = "0";
            }
        }

        private void ButtonA_A_Click(object sender, EventArgs e)
        {
            if (button_IA.Text == "10")
            {
                button_IA.Text = "Off";
                textBox_IA.Text = "10";
            }
            else
            {
                button_IA.Text = "10";
                textBox_IA.Text = "0";
            }
        }

        private void ButtonA_B_Click(object sender, EventArgs e)
        {
            if (button_IB.Text == "10")
            {
                button_IB.Text = "Off";
                textBox_IB.Text = "10";
            }
            else
            {
                button_IB.Text = "10";
                textBox_IB.Text = "0";
            }
        }

        private void ButtonA_C_Click(object sender, EventArgs e)
        {
            if (button_IC.Text == "10")
            {
                button_IC.Text = "Off";
                textBox_IC.Text = "10";
            }
            else
            {
                button_IC.Text = "10";
                textBox_IC.Text = "0";
            }
        }

        private void ButtonPhi_A_Click(object sender, EventArgs e)
        {
            if (button_PHIA.Text == "60")
            {
                button_PHIA.Text = "0";
                textBox_PHIA.Text = "60";
            }
            else
            {
                button_PHIA.Text = "60";
                textBox_PHIA.Text = "0";
            }
        }

        private void ButtonPhi_B_Click(object sender, EventArgs e)
        {
            if (button_PHIB.Text == "60")
            {
                button_PHIB.Text = "0";
                textBox_PHIB.Text = "60";
            }
            else
            {
                button_PHIB.Text = "60";
                textBox_PHIB.Text = "0";
            }
        }

        private void ButtonPhi_C_Click(object sender, EventArgs e)
        {
            if (button_PHIC.Text == "60")
            {
                button_PHIC.Text = "0";
                textBox_PHIC.Text = "60";
            }
            else
            {
                button_PHIC.Text = "60";
                textBox_PHIC.Text = "0";
            }
        }

        private void TextBoxAV_A_Changed(object sender, EventArgs e)
        {
            vProgBar_VA.Value = GetNumericalInput(textBox_AVA, 0, 300);
        }

        private void TextBoxAV_B_Changed(object sender, EventArgs e)
        {
            vProgBar_VB.Value = GetNumericalInput(textBox_AVB, 0, 300);
        }

        private void TextBoxAV_C_Changed(object sender, EventArgs e)
        {
            vProgBar_VC.Value = GetNumericalInput(textBox_AVC, 0, 300);
        }

        private void TextBoxAA_A_Changed(object sender, EventArgs e)
        {
            vProgBar_IA.Value = GetNumericalInput(textBox_AIA, 0, 100);
        }

        private void TextBoxAA_B_Changed(object sender, EventArgs e)
        {
            vProgBar_IB.Value = GetNumericalInput(textBox_AIB, 0, 100);
        }

        private void TextBoxAA_C_Changed(object sender, EventArgs e)
        {
            vProgBar_IC.Value = GetNumericalInput(textBox_AIC, 0, 100);
        }

        private void TextBoxAPhi_A_Changed(object sender, EventArgs e)
        {
            vProgBar_PHIA.Value = GetNumericalInput(textBox_APHIA, 0, 400);
        }

        private void TextBoxAPhi_B_Changed(object sender, EventArgs e)
        {
            vProgBar_PHIB.Value = GetNumericalInput(textBox_APHIB, 0, 400);
        }

        private void TextBoxAPhi_C_Changed(object sender, EventArgs e)
        {
            vProgBar_PHIC.Value = GetNumericalInput(textBox_APHIC, 0, 400);
        }

        private void TextBoxV_A_Changed(object sender, EventArgs e)
        {
            textBox_AVA.BackColor = Color.Orange;
        }

        private void TextBoxV_B_Changed(object sender, EventArgs e)
        {
            textBox_AVB.BackColor = Color.Orange;
        }

        private void TextBoxV_C_Changed(object sender, EventArgs e)
        {
            textBox_AVC.BackColor = Color.Orange;
        }

        private void TextBoxA_A_Changed(object sender, EventArgs e)
        {
            textBox_AIA.BackColor = Color.Orange;
        }

        private void TextBoxA_B_Changed(object sender, EventArgs e)
        {
            textBox_AIB.BackColor = Color.Orange;
        }

        private void TextBoxA_C_Changed(object sender, EventArgs e)
        {
            textBox_AIC.BackColor = Color.Orange;
        }

        private void TextBoxPhi_A_Changed(object sender, EventArgs e)
        {
            textBox_APHIA.BackColor = Color.Orange;
        }

        private void TextBoxPhi_B_Changed(object sender, EventArgs e)
        {
            textBox_APHIB.BackColor = Color.Orange;
        }

        private void TextBoxPhi_C_Changed(object sender, EventArgs e)
        {
            textBox_APHIC.BackColor = Color.Orange;
        }

        private void PowerControls(bool state)
        {
            buttonSetV.Enabled = state;
            button_VAll.Enabled = state;
            button_VA.Enabled = state;
            button_VB.Enabled = state;
            button_VC.Enabled = state;

            textBox_VA.Enabled = state;
            textBox_VB.Enabled = state;
            textBox_VC.Enabled = state;

            buttonSetA.Enabled = state;
            button_IAll.Enabled = state;
            button_IA.Enabled = state;
            button_IB.Enabled = state;
            button_IC.Enabled = state;

            textBox_IA.Enabled = state;
            textBox_IB.Enabled = state;
            textBox_IC.Enabled = state;

            button_PHIAll.Enabled = state;
            button_PHIA.Enabled = state;
            button_PHIB.Enabled = state;
            button_PHIC.Enabled = state;

            textBox_PHIA.Enabled = state;
            textBox_PHIB.Enabled = state;
            textBox_PHIC.Enabled = state;
        }

        private int GetNumericalInput(TextBox tBox, int iMin, int iMax)
        {
            int value = (int)ConvertToDouble(tBox.Text);

            value = value < 0 ? iMin : value;
            value = value > iMax ? iMax : value;

            tBox.Text = value.ToString();

            return value;
        }

        private void timerUpdateStats_Tick(object sender, EventArgs e)
        {
            UpdateStats();
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (string.IsNullOrEmpty(comboBoxMeterType.Text))
            {
                LogMessage(MSG.ERROR, "Please select meter type");
                return;
            }

            try
            {
                if (listBoxProcedures.SelectedItem.ToString() == null)
                {
                    LogMessage(MSG.ERROR, "Please select a procedure");
                    return;
                }
            }
            catch
            {
                LogMessage(MSG.ERROR, "Please select a procedure");
                return;
            }

            DGV_ProcedureList.Rows.Add();
            DGV_ProcedureList.Rows[iProcedureListRow].Cells[0].Value = iProcedureListRow+1;
            DGV_ProcedureList.Rows[iProcedureListRow].Cells[1].Value = comboBoxMeterType.Text;
            DGV_ProcedureList.Rows[iProcedureListRow].Cells[2].Value = listBoxProcedures.SelectedItem.ToString();
            iProcedureListRow++;

            if (GetProcedure(listBoxProcedures.SelectedItem.ToString()))
            {
                return;
            }

            CopyDataGridView(listBoxProcedures.SelectedItem.ToString());
        }

        private void CopyDataGridView(string strProcedure)
        {
            try
            {
                if (DGV_ProcedureRun.Columns.Count == 0)
                {
                    foreach (DataGridViewColumn DGV_Column in DGV_Procedure.Columns)
                    {
                        DGV_ProcedureRun.Columns.Add(DGV_Column.Clone() as DataGridViewColumn);
                    }

                    // Add extra columns for base values
                    DataGridViewTextBoxColumn DGV_Col1 = new DataGridViewTextBoxColumn();
                    DGV_Col1.HeaderText = "Ub";
                    DGV_ProcedureRun.Columns.Add(DGV_Col1);

                    DataGridViewTextBoxColumn DGV_Col2 = new DataGridViewTextBoxColumn();
                    DGV_Col2.HeaderText = "Ib";
                    DGV_ProcedureRun.Columns.Add(DGV_Col2);

                    DataGridViewTextBoxColumn DGV_Col3 = new DataGridViewTextBoxColumn();
                    DGV_Col3.HeaderText = "Im";
                    DGV_ProcedureRun.Columns.Add(DGV_Col3);

                    DataGridViewTextBoxColumn DGV_Col4 = new DataGridViewTextBoxColumn();
                    DGV_Col4.HeaderText = "MeterType";
                    DGV_ProcedureRun.Columns.Add(DGV_Col4);

                    DataGridViewTextBoxColumn DGV_Col5 = new DataGridViewTextBoxColumn();
                    DGV_Col5.HeaderText = "Run";
                    DGV_ProcedureRun.Columns.Add(DGV_Col5);
                }

                DataGridViewRow DGV_Row = new DataGridViewRow();

                string strRun = DateTime.Now.ToString("yy-MM-dd  HHmm") + "  " + strProcedure;

                for (int i = 0; i < DGV_Procedure.Rows.Count; i++)
                {
                    DGV_Row = (DataGridViewRow)DGV_Procedure.Rows[i].Clone();

                    int intColIndex = 0;

                    foreach (DataGridViewCell DGV_Cell in DGV_Procedure.Rows[i].Cells)
                    {
                        DGV_Row.Cells[intColIndex].Value = DGV_Cell.Value;
                        intColIndex++;
                    }

                    DGV_ProcedureRun.Rows.Add(DGV_Row);
                    DGV_ProcedureRun.Rows[iProcedureRunRow].Cells[30].Value = Electricals.iUBase;
                    DGV_ProcedureRun.Rows[iProcedureRunRow].Cells[31].Value = Electricals.iIBase;
                    DGV_ProcedureRun.Rows[iProcedureRunRow].Cells[32].Value = Electricals.iIMax;
                    DGV_ProcedureRun.Rows[iProcedureRunRow].Cells[33].Value = comboBoxMeterType.Text;
                    DGV_ProcedureRun.Rows[iProcedureRunRow].Cells[34].Value = strRun;
                    iProcedureRunRow++;
                }

                DGV_ProcedureRun.AllowUserToAddRows = false;
                DGV_ProcedureRun.Refresh();
            }
            catch (Exception ex)
            {
                LogMessage(MSG.ERROR, ex.ToString());
            }
        }

        private void buttonClearLogs_Click(object sender, EventArgs e)
        {
            ClearMessageLogs();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (File.Exists(Settings.strRootPath + @"\Bypassed.txt"))
            {
                File.Delete(Settings.strRootPath + @"\Bypassed.txt");
                LogMessage(MSG.INFO, "File.Delete(" + Settings.strRootPath + @"\Bypassed.txt)");
            }
        }

        private void buttonClearList_Click(object sender, EventArgs e)
        {
            if (BgWorker.IsBusy)
            {
                LogMessage(MSG.INFO, "Program is busy");
                return;
            }

            ClearMessageLogs();

            progressBar.Value = 0;

            listBox01to24.Items.Clear();
            listBox25to48.Items.Clear();

            DGV_ProcedureList.Rows.Clear();
            iProcedureListRow = 0;
            DGV_ProcedureRun.Columns.Clear();
            iProcedureRunRow = 0;
            DGV_Results.Columns.Clear();

            bResultsSavedAccess = false;
            bResultsSavedSql = false;

            ClearMeterResults();
            UpdateTestBoard();

            LogMessage(MSG.INFO, "Cleared procedure data");
        }

        private void buttonSetV_Click(object sender, EventArgs e)
        {
            DoWorkEventArgs a = null;
            double[] dV = new double[3];

            ButtonsEnabled(false);

            dV[0] = ConvertToDouble(textBox_VA.Text);
            dV[1] = ConvertToDouble(textBox_VB.Text);
            dV[2] = ConvertToDouble(textBox_VC.Text);

            // Safety
            dV[0] = dV[0] > Bench.valMaxMin.uMax ? Bench.valMaxMin.uMax : dV[0];
            dV[1] = dV[1] > Bench.valMaxMin.uMax ? Bench.valMaxMin.uMax : dV[1];
            dV[2] = dV[2] > Bench.valMaxMin.uMax ? Bench.valMaxMin.uMax : dV[2];

            // Sanity check
            dV[0] = dV[0] < 0 ? 0 : dV[0];
            dV[1] = dV[1] < 0 ? 0 : dV[1];
            dV[2] = dV[2] < 0 ? 0 : dV[2];

            SetVoltage(a, false, 1, 1, dV);

            ButtonsEnabled(true);
        }

        private void buttonSetA_Click(object sender, EventArgs e)
        {
            DoWorkEventArgs a = null;
            double[] dI = new double[3];
            double[] dP = new double[3];

            ButtonsEnabled(false);

            dI[0] = ConvertToDouble(textBox_IA.Text);
            dI[1] = ConvertToDouble(textBox_IB.Text);
            dI[2] = ConvertToDouble(textBox_IC.Text);

            // Safety
            dI[0] = dI[0] > Bench.valMaxMin.iMax ? Bench.valMaxMin.iMax : dI[0];
            dI[1] = dI[1] > Bench.valMaxMin.iMax ? Bench.valMaxMin.iMax : dI[1];
            dI[2] = dI[2] > Bench.valMaxMin.iMax ? Bench.valMaxMin.iMax : dI[2];

            // Sanity check
            dI[0] = dI[0] < 0 ? 0 : dI[0];
            dI[1] = dI[1] < 0 ? 0 : dI[1];
            dI[2] = dI[2] < 0 ? 0 : dI[2];

            dP[0] = ConvertToDouble(textBox_PHIA.Text);
            dP[1] = ConvertToDouble(textBox_PHIB.Text);
            dP[2] = ConvertToDouble(textBox_PHIC.Text);

            // Sanity check
            dP[0] = dP[0] > 360 ? 360 : dP[0];
            dP[1] = dP[1] > 360 ? 360 : dP[1];
            dP[2] = dP[2] > 360 ? 360 : dP[2];

            // Sanity check
            dP[0] = dP[0] < 0 ? 0 : dP[0];
            dP[1] = dP[1] < 0 ? 0 : dP[1];
            dP[2] = dP[2] < 0 ? 0 : dP[2];

            SetCurrent(a, false, 1, false, dI, dP);

            ButtonsEnabled(true);
        }
    }

    public class DataGridViewProgressColumn : DataGridViewImageColumn
    {
        public DataGridViewProgressColumn()
        {
            CellTemplate = new DataGridViewProgressCell();
        }
    }

    class DataGridViewProgressCell : DataGridViewImageCell
    {
        static Random Random = new Random();

        // Used to make custom cell consistent with a DataGridViewImageCell
        static System.Drawing.Image emptyImage;

        static DataGridViewProgressCell()
        {
            emptyImage = new Bitmap(1, 1, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
        }

        public DataGridViewProgressCell()
        {
            this.ValueType = typeof(int);
        }

        // Method required to make the Progress Cell consistent with the default Image Cell.
        // The default Image Cell assumes an Image as a value, although the value of the Progress Cell is an int.
        protected override object GetFormattedValue(object value, int rowIndex, ref DataGridViewCellStyle cellStyle, TypeConverter valueTypeConverter, TypeConverter formattedValueTypeConverter, DataGridViewDataErrorContexts context)
        {
            return emptyImage;
        }

        protected override void Paint(System.Drawing.Graphics g, System.Drawing.Rectangle clipBounds, System.Drawing.Rectangle cellBounds, int rowIndex, DataGridViewElementStates cellState, object value, object formattedValue, string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle, DataGridViewPaintParts paintParts)
        {
            int pad = 2;

            int R = Random.Next(0, 256);
            int G = Random.Next(0, 256);
            int B = Random.Next(0, 256);

            if (value == null)
            {
                value = 0;
            }

            try
            {
                float progressVal = (int)value;
                // Need to convert to float before division; otherwise C# returns int which is 0 for anything but 100%
                float percentage = progressVal / 75.0f;
                Brush backColorBrush = new SolidBrush(cellStyle.BackColor);
                Brush foreColorBrush = new SolidBrush(cellStyle.ForeColor);

                // Draws the cell grid
                base.Paint(g, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, (paintParts & ~DataGridViewPaintParts.ContentForeground));

                if (percentage >= 1.0f)
                {
                    percentage = 1.0f;
                }

                // Draw the progress bar and the text
                g.FillRectangle(new SolidBrush(Color.FromArgb(R, G, B)), cellBounds.X + pad, cellBounds.Y + pad, Convert.ToInt32((percentage * cellBounds.Width - (2 * pad))), cellBounds.Height - (2 * pad));
                g.DrawString((progressVal / 100).ToString("F"), cellStyle.Font, foreColorBrush, cellBounds.X + (cellBounds.Width / 2) - 5, cellBounds.Y + 4);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }

    public class Meter
    {
        public bool Active { get; set; }
        public bool Saved { get; set; }
        public int Failed { get; set; }
        public string Status { get; set; }
        public double[] ErrorV = new double[EzyCal.iMaxSteps];
        public int[] ErrorR = new int[EzyCal.iMaxSteps];
        public string MeterType { get; set; }
        public string MSN { get; set; }
        public string OwnerNo { get; set; }
        public string ContractNo { get; set; }
        public string Client { get; set; }
        public string ClientNo { get; set; }
        public string Firmware { get; set; }
        public string MAC { get; set; }
        public string DateTime { get; set; }
        public string Program { get; set; }
        public string Ripple { get; set; }
    }

    public class Extremes
    {
        public double uMax = 288.0;
	    public double iMax = 120.0;
	    public double freqMin;
	    public double freqMax;
    };

    public class Bench
    {
        public int numPhase = 1;
        public int config;
        public Extremes valMaxMin = new Extremes();
        public short ctrlSio = 1;
        public string ctrlSioFmt = "19200,n,8,2";
        public int numPosition = 20;
        public string BenchName;
        public string Owner;
        public string RefStdName;
    };

    public class Actuals
    {
        public int numPhase;
        public double uA, uB, uC;
        public double iA, iB, iC;
        public double phiA, phiB, phiC;
        public double freq;
        public double totalP;
        public double totalQ;
        public double totalS;
        public int isValid;
    }

    public class Electricals
    {
        public bool bA, bB, bC; 
        public double uA, uB, uC;
        public double iA, iB, iC;
        public double phi;
        public double freq;
        public int isImax;
        public int phaseSeq;
        public int waveForm;
        public int lineType;
        public int connectMode;
        public int principle;
        public int iUBase = 240;
        public int iIBase = 10;
        public int iIMax = 100;
    };

    public class ErrorCounter
    {
        public double mtrConst = 1000;
        public long NumPulses = 5;
        public double uLimit = 0.5;
        public double lLimit = -0.5;
        public int ChannelNo = 1;
        public int NumDecPlace;
        public double StCpTestPeriod;
        public int Measurement = 1;
    };

    public class StepValues
    {
        public int ProcedureID;
        public int PStepNo;
        public string Name;
        public double UA;
        public double UB;
        public double UC;
        public double IA;
        public double IB;
        public double IC;
        public int IsImax;
        public int PHI;
        public int FREQ;
        public int Waveform;
        public int PhaseSeq;
        public int TestTypeID;
        public int Measurement;
        public int NumPulses;
        public double ULIMIT;
        public double LLIMIT;
        public int ChannelNo;
        public int Storing;
        public int FileIE;
        public int Duration;
        public int Timeout;
        public int Finally;
        public string ACMDS;
        public string BCMDS;
        public string CCMDS;
        public int WithAmp;
        public int MinimumTime;
        public int BaseUb;
        public int BaseIb;
        public int BaseImax;
    }

    public class Counters
    {
        public int iMsgInfo = 0;
        public int iMsgWarning = 0;
        public int iMsgErrors = 0;
        public int iMsgFailure = 0;
        public int iMsgDebug = 0;
        public int iSetVoltage = 0;
        public int iSetVoltageErrors = 0;
        public int iSetCurrent = 0;
        public int iSetCurrentErrors = 0;
        public int iReadRefMeter = 0;
        public int iReadRefMeterErrors = 0;
        public int iSetErrorCounter = 0;
        public int iSetErrorCounterErrors = 0;
        public int iReadErrorCounter = 0;
        public int iReadErrorCounterErrors = 0;
    }

    public class Settings
    {
        public bool bLogDebug = false;
        public bool bLogInfo = true;
        public int iSleepDivide = 1;
        public bool bErrorsGenerate = false;
        public bool bErrorsType = false;
        public double dErrorsRange = 0.051;
        public bool bRunCmds = true;
        public bool bRunCtrComm = true;
        public bool bFilesExist = true;
        public bool bProcessExitCode = true;
        public bool bSaveAccess = false;
        public bool bSaveSql = false;
        public bool bSaveCsv = false;
        public string strProcedurePC = "None";
        public string strRootPath = Directory.GetCurrentDirectory();
    }
}
