using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
//using SolidWorks.Interop.sldworks;
//using SolidWorks.Interop.swcommands;
//using SolidWorks.Interop.swconst;
//using SolidWorks.Interop.cosworks;

namespace ThermalAnalysisCPU
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // forbidding anything but digits
        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[^0-9]"))
            {
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);
            }
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[^0-9]"))
            {
                textBox2.Text = textBox2.Text.Remove(textBox2.Text.Length - 1);
            }
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox3.Text, "[^0-9]"))
            {
                textBox3.Text = textBox3.Text.Remove(textBox3.Text.Length - 1);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox4.Text, "[^0-9]"))
            {
                textBox4.Text = textBox4.Text.Remove(textBox4.Text.Length - 1);
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox7.Text, "[^0-9]"))
            {
                textBox7.Text = textBox7.Text.Remove(textBox7.Text.Length - 1);
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox5.Text, "[^0-9]"))
            {
                textBox5.Text = textBox5.Text.Remove(textBox5.Text.Length - 1);
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox8.Text, "[^0-9]"))
            {
                textBox8.Text = textBox8.Text.Remove(textBox8.Text.Length - 1);
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox6.Text, "[^0-9]"))
            {
                textBox6.Text = textBox6.Text.Remove(textBox6.Text.Length - 1);
            }
        }

        // closing all current SW processes
        private void DisposeSW()
        {
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (Process process in processes)
            {
                process.CloseMainWindow();
                process.Kill();
            }
        }

        // reference to SW window
        SldWorks swApp;

        private void button1_Click_1(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button1.Text = "Proceeding...";
            // preventing empty fields
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text) 
                || string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(textBox4.Text) 
                || string.IsNullOrWhiteSpace(textBox5.Text) || string.IsNullOrWhiteSpace(textBox6.Text) 
                || string.IsNullOrWhiteSpace(textBox7.Text) || string.IsNullOrWhiteSpace(textBox8.Text))
            {
                MessageBox.Show("Enter all values", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int time = Convert.ToInt32(textBox1.Text);
            int step = Convert.ToInt32(textBox2.Text);
            int initt1 = Convert.ToInt32(textBox3.Text);
            int initt2 = Convert.ToInt32(textBox4.Text);
            int initt3 = Convert.ToInt32(textBox7.Text);
            int conv = Convert.ToInt32(textBox5.Text);
            double bulk = Convert.ToDouble(textBox8.Text) + 273.15; // converting to Kelvin units
            int heatpow = Convert.ToInt32(textBox6.Text);

            DisposeSW();

            // Opening SW 2017
            // Additional GUID: F16137AD-8EE8-4D2A-8CAC-DFF5D1F67522
            Guid Guid1 = new Guid("6B36082E-677B-49E8-BCB2-76698EBD2906");

            object processSW = System.Activator.CreateInstance(System.Type.GetTypeFromCLSID(Guid1));

            swApp = (SldWorks)processSW;
            swApp.Visible = true;

            CosmosWorks COSMOSWORKS = default(CosmosWorks);
            CwAddincallback CWObject = default(CwAddincallback);
            CWModelDoc ActDoc = default(CWModelDoc);
            CWStudyManager StudyMngr = default(CWStudyManager);
            CWStudy Study = default(CWStudy);
            CWThermalStudyOptions ThermalOptions = default(CWThermalStudyOptions);
            CWConvection CWConv = default(CWConvection);
            CWMesh CwMesh = default(CWMesh);
            CWResults CWResult = default(CWResults);
            CWTemperature CWTemp = default(CWTemperature);
            ModelDoc2 Part = default(ModelDoc2);
            CWLoadsAndRestraintsManager LBCMgr = default(CWLoadsAndRestraintsManager);
            CWHeatPower CWHeatPower = default(CWHeatPower);
            ModelDoc2 swDoc = null;
            int bApp = 0;
            int longstatus = 0;
            int longwarnings = 0;
            int errCode = 0;
            double el = 0;
            double tl = 0;
            string cwpath = "";

            Part = swApp.OpenDoc6("G:\\assem(simple).SLDASM", (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
            if (Part == null)
                MessageBox.Show("Failed to open document");

            swDoc = swApp.IActiveDoc2;

            // connecting Simulation to SW
            cwpath = swApp.GetExecutablePath() + @"\Simulation\cosworks.dll";
            errCode = swApp.LoadAddIn(cwpath);
            if (errCode != 0 && errCode != 2)
            {
                MessageBox.Show("Failed to load Simulation library");
                swApp.ExitApp();
                DisposeSW();
                Application.Exit();
            }

            // ready-check Simulation add-on
            CWObject = (CwAddincallback)swApp.GetAddInObject("CosmosWorks.CosmosWorks");
            if (CWObject == null)
                MessageBox.Show("Failed to start Simulation");

            // ready-check add-on object
            COSMOSWORKS = CWObject.CosmosWorks;
            if (COSMOSWORKS == null)
                MessageBox.Show("Failed to create CosmosWorks object");

            // connecting add-on to document
            ActDoc = COSMOSWORKS.ActiveDoc;
            if (ActDoc == null)
                MessageBox.Show("Failed to open document");

            // creating thermal analysis
            bApp = ActDoc.AddDefaultThermalStudyPlot((int)swsThermalResultComponentTypes_e.swsThermalResultComponentTypes_TEMP, true);
            StudyMngr = ActDoc.StudyManager;
            if (StudyMngr == null)
                MessageBox.Show("Failed to start analysis manager");

            Study = StudyMngr.CreateNewStudy3("CPU Usage", (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeThermal,
                (int)swsMeshType_e.swsMeshTypeSolid, out errCode);
            if (Study == null)
                MessageBox.Show("Failed to create thermal analysis");

            ThermalOptions = Study.ThermalStudyOptions;
            Study.ThermalStudyOptions.SolutionType = 0;
            if (ThermalOptions == null)
                MessageBox.Show("No analysis parameters");

            // analysis duration
            ThermalOptions.CheckFlowConvectionCoef = 0;
            ThermalOptions.SolverType = 2;
            ThermalOptions.TotalTime = time;
            ThermalOptions.TimeIncrement = step;

            LBCMgr = Study.LoadsAndRestraintsManager;
            if (LBCMgr == null)
                MessageBox.Show("Failed to load loads and restraints manager");

            // lit coordinates
            swDoc.ShowNamedView2("*Top", 5);
            swDoc.ViewZoomtofit2();
            ISelectionMgr selectionMgr = (ISelectionMgr)swDoc.SelectionManager;
            bool isSelected = swDoc.Extension.SelectByRay(-0.28387922510520769, 0.081999999999993634,
                -0.14613487203691797, 0, -1, 0, 0.0047546274155552031, 2, false, 0, 0);
            if (isSelected)
            {
                Entity swEntity = selectionMgr.GetSelectedObject6(1, -1);
                Component2 swComp = (Component2)swEntity.GetComponent();
                object[] fixedpart = { swComp };
                CWTemp = LBCMgr.AddTemperature(fixedpart, out errCode);
                if (errCode != 0)
                    MessageBox.Show("Failed to add temperature indicator");
            }

            CWTemp.TemperatureBeginEdit();
            CWTemp.TemperatureType = 0;
            CWTemp.Unit = 2;
            CWTemp.TemperatureValue = initt2;

            errCode = CWTemp.TemperatureEndEdit();
            if (errCode != 0)
                MessageBox.Show("Failed to edit temperature indicator");

            swDoc.ClearSelection2(true);

            // silicon board coordinates
            isSelected = swDoc.Extension.SelectByRay(-0.41952594843134139, 0.02199999999999136,
                -0.41742831868918501, 0, -1, 0, 0.0047546274155552031, 2, false, 0, 0);
            if (isSelected)
            {
                Entity swEntity = selectionMgr.GetSelectedObject6(1, -1);
                Component2 swComp = (Component2)swEntity.GetComponent();
                object[] fixedpart = { swComp };
                CWTemp = LBCMgr.AddTemperature(fixedpart, out errCode);
                if (errCode != 0)
                    MessageBox.Show("Failed to add temperature indicator");
            }

            CWTemp.TemperatureBeginEdit();
            CWTemp.TemperatureType = 0;
            CWTemp.Unit = 2;
            CWTemp.TemperatureValue = initt1;

            errCode = CWTemp.TemperatureEndEdit();
            if (errCode != 0)
                MessageBox.Show("Failed to edit temperature indicator");

            swDoc.ClearSelection2(true);

            // stalks coordinates
            swDoc.ShowNamedView2("*Bottom", 6);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(-0.42931488722807265, -0.050999999999987722, 0.28597685484736474, 0, 1, 0, 0.0047546274155552031, 2, false, 0, 0);
            if (isSelected)
            {
                Entity swEntity = selectionMgr.GetSelectedObject6(1, -1);
                Component2 swComp = (Component2)swEntity.GetComponent();
                object[] fixedpart = { swComp };
                CWTemp = LBCMgr.AddTemperature(fixedpart, out errCode);
                if (errCode != 0)
                    MessageBox.Show("Failed to edit temperature indicator");
            }

            CWTemp.TemperatureBeginEdit();
            CWTemp.TemperatureType = 0;
            CWTemp.Unit = 2;
            CWTemp.TemperatureValue = initt3;

            errCode = CWTemp.TemperatureEndEdit();
            if (errCode != 0)
                MessageBox.Show("Failed to edit temperature indicator");

            swDoc.ClearSelection2(true);

            // all faces coordinates
            swDoc.ShowNamedView2("*Front", 1);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(-0.38382731822234351, 0.04982938298755997, 0.40017589628916994, 0, 0, -1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.18695037482381383, 0.049829382987560122, 0.42176499999999351, 0, 0, -1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.35735646868976789, 0.0713369482327776, 0.41521684512093771, 0, 0, -1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.10505618408240869, 0.010950322736590109, 0.4415000000000191, 0, 0, -1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.43015130490435044, -0.017174954891772284, 0.29464212540386825, 0, 0, -1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.42932409085645729, -0.016347740843879296, 0.29483222724246616, 0, 0, -1, 0.0028125277628361371, 2, false, 0, 0);

            swDoc.ShowNamedView2("*Back", 2);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(0.3581836827376611, 0.063064807753847776, -0.4148975912216315, 0, 0, 1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.14558967242916476, 0.053138239179132019, -0.42176499999999351, 0, 0, 1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.37969124798287845, 0.047347740843881027, -0.40327910794843547, 0, 0, 1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.033088561915719295, 0.012604750832376084, -0.4415000000000191, 0, 0, 1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.42766966276067148, -0.014693312748093316, -0.29499856062381014, 0, 0, 1, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.42766966276067131, -0.022138239179130152, -0.29499856062381014, 0, 0, 1, 0.0028125277628361371, 2, false, 0, 0);

            swDoc.ShowNamedView2("*Left", 3);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(-0.42176499999999351, 0.046520526795988171, -0.081066976693512371, 1, 0, 0, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.4415000000000191, 0.0084686805929112003, -0.082721404789298347, 1, 0, 0, 0.0028125277628361371, 2, false, 0, 0);

            swDoc.ShowNamedView2("*Right", 4);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(0.42176499999999351, 0.052311025131239031, 0.1158099667050172, -1, 0, 0, 0.0028125277628361371, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.4415000000000191, 0.010123108688697121, 0.12490932123184006, -1, 0, 0, 0.0028125277628361371, 2, false, 0, 0);

            swDoc.ShowNamedView2("*Top", 5);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(-0.19376272580646192, 0.081999999999993634, -0.13541980560506936, 0, -1, 0, 0.0036397418107291194, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.41535877134202892, 0.02199999999999136, -0.42017607667975893, 0, -1, 0, 0.0036397418107291194, 2, false, 0, 0);

            swDoc.ShowNamedView2("*Bottom", 6);
            swDoc.ViewZoomtofit2();
            isSelected = swDoc.Extension.SelectByRay(-0.20232682418464809, -0.00099999999997635314, 0.23818898614330264, 0, 1, 0, 0.0036397418107291194, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.42713440661203494, -0.050999999999987722, 0.28743255181787342, 0, 1, 0, 0.0036397418107291194, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.42713440661203483, -0.050999999999987722, 0.28529152722332685, 0, 1, 0, 0.0036397418107291194, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(0.4292754312065814, -0.050999999999987722, -0.28422101492605301, 0, 1, 0, 0.0036397418107291194, 2, false, 0, 0);
            isSelected = swDoc.Extension.SelectByRay(-0.42499338201748837, -0.050999999999987722, -0.28636203952059958, 0, 1, 0, 0.0036397418107291194, 2, false, 0, 0);
            if (isSelected)
            {
                object selectedFace1 = (object)selectionMgr.GetSelectedObject6(1, -1);
                object selectedFace2 = (object)selectionMgr.GetSelectedObject6(2, -1);
                object selectedFace3 = (object)selectionMgr.GetSelectedObject6(3, -1);
                object selectedFace4 = (object)selectionMgr.GetSelectedObject6(4, -1);
                object selectedFace5 = (object)selectionMgr.GetSelectedObject6(5, -1);
                object selectedFace6 = (object)selectionMgr.GetSelectedObject6(6, -1);
                object selectedFace7 = (object)selectionMgr.GetSelectedObject6(7, -1);
                object selectedFace8 = (object)selectionMgr.GetSelectedObject6(8, -1);
                object selectedFace9 = (object)selectionMgr.GetSelectedObject6(9, -1);
                object selectedFace10 = (object)selectionMgr.GetSelectedObject6(10, -1);
                object selectedFace11 = (object)selectionMgr.GetSelectedObject6(11, -1);
                object selectedFace12 = (object)selectionMgr.GetSelectedObject6(12, -1);
                object selectedFace13 = (object)selectionMgr.GetSelectedObject6(13, -1);
                object selectedFace14 = (object)selectionMgr.GetSelectedObject6(14, -1);
                object selectedFace15 = (object)selectionMgr.GetSelectedObject6(15, -1);
                object selectedFace16 = (object)selectionMgr.GetSelectedObject6(16, -1);
                object selectedFace17 = (object)selectionMgr.GetSelectedObject6(17, -1);
                object selectedFace18 = (object)selectionMgr.GetSelectedObject6(18, -1);
                object selectedFace19 = (object)selectionMgr.GetSelectedObject6(19, -1);
                object selectedFace20 = (object)selectionMgr.GetSelectedObject6(20, -1);
                object selectedFace21 = (object)selectionMgr.GetSelectedObject6(21, -1);
                object selectedFace22 = (object)selectionMgr.GetSelectedObject6(22, -1);
                object selectedFace23 = (object)selectionMgr.GetSelectedObject6(23, -1);
                object[] fixedpart = { selectedFace1, selectedFace2, selectedFace3, selectedFace4, selectedFace5, selectedFace6, selectedFace7, selectedFace8, selectedFace9, selectedFace10,
                selectedFace11, selectedFace12, selectedFace13, selectedFace14, selectedFace15, selectedFace16, selectedFace17, selectedFace18, selectedFace19, selectedFace20, selectedFace21,
                selectedFace22, selectedFace23 };
                CWConv = LBCMgr.AddConvection(fixedpart, out errCode);
                if (errCode != 0)
                    MessageBox.Show("Failed to add convection coefficient indicator");
            }

            CWConv.ConvectionBeginEdit();
            CWConv.Unit = 0;
            CWConv.ConvectionCoefficient = conv;
            CWConv.BulkAmbientTemperature = bulk;

            errCode = CWConv.ConvectionEndEdit();
            if (errCode != 0)
                MessageBox.Show("Failed to edit convection coefficient indicator");

            swDoc.ClearSelection2(true);

            // stalks coordinates
            isSelected = swDoc.Extension.SelectByRay(-0.42931488722807265, -0.050999999999987722, 0.28597685484736474, 0, 1, 0, 0.0047546274155552031, 2, false, 0, 0);
            if (isSelected)
            {
                Entity swEntity = selectionMgr.GetSelectedObject6(1, -1);
                Component2 swComp = (Component2)swEntity.GetComponent();
                object[] fixedpart = { swComp };
                CWHeatPower = LBCMgr.AddHeatPower(fixedpart, out errCode);
                if (errCode != 0)
                    MessageBox.Show("Failed to add thermal power indicator");
            }

            CWHeatPower.HeatPowerBeginEdit();
            CWHeatPower.Unit = 0;
            CWHeatPower.HPValue = heatpow;
            CWHeatPower.ReverseDirection = 0;
            CWHeatPower.IncludeThermostat = 0;
            errCode = CWHeatPower.HeatPowerEndEdit();
            if (errCode != 0)
                MessageBox.Show("Failed to edit thermal power indicator");

            swDoc.ShowNamedView2("*Isometric", 7);
            swDoc.ViewZoomtofit2();

            // creating mesh model
            CwMesh = Study.Mesh;
            if (CwMesh == null)
                MessageBox.Show("No mesh model");

            CwMesh.Quality = 1;
            CwMesh.GetDefaultElementSizeAndTolerance(0, out el, out tl);
            errCode = Study.CreateMesh(0, el, tl);
            if (errCode != 0)
                MessageBox.Show("Failed to create mesh model");

            // analysis start
            errCode = Study.RunAnalysis();
            if (errCode != 0)
                MessageBox.Show("Failed to start analysi: " + errCode);

            CWResult = Study.Results;
            if (CWResult == null)
            {
                MessageBox.Show("No results");
            }
            else
            {
                MessageBox.Show("Done", "Results");
                button1.Enabled = true;
                button1.Text = "Thermal analysis";
            }
            double[] max = new double[time / step];
            double[] min = new double[time / step];
            double[] avg = new double[4];

            for (int i = 1; i <= time / step; i++)
            {
                CWPlot plot = CWResult.GetPlot("Thermal" + i.ToString(), out errCode);
                plot.SetComponentUnitAndValueByElem(0, 2, false);
                object[] results = null;
                results = (object[])plot.GetMinMaxResultValues(out errCode);
                max[i - 1] = Convert.ToDouble(results[3].ToString());
                min[i - 1] = Convert.ToDouble(results[1].ToString());
                if (i == 1)
                {
                    avg[0] = Convert.ToDouble(results[1].ToString());
                    avg[1] = Convert.ToDouble(results[3].ToString());
                }
                if (i == time / step)
                {
                    avg[2] = Convert.ToDouble(results[1].ToString());
                    avg[3] = Convert.ToDouble(results[3].ToString());
                }
            }

            double maxValue = max.Max();
            double minValue = min.Min();

            textBox9.Text = maxValue.ToString("#.##");
            textBox10.Text = minValue.ToString("#.##");

            double heatdens = Math.Abs(conv * (((initt2 + initt3) / 2) - (initt3 - (maxValue - minValue))));
            textBox12.Text = heatdens.ToString("#.##");

            double avg1 = (avg[0] + avg[1]) / 2;
            double avg2 = (avg[2] + avg[3]) / 2;
            double avgr = avg1 - avg2;
            textBox15.Text = avgr.ToString("#.##");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox5.Image = ThermalAnalysisCPU.Properties.Resources.full;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            pictureBox5.Image = ThermalAnalysisCPU.Properties.Resources.part1;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            pictureBox5.Image = ThermalAnalysisCPU.Properties.Resources.part2;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            pictureBox5.Image = ThermalAnalysisCPU.Properties.Resources.part3;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            pictureBox5.Image = ThermalAnalysisCPU.Properties.Resources.part3_simple_;
        }
    }
}
