#region Version
/*
 * This Version is used for COMPUTER SYSTEM VALIDATION 
 * Test for Single-Use
 * Author: Li Shao 30-JUN-2015
 */
#endregion

#region Configuration and Technical Requirements
// Configuration Requirements


// Technical Requirements
 

#endregion

#region References
//** Initial Reference
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;

using System.Threading;

//Test
using System.Diagnostics;

//** Project -> Add Reference... -> COM -> 
//** Microsoft Excel 15.0 Object Library
using Excel = Microsoft.Office.Interop.Excel;

/*
* Third-party Reference
* Third-party library downloaded from 
* http://sourceforge.net/projects/itextsharp/
* version: 5.5.0
*/
//** Project -> Add Reference... -> Browse -> add .dll files
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

#endregion

namespace QN_Excel_Tool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            //** Disable converter button
            button1.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //***********************************************Pre-Defined Structures********************************************************//
        public string[] selected_files; //** Define global variables

        //** Operator for filtering Basic Information
        private string[] BasicInfo = { "Notification :", "QN Type :", "Material:", "Plant:", 
                                     "Description:", "QN Responsible:", "Notification Date : "};

        //** Listed on Excel form (Category)
        private string[] listed_items_excel = { "Notification","Plant","Tasks.QE & Activities" };
        private string[] RMM_info_item = { "QN File Name", "QN ID", "Part Risk Classification", "RMM Doc ID", "RMM Entry", "pFMEA", "Description", "Notif Date"+'\n'+"(dd.mm.yyyy)", "Material" };

        //** Define the strings for delimiters    
        
        private string[] stringSeparates = new string[] { "Tasks:", "Activities:" };
        
        private string[] delimiter5 = new string[] { "_______________________________________________________________" };
        private string[] task_Activity_parts = { };

        #region Regular Expression for Risk Level Determining
        /*
         * Regular Expression
         * 
         * Pattern 1: ([^a-zA-Z]W[^a-zA-Z])
         * Looking for a group that has three characters: 
         * non-letter, Captical-Letter(W, X, Y, Z), non-letter.
         */
        string pattern_1 = "([^a-zA-Z]W[^a-zA-Z])";
        string pattern_2 = "([^a-zA-Z]X[^a-zA-Z])";
        string pattern_3 = "([^a-zA-Z]Y[^a-zA-Z])";
        string pattern_4 = "([^a-zA-Z]Z[^a-zA-Z])";
        string pattern_risk = "([Rr]isk[ \n][Cc]las[s]ification)"; // Do some typo fix
        string pattern_risk_1 = "([Rr]isk[ \n][Cc]at[ae]gory)"; //Do some typo fix
        string pattern_risk_2 = "([Rr]isk[ \n][Cc]lass[ified])"; //Do some typo fix
        #endregion


        bool Risk_classification_status = false; 

        #region Regular Expression for RMM file
        string pattern_RMM_file1 = "(C16U-0266)"; //Ingenuity
        string pattern_RMM_file2 = "(CP2H-0019)"; //
        string pattern_RMM_file3 = "(DHF-204922)|(DHF204922)|(204922)";
        string pattern_RMM_file4 = "(DHF-204923)|(DHF204923)|(204923)";
        #endregion

        #region Regular Expression for pFMEA file
        string pattern_pFMEA = "(F[ME][ME]A)";
        #endregion
        

        //**********************************************Defined ToolItems Functions******************************************************//
        //** Functions for ToolItems:
        private void filesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        
        private void openFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog open_file = new OpenFileDialog();
            open_file.Multiselect = true; //** Ensure multiple files selection 
            open_file.Filter = "PDF Files(*.PDF)|*.PDF|All Files(*.*)|*.*"; //Read PDF files
            open_file.ShowDialog(); //Show Dialog for selecting
            selected_files = open_file.FileNames;
            //** After choosing files, enable convertion button
            if (selected_files.Length != 0)
            {
                button1.Enabled = true;
            }
        }//** openFilesToolStripMenuItem_Click

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("This Tool is designed for assisting to collect RMM and pFMEA information from QN PDF files."+'\n'+
            "The Tool should generate three Excel worksheets:"+'\n'+"a. The first sheet is optional, if the user check Export Task.QE and Activity, then the sheet would contain the original information regarding Task.QE and Activity from each QN files"+'\n'+
            "b. The second sheet would contain the following information:"+'\n'+"   1. QN file name; 2. QN ID; 3. Part Risk Classification; 4. RMM Doc ID; 5. RMM Entry; 6. pFMEA; 7. Description; 8. Notif Date; 9. Material"+'\n'+"c. The third sheet would contain the original infromation about RMM and pFMEA from QN files"
            , "Help", MessageBoxButtons.OK);
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void US8M_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void US8W_CheckedChanged(object sender, EventArgs e)
        {

        }
        
        private void openExcelFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            #region Initialization of Excel sheet
            //** Should select files
            if (selected_files.Length == 0)
            {
                MessageBox.Show("Please select at least one pdf file", "ERROR", MessageBoxButtons.OK);
                return;
            }
            int file_numbers = selected_files.Length; //** Find the numbers of files

            //** Open Excel
            Excel._Application newExcelForm = new Microsoft.Office.Interop.Excel.Application();
            newExcelForm.Visible = true;
            object missing = System.Reflection.Missing.Value;// set missing for default input variable

            //** Check if Excel has been installed
            if (newExcelForm == null)
            {
                MessageBox.Show("Cannot find Microsoft Excel, make sure it installed", "CRASH", MessageBoxButtons.OK);
                return;
            }

            //** Create new Workbook & Worksheet in Excel
            Excel.Workbook xls_QN_workbook;
            xls_QN_workbook = newExcelForm.Workbooks.Add(missing);

            //** Creating a string buffer
            string[] part_string;
            
            #endregion
            /*
             * SETTING OPTIONS FOR DIFFERENT PLANT TYPE:
             * US8M --- CT
             * US8W --- AMI
             * 1. ONLY CHECK US8M
             * 2. ONLY CHECK US8W
             * 3. BOTH CHECK US8M && US8W
             */
            //** Exporting information base on the Checkbox
            #region CT QN checked
            if (US8M.Checked == true && US8W.Checked == false)
            {
                #region Initialization
                //** Create new Worksheet for US8M
                Excel.Worksheet xls_QN_sheet;
                xls_QN_sheet = (Excel.Worksheet)xls_QN_workbook.Worksheets.get_Item(1);
                xls_QN_sheet.Name = "US8M_QN_files";

                //** Create new Worksheet for RMM files
                Excel.Worksheet xls_QN_sheet_RMM;
                xls_QN_sheet_RMM = (Excel.Worksheet)xls_QN_workbook.
                                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_sheet_RMM.Name = "QN_files_RMM";

                //** Create new Worksheet for RMM Reference information
                Excel.Worksheet xls_QN_RMM_Ref;
                xls_QN_RMM_Ref = (Excel.Worksheet)xls_QN_workbook.
                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_RMM_Ref.Name = "QN_files_RMM_Reference";


                //** Layout of matrix is in order of row 
                int column = 1; //Initialize the column number
                foreach (string info in listed_items_excel)
                {
                    xls_QN_sheet.Cells[1, column].Value = info;
                    xls_QN_sheet.Cells[1, column].Font.Bold = true;
                    xls_QN_sheet.Rows.AutoFit();
                    column++;
                }//** Export Information tilte into Excel sheet

                //** Layout for RMM Worksheet
                int RMM_column = 1;
                foreach (string info in RMM_info_item)
                {
                    xls_QN_sheet_RMM.Cells[1, RMM_column].Value = info;
                    xls_QN_sheet_RMM.Cells[1, RMM_column].Font.Bold = true;
                    xls_QN_sheet_RMM.Rows.AutoFit();
                    RMM_column++;
                }
                // Adding comment for Risk Classification Level
                xls_QN_sheet_RMM.Cells[1, 3].AddComment("W-Critical" + '\n' + "X-Major" + '\n' +
                    "Y-Moderate" + '\n' + "Z-Negligible");
                xls_QN_RMM_Ref.Cells[1, 1] = "QN ID";
                xls_QN_RMM_Ref.Cells[1, 1].Font.Bold = true;
                int row_num = 1; //** indicate the file numbers by iteration
                //***************** iterating File names into exporting info ******************************************//
                //foreach (string file_name in selected_files)
                Stopwatch sw = new Stopwatch();

                sw.Start();
                #endregion
                #region Checking files one by one
                for (int i_file = 0; i_file < file_numbers; i_file++)
                {
                    row_num++;// refresh the index to next column.

                    //** Converting PDF into Text by using thrid-party library: iTextSharp
                    string Text_PDF = string.Empty;
                    
                    //** try to exporting PDF format into text 
                    try
                    {
                        //** OPERATOR for converting PDF to Text
                        Text_PDF = PDF_to_Text(selected_files[i_file]);
                    }
                    catch (Exception Error)
                    {
                        MessageBox.Show(Error.Message);
                    }

                    /** Exporting information into Excel sheet **/
                    //** Separating The whole string into several substring
                    /*Structure of the QN file:
                     * ------------ basic info -----------------
                     * --------Problem Description -------------
                     * ---------------Tasks---------------------
                     * ---------------Activities----------------
                     * ---------Detailed Items------------------
                     */
                    part_string = Text_PDF.Split(stringSeparates, StringSplitOptions.None); //** Split the whole string into several substring by seperating Tasks and Activities

                    //** Extracting Basic info from QN files. 
                    int position1 = part_string[0].IndexOf("Plant:") + "Plant:".Length;

                    //** Filtering US8M into Excel
                    if (part_string[0].Substring(position1, 5) == " US8M")
                    {
                        //** Export Basic Information
                        ExcelBasicInfoTool(row_num, selected_files[i_file], part_string[0], BasicInfo, xls_QN_sheet,
                                           xls_QN_sheet_RMM, xls_QN_RMM_Ref);
                    
                        //** Export Task Description
                        // Recombine Tasks and Activities strings
                        string Task_Activities = String.Empty; //** Define the string of Tasks and Activiteis 
                        for (int index = 1; index < part_string.Length; index++)
                        {
                            Task_Activities = Task_Activities + part_string[index]; //** Re-assemble substrings into a whole string
                            //** NOTE: The reason is because for the following pages in QN files, there are multiple 
                            //**       Tasks and Activities on the headline. It is hard to process substrings. Therefore, re-combine into a whole string
                        }

                        // Filting for Tasks and description
                        /* Separating by 
                         * "_______________________________________________________________"
                         */
                        task_Activity_parts = Task_Activities.Split(delimiter5, StringSplitOptions.None);//** Seperating 

                        //** DETECTING RMM AS KEYWORD
                        //** Trying to BOLD & UNDERSTORE Keyword
                        ExcelRMM_RiskLevelTool(row_num, task_Activity_parts, xls_QN_sheet_RMM);
                        //ExcelRMMPaternMatch(row_num, task_Activity_parts, xls_QN_sheet_RMM);
                        ExcelRMM_pFMEA_Entry(row_num, selected_files[i_file], task_Activity_parts, xls_QN_sheet_RMM,xls_QN_RMM_Ref);

                        //OPERATOR for QN Tasks and Activities exporting
                        if (checkBox1.Checked == true)
                           ExcelFilterTool(row_num,selected_files[i_file] ,task_Activity_parts, xls_QN_sheet);

                    }// IF-condition "US8M"
                    else
                    {
                        //** If cannot find US8M (CT), move to next file and reset the location in the sheet
                        row_num--; 
                        continue;
                    }

                }// FOR-loop iterating input files
                #endregion
                #region Formatting the Excel cells alignment
                //** Changing Format for more usabiliy
                xls_QN_sheet.Range[xls_QN_sheet.Cells[1, 1], xls_QN_sheet.Cells[row_num]].WrapText = true; //** wrap text
                xls_QN_sheet_RMM.Range[xls_QN_sheet_RMM.Cells[1, 1], xls_QN_sheet_RMM.Cells[row_num]].WrapText = true; //** wrap text
                xls_QN_sheet.Columns.AutoFit(); //** auto fit cells
                xls_QN_sheet_RMM.Columns.AutoFit();
                xls_QN_sheet.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop; //** align text into top
                xls_QN_sheet_RMM.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                xls_QN_RMM_Ref.Range[xls_QN_RMM_Ref.Cells[1,1], xls_QN_RMM_Ref.Cells[row_num]].WrapText = true;
                xls_QN_RMM_Ref.Columns.AutoFit();
                xls_QN_RMM_Ref.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                button1.Enabled = false; //** Disable button "" untill next select

                sw.Stop();
                #endregion
                MessageBox.Show(string.Format("Elapsed={0}",sw.Elapsed), "TEST", MessageBoxButtons.OK);

            }// IF-condition "US8M checked& US8W unchecked"
                
            #endregion
            #region AMI QN checked
            else if (US8M.Checked == false && US8W.Checked == true)
            {
                #region Initialization
                //** Create new Worksheet for US8W
                Excel.Worksheet xls_QN_sheet;
                xls_QN_sheet = (Excel.Worksheet)xls_QN_workbook.Worksheets.get_Item(1);
                xls_QN_sheet.Name = "US8W_QN_files";

                //** Create new Worksheet for RMM files
                Excel.Worksheet xls_QN_sheet_RMM;
                xls_QN_sheet_RMM = (Excel.Worksheet)xls_QN_workbook.
                                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_sheet_RMM.Name = "QN_files_RMM";

                //** Create new Worksheet for RMM Reference information
                Excel.Worksheet xls_QN_RMM_Ref;
                xls_QN_RMM_Ref = (Excel.Worksheet)xls_QN_workbook.
                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_RMM_Ref.Name = "QN_files_RMM_Reference";

                //** Layout of matrix is in order of row 
                int column = 1; //Initialize the column number
                foreach (string info in listed_items_excel)
                {
                    xls_QN_sheet.Cells[1, column].Value = info;
                    xls_QN_sheet.Cells[1, column].Font.Bold = true;
                    xls_QN_sheet.Rows.AutoFit();
                    column++;
                }//** Export Information tilte into Excel sheet

                //** Layout for RMM Worksheet
                int RMM_column = 1;
                foreach (string info in RMM_info_item)
                {
                    xls_QN_sheet_RMM.Cells[1, RMM_column].Value = info;
                    xls_QN_sheet_RMM.Cells[1, RMM_column].Font.Bold = true;
                    xls_QN_sheet_RMM.Rows.AutoFit();
                    RMM_column++;
                }
                // Adding comment for Risk Classification Level
                xls_QN_sheet_RMM.Cells[1, 3].AddComment("W-Critical" + '\n' + "X-Major" + '\n' +
                    "Y-Moderate" + '\n' + "Z-Negligible");

                xls_QN_RMM_Ref.Cells[1, 1] = "QN ID";
                xls_QN_RMM_Ref.Cells[1, 1].Font.Bold = true;

                int row_num = 1; //** indicate the file numbers by iteration
                //***************** iterating File names into exporting info ******************************************//
                //foreach (string file_name in selected_files)
                for (int j_file = 0; j_file < file_numbers; j_file++)
                {
                    row_num++;// refresh the index to next column.
                    
                    //** Converting PDF into Text by using thrid-party library: iTextSharp
                    string Text_PDF = string.Empty;
                    
                    //** try to exporting PDF format into text **//
                    try
                    {
                        //OPERATOR for PDF to Text
                        Text_PDF = PDF_to_Text(selected_files[j_file]);
                    }
                    catch (Exception Error)
                    {
                        MessageBox.Show(Error.Message);
                    }

                    /** Exporting information into Excel sheet **/
                    //** Separating The whole string into several substring
                    /*Structure of the QN file:
                     * ------------ basic info -----------------
                     * --------Problem Description -------------
                     * ---------------Tasks---------------------
                     * ---------------Activities----------------
                     * ---------Detailed Items------------------
                     */
                    part_string = Text_PDF.Split(stringSeparates, StringSplitOptions.None); //** Split the whole string into several substring by seperating Tasks and Activities

                    //** Extracting Basic info from QN files. **//
                    int position1 = part_string[0].IndexOf("Plant:") + "Plant:".Length;

                    //** Filtering US8W into Excel
                    if (part_string[0].Substring(position1, 5) == " US8W")
                    {
                        //** Export Basic Information
                        ExcelBasicInfoTool(row_num, selected_files[j_file], part_string[0], BasicInfo, xls_QN_sheet,
                                           xls_QN_sheet_RMM, xls_QN_RMM_Ref);
  
                        //** Export Task Description 
                        // Recombine Tasks and Activities strings
                        string Task_Activities = String.Empty; //** Define the string of Tasks and Activiteis 
                        for (int index = 1; index < part_string.Length; index++)
                        {
                            Task_Activities = Task_Activities + part_string[index]; //** Re-assemble substrings into a whole string
                            //** NOTE: The reason is because for the following pages in QN files, there are multiple 
                            //**       Tasks and Activities on the headline. It is hard to process substrings. Therefore, re-combine into a whole string
                        }

                        // Filting for Tasks and description
                        /* Separating by 
                         * "_______________________________________________________________"
                         */
                        task_Activity_parts = Task_Activities.Split(delimiter5, StringSplitOptions.None);//** Seperating 

                        //** DETECTING RMM AS KEYWORD
                        //** Trying to BOLD & UNDERSTORE Keyword
                        ExcelRMM_RiskLevelTool(row_num, task_Activity_parts, xls_QN_sheet_RMM);
                        //ExcelRMMPaternMatch(row_num, task_Activity_parts, xls_QN_sheet_RMM);
                        ExcelRMM_pFMEA_Entry(row_num, selected_files[j_file], task_Activity_parts, xls_QN_sheet_RMM,xls_QN_RMM_Ref);
                        //** Exporting Task.QE and Activities into Excel
                        if (checkBox1.Checked == true)
                            ExcelFilterTool(row_num, selected_files[j_file], task_Activity_parts, xls_QN_sheet);
                    }// ELSE IF condition for filtering US8W
                    else
                    {
                        row_num--;//** If Plant is not "US8W", then go back to last column and continue this loop for another file 
                        continue; //** Continue for another loop to next file
                    }

                }//** END with iterating pdf files
                #endregion
                #region Formatting the Excel cells alignment
                //** Changing Format for more usabiliy
                xls_QN_sheet.Range[xls_QN_sheet.Cells[1, 1], xls_QN_sheet.Cells[row_num]].WrapText = true; //** wrap text
                xls_QN_sheet.Columns.AutoFit(); //** auto fit cells
                xls_QN_sheet.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop; //** align text into top

                xls_QN_sheet_RMM.Range[xls_QN_sheet_RMM.Cells[1, 1], xls_QN_sheet_RMM.Cells[row_num]].WrapText = true; //** wrap text
                xls_QN_sheet_RMM.Columns.AutoFit();
                xls_QN_sheet_RMM.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                xls_QN_RMM_Ref.Range[xls_QN_RMM_Ref.Cells[1, 1], xls_QN_RMM_Ref.Cells[row_num]].WrapText = true;
                xls_QN_RMM_Ref.Columns.AutoFit();
                xls_QN_RMM_Ref.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                
                button1.Enabled = false; //** Disable button untill next select
                #endregion
            }// ELSE IF-condition "US8M unchecked && US8W checked"
            #endregion
            #region Both CT and AMI checked
            else if (US8M.Checked == true && US8W.Checked == true)
            {
                #region Initialization
                //** Create new Worksheet for US8M and US8W
                Excel.Worksheet xls_QN_sheet;
                xls_QN_sheet = (Excel.Worksheet)xls_QN_workbook.Worksheets.get_Item(1);
                xls_QN_sheet.Name = "US8M_CT_QN_files";

                Excel.Worksheet xls_QN_sheet_US8W;
                xls_QN_sheet_US8W = (Excel.Worksheet)xls_QN_workbook.
                                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_sheet_US8W.Name = "US8W_AMI_QN_files";

                //** Create new Worksheet for RMM 
                Excel.Worksheet xls_QN_sheet_US8M_RMM;
                xls_QN_sheet_US8M_RMM = (Excel.Worksheet)xls_QN_workbook.
                                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_sheet_US8M_RMM.Name = "US8M_CT_QN_RMM";

                Excel.Worksheet xls_QN_sheet_US8W_RMM;
                xls_QN_sheet_US8W_RMM = (Excel.Worksheet)xls_QN_workbook.
                                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_sheet_US8W_RMM.Name = "US8W_AMI_QN_RMM";

                //** Create new Worksheet for CT RMM Reference information
                Excel.Worksheet xls_QN_US8M_RMM_Ref;
                xls_QN_US8M_RMM_Ref = (Excel.Worksheet)xls_QN_workbook.
                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_US8M_RMM_Ref.Name = "QN_files_CT_RMM_Reference";

                //** Create new Worksheet for AMI RMM Reference information
                Excel.Worksheet xls_QN_US8W_RMM_Ref;
                xls_QN_US8W_RMM_Ref = (Excel.Worksheet)xls_QN_workbook.
                        Worksheets.Add(After: xls_QN_workbook.Sheets[xls_QN_workbook.Sheets.Count]);
                xls_QN_US8W_RMM_Ref.Name = "QN_files_AMI_RMM_Reference";

                //** int info_items = listed_items.Length;
                //** Layout of matrix is in order of row 
                int column = 1; //Initialize the column number
                foreach (string info in listed_items_excel)
                {
                    xls_QN_sheet.Cells[1, column].Value = info;
                    xls_QN_sheet_US8W.Cells[1, column].Value = info;
                    xls_QN_sheet.Cells[1, column].Font.Bold = true;
                    xls_QN_sheet_US8W.Cells[1, column].Font.Bold = true; 
                    xls_QN_sheet.Rows.AutoFit();
                    xls_QN_sheet_US8W.Rows.AutoFit();
                    column++;
                }//** Export Information tilte into Excel sheet

                //** Layout for RMM Worksheet
                int RMM_column = 1;
                foreach (string info in RMM_info_item)
                {
                    xls_QN_sheet_US8M_RMM.Cells[1, RMM_column].Value = info;
                    xls_QN_sheet_US8M_RMM.Cells[1, RMM_column].Font.Bold = true;
                    xls_QN_sheet_US8M_RMM.Rows.AutoFit();
                    xls_QN_sheet_US8W_RMM.Cells[1, RMM_column].Value = info;
                    xls_QN_sheet_US8W_RMM.Cells[1, RMM_column].Font.Bold = true;
                    xls_QN_sheet_US8W_RMM.Rows.AutoFit();
                    RMM_column++;
                }
                // Adding comment for Risk Classification Level
                xls_QN_sheet_US8M_RMM.Cells[1, 3].AddComment("W-Critical" + '\n' + "X-Major" + '\n' +
                    "Y-Moderate" + '\n' + "Z-Negligible");
                xls_QN_sheet_US8W_RMM.Cells[1, 3].AddComment("W-Critical" + '\n' + "X-Major" + '\n' +
                    "Y-Moderate" + '\n' + "Z-Negligible");

                xls_QN_US8M_RMM_Ref.Cells[1, 1] = "QN ID";
                xls_QN_US8M_RMM_Ref.Cells[1, 1].Font.Bold = true;

                xls_QN_US8W_RMM_Ref.Cells[1, 1] = "QN ID";
                xls_QN_US8W_RMM_Ref.Cells[1, 1].Font.Bold = true;


                int row_num = 1; //** indicate the file numbers by iteration
                #endregion
                #region Checking Files one by one
                int US8M_row_num = 1; //** indicate the CT file numbers by iteration
                int US8W_row_num = 1; //** indicate the AMI file numbers by iteration
                //***************** iterating File names into exporting info ******************************************//
                Stopwatch sw = new Stopwatch();

                sw.Start();
                for (int k_file = 0; k_file < file_numbers; k_file++)
                {
                    //** Converting PDF into Text by using thrid-party library: iTextSharp
                    string Text_PDF = string.Empty;
                    
                    //** try to exporting PDF format into text **//
                    try
                    {
                        Text_PDF = PDF_to_Text(selected_files[k_file]);
                    }
                    catch (Exception Error)
                    {
                        MessageBox.Show(Error.Message);
                    }

                    /** Exporting information into Excel sheet **/
                    //** Separating The whole string into several substring
                    /*
                     * Structure of the QN file:
                     * ------------ basic info -----------------
                     * --------Problem Description -------------
                     * ---------------Tasks---------------------
                     * ---------------Activities----------------
                     * ---------Detailed Items------------------
                     */
                    part_string = Text_PDF.Split(stringSeparates, StringSplitOptions.None); //** Split the whole string into several substring by seperating Tasks and Activities

                    //** Extracting Basic info from QN files. **//
                    int position1 = part_string[0].IndexOf("Plant:") + "Plant:".Length;
                    //** Filtering US8M into Excel
                    #region For CT case
                    if (part_string[0].Substring(position1, 5) == " US8M")
                    {
                        US8M_row_num++;// refresh the index to next column.
                        //** Export Basic Information
                        ExcelBasicInfoTool(US8M_row_num, selected_files[k_file], part_string[0], BasicInfo, xls_QN_sheet,
                                           xls_QN_sheet_US8M_RMM,xls_QN_US8M_RMM_Ref);

                        //** Export Task Description 
                        // Recombine Tasks and Activities strings
                        string Task_Activities = String.Empty; //** Define the string of Tasks and Activiteis 
                        for (int index = 1; index < part_string.Length; index++)
                        {
                            Task_Activities = Task_Activities + part_string[index]; //** Re-assemble substrings into a whole string
                            //** NOTE: The reason is because for the following pages in QN files, there are multiple 
                            //**       Tasks and Activities on the headline. It is hard to process substrings. Therefore, re-combine into a whole string
                        }

                        // Filting for Tasks and description
                        /* Separating by 
                         * "_______________________________________________________________"
                         */
                        task_Activity_parts = Task_Activities.Split(delimiter5, StringSplitOptions.None);//** Seperating 

                        //** DETECTING RMM AS KEYWORD
                        ExcelRMM_RiskLevelTool(US8M_row_num, task_Activity_parts, xls_QN_sheet_US8M_RMM);
                        //ExcelRMMPaternMatch(row_num, task_Activity_parts, xls_QN_sheet_RMM);
                        ExcelRMM_pFMEA_Entry(US8M_row_num, selected_files[k_file], task_Activity_parts, xls_QN_sheet_US8M_RMM, xls_QN_US8M_RMM_Ref);
                        //** Operator for Task and Activity exporting
                        if (checkBox1.Checked == true)
                            ExcelFilterTool(row_num, selected_files[k_file], task_Activity_parts, xls_QN_sheet);
                    }//** END with US8M
                    #endregion
                    #region For AMI case
                    else if (part_string[0].Substring(position1, 5) == " US8W")
                    {
                        US8W_row_num++;
                        //** Export Basic Information
                        ExcelBasicInfoTool(US8W_row_num, selected_files[k_file], part_string[0], BasicInfo, xls_QN_sheet_US8W,
                                           xls_QN_sheet_US8W_RMM,xls_QN_US8W_RMM_Ref);

                        //** Export Task Description 
                        // Recombine Tasks and Activities strings
                        string Task_Activities = String.Empty; //** Define the string of Tasks and Activiteis 
                        for (int index = 1; index < part_string.Length; index++)
                        {
                            Task_Activities = Task_Activities + part_string[index]; //** Re-assemble substrings into a whole string
                            //** NOTE: The reason is because for the following pages in QN files, there are multiple 
                            //**       Tasks and Activities on the headline. It is hard to process substrings. Therefore, re-combine into a whole string
                        }

                        // Filting for Tasks and description
                        /* Separating by 
                            * "_______________________________________________________________"
                            */
                        task_Activity_parts = Task_Activities.Split(delimiter5, StringSplitOptions.None);//** Seperating 

                        //** DETECTING RMM AS KEYWORD
                        ExcelRMM_RiskLevelTool(US8W_row_num, task_Activity_parts, xls_QN_sheet_US8W_RMM);
                        //ExcelRMMPaternMatch(row_num, task_Activity_parts, xls_QN_sheet_RMM);
                        ExcelRMM_pFMEA_Entry(US8W_row_num, selected_files[k_file], task_Activity_parts, xls_QN_sheet_US8W_RMM,xls_QN_US8W_RMM_Ref);
                        //** Exporting Task.QE and Activities
                        if (checkBox1.Checked == true)
                            ExcelFilterTool(row_num, selected_files[k_file], task_Activity_parts, xls_QN_sheet_US8W);
                    }//** END with US8W
                    #endregion
                    else
                    {
                        MessageBox.Show("This file's plant is out of US8M and US8W", "ERROR", MessageBoxButtons.OK);
                        continue;
                    }
                
                }
                #endregion
                #region Formatting Excel alignment
                //** Changing Format for more usabiliy
                xls_QN_sheet.Range[xls_QN_sheet.Cells[1, 1], xls_QN_sheet.Cells[US8M_row_num]].WrapText = true; //** wrap text
                xls_QN_sheet.Columns.AutoFit(); //** auto fit cells
                xls_QN_sheet.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop; //** align text into top

                //** Changing another sheet Format for more usabiliy
                xls_QN_sheet_US8W.Range[xls_QN_sheet_US8W.Cells[1, 1], xls_QN_sheet_US8W.Cells[US8W_row_num]].WrapText = true; //** wrap text
                xls_QN_sheet_US8W.Columns.AutoFit(); //** auto fit cells
                xls_QN_sheet_US8W.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop; //** align text into top

                xls_QN_sheet_US8M_RMM.Range[xls_QN_sheet_US8M_RMM.Cells[1, 1], xls_QN_sheet_US8M_RMM.Cells[US8M_row_num]].WrapText = true; //** wrap text
                xls_QN_sheet_US8M_RMM.Columns.AutoFit();
                xls_QN_sheet_US8M_RMM.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                xls_QN_sheet_US8W_RMM.Range[xls_QN_sheet_US8W_RMM.Cells[1, 1], xls_QN_sheet_US8W_RMM.Cells[US8W_row_num]].WrapText = true; //** wrap text
                xls_QN_sheet_US8W_RMM.Columns.AutoFit();
                xls_QN_sheet_US8W_RMM.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                xls_QN_US8M_RMM_Ref.Range[xls_QN_US8M_RMM_Ref.Cells[1, 1], xls_QN_US8M_RMM_Ref.Cells[row_num]].WrapText = true;
                xls_QN_US8M_RMM_Ref.Columns.AutoFit();
                xls_QN_US8M_RMM_Ref.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                xls_QN_US8W_RMM_Ref.Range[xls_QN_US8W_RMM_Ref.Cells[1, 1], xls_QN_US8W_RMM_Ref.Cells[row_num]].WrapText = true;
                xls_QN_US8W_RMM_Ref.Columns.AutoFit();
                xls_QN_US8W_RMM_Ref.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;


                
                button1.Enabled = false; //** Disable button "" untill next select

                sw.Stop();
                #endregion
                MessageBox.Show(string.Format("Elapsed={0}", sw.Elapsed), "TEST", MessageBoxButtons.OK);
            }//End condition (US8M.Checked == true && US8W.Checked == true)
            #endregion
            #region No box checked
            else
            {
                MessageBox.Show(new Form() {TopMost = true}, "Please select Plant Type US8M (CT) / US8W (AMI)", 
                                            "Reminder", MessageBoxButtons.OKCancel);
                return;
            }
            #endregion
            MessageBox.Show(new Form() { TopMost = true }, "Work is done successfully.",
                                            "Work Status", MessageBoxButtons.OK); //** Make sure MessageBox comes the most top            
        }//** button1_click function quick filter
        
        //**********************************************Defined Self Data Parsing Tool Functions************************************************//
        //** Operator for converting PDF file into Text
        public string PDF_to_Text(string filename)
        {
            //** Using iTextSharp library converting PDF to text:
            //**      class iTextSharp.text.pdf.PdfReader
            //**      class iTextSharp.text.pdf.parser.ITextExtractionStrategy
            //**      class iTextSharp.text.pdf.parser.PdfTextExtractor

            string output = String.Empty;
            PdfReader pdfread = new PdfReader(filename);
            
            for (int index = 1; index <= pdfread.NumberOfPages; index++)
            {
                ITextExtractionStrategy singlestring = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
                String str = PdfTextExtractor.GetTextFromPage(pdfread, index, singlestring);

                str = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(str)));

                output = output + str; //** Assemble the contents of PDF into Text
            }
            pdfread.Close();
            return output;
        }

        //** Operator for data parsing RMM info
        public void ExcelBasicInfoTool(int index, string fileName, string input, string[] input_items, Excel.Worksheet input_worksheet_main,
                                       Excel.Worksheet input_worksheet_RMM, Excel.Worksheet input_worksheet_RMM_Ref)
        {
            #region Export Notification QN id
            int position1 = input.IndexOf(BasicInfo[0]) + BasicInfo[0].Length;//** Find the position of the last letter of the keyword 
            int position2 = input.IndexOf(BasicInfo[1]); //** Find the postion of the first letter of the keyword behind
            int info_length = position2 - position1; //** Find the position of the first letter in the contents
            input_worksheet_main.Cells[index, 1].Value = input.Substring(position1, info_length);//** Exporting to Excel Cell
            input_worksheet_main.Cells[index, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//** Align the contents inside the cell
            //input_worksheet_main.Cells[index, 1].AddComment(fileName);//** Adding filenames into Comment

            // Exporting filename into RMM sheet
            input_worksheet_RMM.Cells[index, 1].Value = fileName.Substring(fileName.Length-28,28);//** Exporting to Excel Cell _ FOR RMM
            input_worksheet_RMM.Cells[index, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//** Align the contents inside the cell _FOR RMM
            
             // Exporting QN id into RMM sheet
            input_worksheet_RMM.Cells[index, 2].Value = input.Substring(position1, info_length);//** Exporting to Excel Cell _ FOR RMM
            input_worksheet_RMM.Cells[index, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//** Align the contents inside the cell _FOR RMM
            // Exporting QN ID into RMM_Ref sheet
            input_worksheet_RMM_Ref.Cells[index, 1].Value = input.Substring(position1, info_length);//** Exporting to Excel Cell _ FOR RMM
            input_worksheet_RMM_Ref.Cells[index, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//** Align the contents inside the cell _FOR RMM

            #endregion

            #region Export Notification Date
            input_worksheet_RMM.Cells[index, 8].Value = DateExport(input);//** Exporting to Excel Cell
            input_worksheet_RMM.Cells[index, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;//** Align the contents inside the cell
            #endregion

            #region Export Material
            position1 = input.IndexOf(BasicInfo[2]) + BasicInfo[2].Length;
            position2 = input.IndexOf(BasicInfo[3]);
            info_length = position2 - position1;
            input_worksheet_RMM.Cells[index, 9].Value = input.Substring(position1, info_length);
            #endregion

            #region Export Plant
            position1 = input.IndexOf("Plant:") + "Plant:".Length;
            input_worksheet_main.Cells[index, 2].Value = input.Substring(position1, 5);
            #endregion

            #region Export Description
            position1 = input.IndexOf(BasicInfo[4]) + BasicInfo[4].Length;
            position2 = input.IndexOf(BasicInfo[5]);
            info_length = position2 - position1;
            //input_worksheet_main.Cells[index, 5].Value = input.Substring(position1, info_length);
            input_worksheet_RMM.Cells[index, 7].Value = input.Substring(position1, info_length);
            #endregion
        }

        //** Operator for Filtering Task.QE and Activity
        public void ExcelFilterTool(int rows_index, string filename, string[] input, Excel.Worksheet input_worksheet)
        {
            //** Initializing the index in the Excel sheet
            int column_task_index = 0; //Define the index for each task in Excel sheet
            int ac_index = 0; // Define a global index to connect Tasks with Activities
            int column_activity_index = 0; //Define the begin index for activities 
            char newLine = '\n'; //Define newline to seperate the strign into parts
            string[] subString; //Define string buffer for each seperated strings 

            #region Export Task.QE 
            for (int subIndex = 0; subIndex < input.Length; subIndex += 2)
            {
                subString = input[subIndex].Split(newLine); //** Seperating string line by line
                try
                {
                if (subString[1].Length > 1) //** The second line is not Empty
                {
                    if (subString[1].Substring(0, 4) == "Task") //** Check for Task
                    {
                        if (subString[2].Substring(10, 2) == "QE")//** Only care about QE
                        {
                            if (subIndex + 1 < input.Length)//** Make sure not the last part
                            {
                                //index buffer: 6 is the location right behind the basic infomation
                                input_worksheet.Cells[rows_index, 3 + column_task_index].Value = subString[1].Substring(0, 9) + '\n' + input[subIndex + 1]; //** Exporting the info into the cell
                                column_task_index++; //** Move to next column
                            }
                            else
                                break;//** break the loop because of the last part of the string 
                        }
                        else
                        {
                            if (subString[subString.Length - 6].Length < 8 || subString[subString.Length - 4].Length < 8)//** Consider the case there is no contents between Tasks and Activies
                            {
                                continue;
                            }
                            else if (subString[subString.Length - 6].Substring(10, 2) == "QE")//** Consdier the case the last Task part is "QE"
                            {
                                input_worksheet.Cells[rows_index, 3 + column_task_index].Value = subString[subString.Length - 7].Substring(0, 9) + '\n' + input[subIndex + 1];
                                column_task_index++;
                            }
                            else if (subString[subString.Length - 4].Substring(0, 8) == "Activity")//** Consider the case the last part is Activities
                            {
                                input_worksheet.Cells[rows_index, 2 + column_task_index].Value = "Activity" + '\n' + input[subIndex + 1];
                                column_task_index++;
                                break;
                            }
                            else
                                continue;
                        }
                    }
                    else
                    {
                        ac_index = subIndex;//** pass the index to global index for Activities
                        break; //** continue to search Activities
                    }
                }
                else
                {
                    if (subString[1].Length == 1)//** Consider the third line is Tasks
                    {
                        if (subString[2].Substring(0, 4) == "Task")
                        {
                            if (subString[3].Substring(10, 2) == "QE")
                            {
                                input_worksheet.Cells[rows_index, 3 + column_task_index].Value = subString[2].Substring(0, 9) + '\n' + input[subIndex + 1];
                                column_task_index++;
                            }
                            else
                            {
                                if (subString[subString.Length - 6].Length < 8 && subString[subString.Length - 4].Length < 8)
                                {
                                    continue;
                                }
                                else if (subString[subString.Length - 6].Substring(10, 2) == "QE")
                                {
                                    input_worksheet.Cells[rows_index, 3 + column_task_index].Value = subString[subString.Length - 7].Substring(0, 9) + '\n' + input[subIndex + 1];
                                    column_task_index++;
                                }
                                else if (subString[subString.Length - 4].Substring(0, 8) == "Activity")
                                {
                                    input_worksheet.Cells[rows_index, 3 + column_task_index].Value = "Activity" + '\n' + input[subIndex + 1];
                                    column_task_index++;
                                    break;
                                }
                                else
                                    continue;
                            }
                        }
                    }
                    else
                    {
                        ac_index = subIndex;
                        break;
                    }
                }
                }                          
                catch
                {
                        MessageBox.Show("The file: " + filename +'\n'+"Tasks are not properly exported to Excel. Please manually check it", "ERROR", MessageBoxButtons.OK);
                        continue;
                }
            }
            #endregion
            #region Export Activity 
            for (int activityIndex = ac_index; activityIndex < input.Length; activityIndex++)
            {
                subString = input[activityIndex].Split(newLine);
                try
                {
                if (subString[1].Length <= 2) //** If the second line does not have Actiivties
                {
                    if (subString.Length > 2) //** Meanwhile the substring haves contents
                    {
                        if (subString[2].Substring(0, 8) == "Activity")
                        {
                            if (activityIndex + 1 == input.Length)
                            {
                                break; //** The string is ending
                            }
                            else
                            {
                                input_worksheet.Cells[rows_index, column_task_index + 3 + column_activity_index].Value = "Activity" + '\n' + input[activityIndex + 1];
                                column_activity_index++;
                            }
                        }
                    }
                }
                else
                {
                    if (subString[1].Substring(0, 8) == "Activity")
                    {
                        if (activityIndex + 1 == input.Length)
                        {
                            break;
                        }
                        else
                        {
                            input_worksheet.Cells[rows_index, column_task_index + 3 + column_activity_index].Value = "Activity" + '\n' + input[activityIndex + 1];
                            column_activity_index++; // move to another substring
                        }
                    }
                }
                }
                catch
                {
                            MessageBox.Show("The file: " + filename + '\n'+ "Activities are not properly exported into Excel. Please manually check it", "ERROR", MessageBoxButtons.OK);
                            continue;
                }
            }//**
            #endregion
        }
             
        //**
        // Call ExcelRMMTool(...) -> ExcelRiskClassificationAnalysisTool(...) -> ExcelRiskLevelPatternMatch(...)
        //** Operator for determining the Part Risk Classficaition (W, X, Y, Z) -- Core Function
        public void ExcelRiskLevelPaternMatch(int index, string input, Excel.Worksheet input_worksheet)
        {               
            if (Regex.IsMatch(input,pattern_1))
            {
                input_worksheet.Cells[index, 3].Value = "W";
                Risk_classification_status = true;
            }
            else if (Regex.IsMatch(input, pattern_2))
            {
                input_worksheet.Cells[index, 3].Value = "X";
                Risk_classification_status = true;
            }
            else if (Regex.IsMatch(input, pattern_3))
            {
                input_worksheet.Cells[index, 3].Value = "Y";
                Risk_classification_status = true;
            }
            else if (Regex.IsMatch(input, pattern_4))
            {
                input_worksheet.Cells[index, 3].Value = "Z";
                Risk_classification_status = true;
            }
            else
            {
                input_worksheet.Cells[index, 3].Value = input;
            }
            //Formatting the Excel cell alignement
            input_worksheet.Cells[index, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        //** Operator for determining the Part Risk Classficaition (W, X, Y, Z) -- Handle Function 
        public void ExcelRiskClassificationAnalysisTool(int row_index, int column_index, string input, Excel.Worksheet input_worksheet)
        {
            #region Detect Risk Classification Level
            string[] part = input.Split('.');
            int index = 0;
            for (int num = 0; num < part.Length; num++)
            {
                if (Regex.IsMatch(part[num], pattern_risk) || Regex.IsMatch(part[num], pattern_risk_1) || Regex.IsMatch(part[num], pattern_risk_2))
                {
                    if (Risk_classification_status == false)
                    {
                        ExcelRiskLevelPaternMatch(row_index, part[num] + '.', input_worksheet);
                    }
                    else
                        break; //If the Risk Classification has been detected, then break
                }
                else
                    continue;
                index = num;
            }

            #endregion
        }

        //** Operator for RMM Analysis Tool (Part Risk Classification, RMM, DHF, FMEA/FEMA)
        public void ExcelRMM_RiskLevelTool(int index, string[] input, Excel.Worksheet input_worksheet)
        {
            int length = input.Length;
            int num_RMM = 0;
            
            for (int m_index = 0; m_index < length; m_index++)
            {
                if (Risk_classification_status == false)
                {
                    if (Regex.IsMatch(input[m_index], pattern_risk) || Regex.IsMatch(input[m_index], pattern_risk_1 )|| Regex.IsMatch(input[m_index], pattern_risk_2))
                    {
                        ExcelRiskClassificationAnalysisTool(index, num_RMM, input[m_index], input_worksheet);
                        Risk_classification_status = true; // Make sure when Risk Classification is determined just for once
                        break;
                    }
                    else
                    {
                        continue;
                    }
                }
                else
                    break;
            }
            if (input_worksheet.Cells[index, 3].Value == null)
                input_worksheet.Cells[index, 3].Value = "N/A";
            Risk_classification_status = false;
         }

        //**
        // - Risk Classification can be detected by regular expression
        // - RMM doc ID can be determined by C16U-0266, CP2H-0019, DHF204922, DHF204923
        // - RMM entry ID: 
        //      If cannot find RMM doc ID, then break;
        //      Else, export the whole string to the cell
        // - FMEA can be processed by this way: 
        //      1. Get the string of the whole FMEA or FEMA
        //      2. Else, find the DHF but not the DHF204922, DHF204923
        //
        //
        //
        public void ExcelRMM_pFMEA_Entry(int index,string filename, string[] input, Excel.Worksheet input_worksheet, Excel.Worksheet input_RMM_Ref)
        {
            int length = input.Length;
            #region RMM Doc ID and Entry ID
            for (int k_index = 0; k_index < length; k_index++)
            {
                if (Regex.IsMatch(input[k_index], pattern_RMM_file1))
                {
                    input_worksheet.Cells[index, 4].Value = "C16U-0266";
                    input_worksheet.Cells[index, 5].Value = input[k_index];
                }
                else if (Regex.IsMatch(input[k_index], pattern_RMM_file2))
                {
                    input_worksheet.Cells[index, 4].Value = "CP2H-0019";
                    input_worksheet.Cells[index, 5].Value = input[k_index];
                }
                else if (Regex.IsMatch(input[k_index], pattern_RMM_file3))
                {
                    input_worksheet.Cells[index, 4].Value = "DHF204922";
                    RMMEntryID(index, input[k_index], input_worksheet);
                    input_RMM_Ref.Cells[index, 2].Value = input[k_index];
                }
                else if (Regex.IsMatch(input[k_index], pattern_RMM_file4))
                {
                    input_worksheet.Cells[index, 4].Value = "DHF204923";
                    RMMEntryID(index, input[k_index], input_worksheet); 
                    input_RMM_Ref.Cells[index, 2].Value = input[k_index];
                }
                else
                    continue;
            }
            if (input_worksheet.Cells[index, 4].Value == null)
            {
                input_worksheet.Cells[index, 4].Value = "N/A";
                input_worksheet.Cells[index, 5].Value = "N/A";
            }
            if (input_worksheet.Cells[index, 5].Value == null)
            {
                input_worksheet.Cells[index, 5].Value = "N/A";
            }
            #endregion
            
            #region pFMEA Doc ID
            for (int l_index = 0; l_index < length; l_index++)
            {
                if (Regex.IsMatch(input[l_index], pattern_pFMEA))
                {
                    if (input[l_index].IndexOf("DHF") > 0 && (input[l_index].IndexOf("204922") < 0 || input[l_index].IndexOf("204923") < 0))
                    {
                        input_worksheet.Cells[index, 6].Value = input[l_index].Substring(input[l_index].IndexOf("DHF"), 10);
                        input_RMM_Ref.Cells[index, 3].Value = input[l_index];
                    }
                    //input_worksheet.Cells[index, 6].AddComment(input[l_index]);
                    else
                    {
                        input_worksheet.Cells[index, 6].Value = "N/A";
                        // input_worksheet.Cells[index,6].AddComment(input[l_index]);
                    }
                    break;
                }
                else if (input[l_index].IndexOf("DHF") > 0 && (input[l_index].IndexOf("204922") < 0 && input[l_index].IndexOf("204923") < 0))
                {
                    input_worksheet.Cells[index, 6].Value = input[l_index].Substring(input[l_index].IndexOf("DHF"), 10);
                    input_RMM_Ref.Cells[index, 3].Value = input[l_index];
                    break;
                }
            }
            if (input_worksheet.Cells[index, 6].Value == null)
                input_worksheet.Cells[index, 6].Value = "N/A";
            #endregion
        }
        
        //** Operator for RMM Entry ID Input for CT-NM.RMM.RMM style
        private string RMMEntry_pattern = "(RMM.[0-9][0-9][0-9])";
       

        public void RMMEntryID(int index, string input, Excel.Worksheet input_worksheet)
        {
            MatchCollection matches = Regex.Matches(input, RMMEntry_pattern);
            string previous_string;
            foreach (Match match in matches)
            {
                if (input_worksheet.Cells[index, 5].Value != null)
                {
                    previous_string = input_worksheet.Cells[index, 5].Value2.ToString();
                    // The only difference between this property and the Value property is that the Value2 property doesn’t use the Currency and Date data types. 
                    // You can return values formatted with these data types as floating-point numbers by using the Double data type.
                    if (previous_string.IndexOf(match.Value) >= 0)
                        continue;
                    else
                        input_worksheet.Cells[index, 5].Value = input_worksheet.Cells[index, 5].Value + '\n' + match.Value;
                }
                else
                    input_worksheet.Cells[index, 5].Value = input_worksheet.Cells[index, 5].Value + '\n' + match.Value;
            }
        }
            
        //** Operator for Date format transfer
        public string DateExport(string Input)
        {
            //** For QN Date not GDP Format 
            string Output = String.Empty;
            string Raw_Date = String.Empty;
            int position_notif = Input.IndexOf("Notification Date : ");
            Raw_Date = Input.Substring(position_notif + 20, 11);
            return Raw_Date;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            label1.AutoSize = true;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
        /*
        // GDP for Date Format, but cost efficience. 
        public string[] Month = {"JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL",
                                 "AUG", "SEP", "OCT", "NOV", "DEC"};
        public string DateExport(string Input)
        {
            //** For QN Date format GDP 
            string Output = String.Empty;
            string Raw_Date = String.Empty;
            int position_notif = Input.IndexOf("Notification Date : ");
            Raw_Date = Input.Substring(position_notif+24, 2); //** Get the string of the Month
           
            switch (Raw_Date)
            {
                case "01":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[0] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "02":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[1] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "03":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[2] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "04":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[3] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "05":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[4] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "06":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[5] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "07":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[6] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "08":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[7] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "09":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[8] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "10":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[9] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "11":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[10] + "-" + Input.Substring(position_notif + 27, 4);
                    break;
                case "12":
                    Output = Input.Substring(position_notif + 21, 2) + "-" + Month[11] + "-" + Input.Substring(position_notif + 27, 4);
                    break;                
            }
            return Output;
        }
        */

     } //** END class MainForm: Form
} //** END namespace QN_Excel_Tool


