using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Drawing;
using System.IO;

namespace MDM
{
    public class DataExporter
    {
        public DataExporter()
        {

        }

        /// <summary>
        /// Generate appropriate file name for checksheets file
        /// </summary>
        /// <param name="MoldNo"></param>
        /// <param name="PartNo"></param>
        /// <param name="Type">Type of checksheet: QA, IPQC or ENG</param>
        /// <returns></returns>
        /// 
        public string GenerateFileName(string MoldNo, string PartNo, string Type)
        {
            string FileName;
            FileName = Type + "_" + MoldNo + "_" + PartNo + ".xlsx";
            FileName = FileName.Replace(" ", "");
            //remove illegal characters, just in case
            FileName = string.Join("", FileName.Split(Path.GetInvalidFileNameChars()));
            return FileName;
        }

        //QA checksheet functions
        /// <summary>
        /// Add part info to QA checksheets and do some formatting.
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddQAInfo(ExcelWorksheet Sheet, DataTable table)
        {
            ReplaceFirstCellContent("<MoldNo>", table.Rows[0]["MoldNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<PartName>", table.Rows[0]["PartName"].ToString(), Sheet);
            ReplaceFirstCellContent("<PartNo>", table.Rows[0]["PartNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<DieNo>", table.Rows[0]["DieNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<DWGRev>", table.Rows[0]["DWGrev"].ToString(), Sheet);
            ReplaceFirstCellContent("<Material>", table.Rows[0]["Material"].ToString(), Sheet);
            ReplaceFirstCellContent("<His>", table.Rows[0]["His"].ToString(), Sheet);
            ReplaceFirstCellContent("<Cavity>", table.Rows[0]["NoOfCav"].ToString(), Sheet);
            ReplaceFirstCellContent("<Rev>", table.Rows[0]["RevNo"].ToString(), Sheet);
            string FilePath = table.Rows[0]["Illustration"].ToString();
            FileInfo ImageFile = new FileInfo(FilePath);
            if (ImageFile.Exists)
            {
                Image image = Image.FromFile(FilePath);
                int Width = 170 * image.Width / image.Height;
                if (Width > 340)
                {
                    Width = 340;
                }
                var Illus = Sheet.Drawings.AddPicture("Illus", image);
                Illus.SetPosition(140, 450);
                Illus.SetSize(Width, 170);
            }
            //Hide unused Cavity column
            for (int i = int.Parse(table.Rows[0]["NoOfCav"].ToString()) + 12; i < 44; i++)
            {
                if (i > 15)
                {
                    Sheet.Column(i).Hidden = true;
                }
            }
            //Dealing with special cavity designation
            if (table.Rows[0]["SpecialCavityLetter"].ToString() == "True")
            {
                String[] CavityLetter = table.Rows[0]["SpecialCavityList"].ToString().Split(',');
                int ActualCav = int.Parse(table.Rows[0]["NoOfCav"].ToString());
                if (int.Parse(table.Rows[0]["NoOfCav"].ToString()) > CavityLetter.Length)
                {
                    ActualCav = CavityLetter.Length;
                }

                for (int i = 0; i < ActualCav; i++)
                {
                    ReplaceFirstCellContent("<Cav>", "Cav " + CavityLetter[i], Sheet);
                }
            }
            else
            {
                for (int i = 1; i <= int.Parse(table.Rows[0]["NoOfCav"].ToString()); i++)
                {
                    ReplaceFirstCellContent("<Cav>", "Cav " + i.ToString(), Sheet);
                }
            }
            ReplaceAllCellContent("<Cav>", "", Sheet);
        }

        /// <summary>
        /// Add appearance checkpoints to QA checksheet.
        /// </summary>
        /// <param name="Sheet"></param> Sheet to add appearance checkpoints to
        /// <param name="table"></param> Table contain appearance checkpoints
        public void AddQAAppr(ExcelWorksheet Sheet, DataTable table)
        {

            //find rows and col address
            int FirstApprRow = FindCellRowAddress("<ApprItemNo>", Sheet);
            int ItemNoCol = FindCellColAddress("<ApprItemNo>", Sheet);
            int CheckContent = FindCellColAddress("<CheckContent>", Sheet);
            int Specifications = FindCellColAddress("<Specifications>", Sheet);
            int ApprTool = FindCellColAddress("<ApprTool>", Sheet);
            int ApprID = FindCellColAddress("<ApprID>", Sheet);

            int CurrentRow = FirstApprRow;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["QA"].ToString() == "True")
                {
                    Sheet.Cells[CurrentRow, ItemNoCol].Value = table.Rows[i]["ApprItemNo"].ToString();
                    Sheet.Cells[CurrentRow, CheckContent].Value = table.Rows[i]["CheckContent"].ToString();
                    Sheet.Cells[CurrentRow, Specifications].Value = table.Rows[i]["Specifications"].ToString();
                    Sheet.Cells[CurrentRow, ApprTool].Value = table.Rows[i]["ApprTool"].ToString();
                    Sheet.Cells[CurrentRow, ApprID].Value = "(ApprID)" + table.Rows[i]["ApprID"].ToString();

                    CurrentRow++;
                }
            }
            //hide unused apperance checkpoint rows
            for (int i = FirstApprRow + 29; i >= CurrentRow; i--)
            {
                Sheet.DeleteRow(i);
            }

        }

        /// <summary>
        /// Add Dimension checkpoints to QA checksheet
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddQADim(ExcelWorksheet Sheet, DataTable table)
        {

            int FirstDimRow = FindCellRowAddress("<DimItemNo>", Sheet);
            int ItemNoCol = FindCellColAddress("<DimItemNo>", Sheet);
            int GeoSymbol = FindCellColAddress("<GeometricTolerance>", Sheet);
            int Specs = FindCellColAddress("<DimSpecifications>", Sheet);
            int Tolerance1 = FindCellColAddress("<Tolerance1>", Sheet);
            int Tolerance2 = FindCellColAddress("<Tolerance2>", Sheet);
            int Upper = FindCellColAddress("<Upper>", Sheet);
            int Lower = FindCellColAddress("<Lower>", Sheet);
            int FaMax = FindCellColAddress("<FaAcceptMax>", Sheet);
            int FaMin = FindCellColAddress("<FaAcceptMin>", Sheet);
            int Position = FindCellColAddress("<Position>", Sheet);
            int Tool = FindCellColAddress("<DimTool>", Sheet);
            int DimID = FindCellColAddress("<DimID>", Sheet);
            int CurrentRow = FirstDimRow;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["QA"].ToString() == "True")
                {
                    Sheet.Cells[CurrentRow, ItemNoCol].Value = table.Rows[i]["DimItemNo"].ToString();
                    Sheet.Cells[CurrentRow, GeoSymbol].Value = table.Rows[i]["GeometricTolerance"].ToString();
                    Sheet.Cells[CurrentRow, Specs].Value = table.Rows[i]["Specifications"].ToString();
                    if (table.Rows[i]["Tolerance"].ToString().Contains("/"))
                    {
                        Sheet.Cells[CurrentRow, Tolerance1].Value = table.Rows[i]["Tolerance"].ToString().Split('/').First();
                        Sheet.Cells[CurrentRow, Tolerance2].Value = table.Rows[i]["Tolerance"].ToString().Split('/').Last();
                    }
                    else
                    {
                        Sheet.Cells[CurrentRow, Tolerance1].Value = table.Rows[i]["Tolerance"].ToString();
                        Sheet.Cells[CurrentRow, Tolerance2].Value = "";
                    }
                    Sheet.Cells[CurrentRow, Upper].Value = table.Rows[i]["Upper"].ToString();
                    Sheet.Cells[CurrentRow, Lower].Value = table.Rows[i]["Lower"].ToString();
                    Sheet.Cells[CurrentRow, FaMax].Value = table.Rows[i]["FaAcceptMax"].ToString();
                    Sheet.Cells[CurrentRow, FaMin].Value = table.Rows[i]["FaAcceptMin"].ToString();
                    Sheet.Cells[CurrentRow, Position].Value = table.Rows[i]["Position"].ToString();
                    Sheet.Cells[CurrentRow, Tool].Value = table.Rows[i]["DimTool"].ToString();
                    Sheet.Cells[CurrentRow, DimID].Value = "(DimID)" + table.Rows[i]["DimID"].ToString();
                    CurrentRow++;
                }
            }
            //Hide unused dimension checkpoint rows
            for (int i = FirstDimRow + 64; i >= CurrentRow; i--)
            {
                Sheet.DeleteRow(i);
            }
        }

        /// <summary>
        /// Add measurement results to QA checksheet
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="ResultTable"></param>
        public void AddResult(ExcelWorksheet Sheet, DataTable ResultTable, DataTable InfoTable)
        {
            //Adding time info
            DateTime MeasDate = (DateTime)ResultTable.Rows[0]["MeasureDate"];
            ReplaceAllCellContent("<InjectionDate>", ResultTable.Rows[0]["InjectionDate"].ToString(), Sheet);
            ReplaceAllCellContent("<MeasureDate>", MeasDate.ToString("dd/MM/yyyy"), Sheet);
            ReplaceAllCellContent("<Temp>", ResultTable.Rows[0]["Temp"].ToString(), Sheet);
            ReplaceAllCellContent("<Shift>", ResultTable.Rows[0]["Shift"].ToString(), Sheet);
            ReplaceAllCellContent("<Humi>", ResultTable.Rows[0]["Humid"].ToString(), Sheet);
            ReplaceAllCellContent("<MachineNo>", ResultTable.Rows[0]["MachineNo"].ToString(), Sheet);
            int.TryParse(InfoTable.Rows[0]["NoOfCav"].ToString(), out int NoOfCav);
            string[] CavityList = InfoTable.Rows[0]["SpecialCavityList"].ToString().Split(',');
            if (InfoTable.Rows[0]["SpecialCavityLetter"].ToString() == "True")
            {
                NoOfCav = CavityList.Length;
            }
            //Find columns addresses for each cavities
            int[] cols = new int[NoOfCav];
            for (int i = 0; i < NoOfCav; i++)
            {
                //find the columns of each cavity
                if (InfoTable.Rows[0]["SpecialCavityLetter"].ToString() == "True")
                {
                    cols[i] = FindCellColAddress("Cav " + CavityList[i].Trim(), Sheet);
                }
                else
                {
                    cols[i] = FindCellColAddress("Cav " + i.ToString(), Sheet);
                }
            }
            //Sweeping through all results in table
            for (int i = 0; i < ResultTable.Rows.Count; i++)
            {
                //Split the result list into result for each cavity
                string[] ResultList = ResultTable.Rows[i]["Valuelist"].ToString().Split(',');
                int row;
                //find the row of the result
                if (ResultTable.Rows[i]["DimID"].ToString() == "")
                {
                    row = FindCellRowAddress("(ApprID)" + ResultTable.Rows[i]["ApprID"].ToString(), Sheet);
                }
                else
                {
                    row = FindCellRowAddress("(DimID)" + ResultTable.Rows[i]["DimID"].ToString(), Sheet);
                }
                for (int j = 0; j < NoOfCav; j++)
                {
                    Sheet.Cells[row, cols[j]].Value = ResultList[j];
                }
                Sheet.Cells[row, FindCellColAddress("Judge", Sheet)].Value = ConvertON(ResultTable.Rows[i]["Judge"].ToString());
            }

        }

        //Eng checking standard functions
        /// <summary>
        /// Add part info to Eng checking standard.
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddENGInfo(ExcelWorksheet Sheet, DataTable table)
        {
            ReplaceFirstCellContent("<Customer>", table.Rows[0]["Customer"].ToString(), Sheet);
            ReplaceFirstCellContent("<Model>", table.Rows[0]["Model"].ToString(), Sheet);
            ReplaceFirstCellContent("<MoldNo>", table.Rows[0]["MoldNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<PartName>", table.Rows[0]["PartName"].ToString(), Sheet);
            ReplaceFirstCellContent("<PartNo>", table.Rows[0]["PartNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<DieNo>", table.Rows[0]["DieNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<DWGRev>", table.Rows[0]["DWGrev"].ToString(), Sheet);
            ReplaceFirstCellContent("<Material>", table.Rows[0]["Material"].ToString(), Sheet);
            ReplaceFirstCellContent("<His>", table.Rows[0]["His"].ToString(), Sheet);
            ReplaceFirstCellContent("<Cavity>", table.Rows[0]["NoOfCav"].ToString(), Sheet);
            ReplaceFirstCellContent("<Rev>", table.Rows[0]["RevNo"].ToString(), Sheet);
        }

        /// <summary>
        /// Add appearance checkpoints to ENG Checking standard
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddEngAppr(ExcelWorksheet Sheet, DataTable table)
        {
            int FirstApprRow = FindCellRowAddress("<ApprItemNo>", Sheet);
            int ItemNoCol = FindCellColAddress("<ApprItemNo>", Sheet);
            int CheckContent = FindCellColAddress("<CheckContent>", Sheet);
            int Specifications = FindCellColAddress("<Specifications>", Sheet);
            int ApprTool = FindCellColAddress("<ApprTool>", Sheet);
            int SupportingJig = FindCellColAddress("<SupportingJig>", Sheet);
            int QA = FindCellColAddress("<QA>", Sheet);
            int QASampleSize = FindCellColAddress("<QASampleSize>", Sheet);
            int QAFreq = FindCellColAddress("<QAFreq>", Sheet);
            int IPQC = FindCellColAddress("<IPQC>", Sheet);
            int IPQCSampleSize = FindCellColAddress("<IPQCSampleSize>", Sheet);
            int IPQCFreq = FindCellColAddress("<IPQCFreq>", Sheet);
            int OQC = FindCellColAddress("<OQC>", Sheet);
            int OQCSampleSize = FindCellColAddress("<OQCSampleSize>", Sheet);
            int OQCFreq = FindCellColAddress("<OQCFreq>", Sheet);
            int CurrentRow = FirstApprRow;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["ENG"].ToString() == "True")
                {
                    Sheet.Cells[CurrentRow, ItemNoCol].Value = table.Rows[i]["ApprItemNo"].ToString();
                    Sheet.Cells[CurrentRow, CheckContent].Value = table.Rows[i]["CheckContent"].ToString();
                    Sheet.Cells[CurrentRow, Specifications].Value = table.Rows[i]["Specifications"].ToString();
                    Sheet.Cells[CurrentRow, ApprTool].Value = table.Rows[i]["ApprTool"].ToString();
                    Sheet.Cells[CurrentRow, SupportingJig].Value = table.Rows[i]["SupportingJig"].ToString();
                    Sheet.Cells[CurrentRow, QA].Value = ConvertYN(table.Rows[i]["QA"].ToString());
                    Sheet.Cells[CurrentRow, QASampleSize].Value = table.Rows[i]["QASampleSize"].ToString();
                    Sheet.Cells[CurrentRow, QAFreq].Value = table.Rows[i]["QAFreq"].ToString();
                    Sheet.Cells[CurrentRow, IPQC].Value = ConvertYN(table.Rows[i]["IPQC"].ToString());
                    Sheet.Cells[CurrentRow, IPQCSampleSize].Value = table.Rows[i]["IPQCSampleSize"].ToString();
                    Sheet.Cells[CurrentRow, IPQCFreq].Value = table.Rows[i]["IPQCFreq"].ToString();
                    Sheet.Cells[CurrentRow, OQC].Value = ConvertYN(table.Rows[i]["OQC"].ToString());
                    Sheet.Cells[CurrentRow, OQCSampleSize].Value = table.Rows[i]["OQCSampleSize"].ToString();
                    Sheet.Cells[CurrentRow, OQCFreq].Value = table.Rows[i]["OQCFreq"].ToString();
                    CurrentRow++;
                    Sheet.Cells[CurrentRow, 56].Value = CurrentRow;
                }
            }
            //hide unused apperance checkpoint rows
            for (int i = FirstApprRow + 41; i >= CurrentRow; i--)
            {
                Sheet.DeleteRow(i);
            }
        }

        /// <summary>
        /// Add Dimension checkpoints to ENG checking standard
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddEngDim(ExcelWorksheet Sheet, DataTable table)
        {
            int FirstDimRow = FindCellRowAddress("<DimItemNo>", Sheet);
            int ItemNoCol = FindCellColAddress("<DimItemNo>", Sheet);
            int GeoSymbol = FindCellColAddress("<GeometricTolerance>", Sheet);
            int Specifications = FindCellColAddress("<DimSpecifications>", Sheet);
            int Tolerance1 = FindCellColAddress("<Tolerance1>", Sheet);
            int Tolerance2 = FindCellColAddress("<Tolerance2>", Sheet);
            int SupportingJig = FindCellColAddress("<DimSupportingJig>", Sheet);
            int Upper = FindCellColAddress("<Upper>", Sheet);
            int Lower = FindCellColAddress("<Lower>", Sheet);
            int FaMax = FindCellColAddress("<FaAcceptMax>", Sheet);
            int FaMin = FindCellColAddress("<FaAcceptMin>", Sheet);
            int Position = FindCellColAddress("<Position>", Sheet);
            int Tool = FindCellColAddress("<DimTool>", Sheet);
            int QA = FindCellColAddress("<DimQA>", Sheet);
            int QASampleSize = FindCellColAddress("<DimQASampleSize>", Sheet);
            int QAFreq = FindCellColAddress("<DimQAFreq>", Sheet);
            int IPQC = FindCellColAddress("<DimIPQC>", Sheet);
            int IPQCSampleSize = FindCellColAddress("<DimIPQCSampleSize>", Sheet);
            int IPQCFreq = FindCellColAddress("<DimIPQCFreq>", Sheet);
            int OQC = FindCellColAddress("<DimOQC>", Sheet);
            int OQCSampleSize = FindCellColAddress("<DimOQCSampleSize>", Sheet);
            int OQCFreq = FindCellColAddress("<DimOQCFreq>", Sheet);
            int CurrentRow = FirstDimRow;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["ENG"].ToString() == "True")
                {
                    Sheet.Cells[CurrentRow, ItemNoCol].Value = table.Rows[i]["DimItemNo"].ToString();
                    Sheet.Cells[CurrentRow, GeoSymbol].Value = table.Rows[i]["GeometricTolerance"].ToString();
                    Sheet.Cells[CurrentRow, Specifications].Value = table.Rows[i]["Specifications"].ToString();
                    if (table.Rows[i]["Tolerance"].ToString().Contains("/"))
                    {
                        Sheet.Cells[CurrentRow, Tolerance1].Value = table.Rows[i]["Tolerance"].ToString().Split('/').First();
                        Sheet.Cells[CurrentRow, Tolerance2].Value = table.Rows[i]["Tolerance"].ToString().Split('/').Last();
                    }
                    else
                    {
                        Sheet.Cells[CurrentRow, Tolerance1].Value = table.Rows[i]["Tolerance"].ToString();
                        Sheet.Cells[CurrentRow, Tolerance2].Value = "";
                    }
                    Sheet.Cells[CurrentRow, Upper].Value = table.Rows[i]["Upper"].ToString();
                    Sheet.Cells[CurrentRow, Lower].Value = table.Rows[i]["Lower"].ToString();
                    Sheet.Cells[CurrentRow, FaMax].Value = table.Rows[i]["FaAcceptMax"].ToString();
                    Sheet.Cells[CurrentRow, FaMin].Value = table.Rows[i]["FaAcceptMin"].ToString();
                    Sheet.Cells[CurrentRow, Position].Value = table.Rows[i]["Position"].ToString();
                    Sheet.Cells[CurrentRow, Tool].Value = table.Rows[i]["DimTool"].ToString();
                    Sheet.Cells[CurrentRow, SupportingJig].Value = table.Rows[i]["SupportingJig"].ToString();
                    Sheet.Cells[CurrentRow, QA].Value = ConvertYN(table.Rows[i]["QA"].ToString());
                    Sheet.Cells[CurrentRow, QASampleSize].Value = table.Rows[i]["QASampleSize"].ToString();
                    Sheet.Cells[CurrentRow, QAFreq].Value = table.Rows[i]["QAFreq"].ToString();
                    Sheet.Cells[CurrentRow, IPQC].Value = ConvertYN(table.Rows[i]["IPQC"].ToString());
                    Sheet.Cells[CurrentRow, IPQCSampleSize].Value = table.Rows[i]["IPQCSampleSize"].ToString();
                    Sheet.Cells[CurrentRow, IPQCFreq].Value = table.Rows[i]["IPQCFreq"].ToString();
                    Sheet.Cells[CurrentRow, OQC].Value = ConvertYN(table.Rows[i]["OQC"].ToString());
                    Sheet.Cells[CurrentRow, OQCSampleSize].Value = table.Rows[i]["OQCSampleSize"].ToString();
                    Sheet.Cells[CurrentRow, OQCFreq].Value = table.Rows[i]["OQCFreq"].ToString();
                    CurrentRow++;
                }
            }
            //Hide unused dimension checkpoint rows
            for (int i = FirstDimRow + 64; i >= CurrentRow; i--)
            {
                Sheet.DeleteRow(i);
            }
        }

        public void AddEngRev(ExcelWorksheet Sheet, DataTable table)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                ReplaceFirstCellContent("<Revision>" + i.ToString(), table.Rows[i]["Revision"].ToString(), Sheet);
                ReplaceFirstCellContent("<Reason>" + i.ToString(), table.Rows[i]["Reason"].ToString(), Sheet);
                ReplaceFirstCellContent("<Date>" + i.ToString(), table.Rows[i]["Date"].ToString(), Sheet);
                ReplaceFirstCellContent("<PrepairedBy>" + i.ToString(), table.Rows[i]["Prepaired"].ToString(), Sheet);
                ReplaceFirstCellContent("<ApprovedBy>" + i.ToString(), table.Rows[i]["Approved"].ToString(), Sheet);
            }
            int RevStart = FindCellRowAddress("<RevStart>", Sheet);
            int RevEnd = FindCellRowAddress("RevEnd", Sheet);
            Sheet.Cells[RevStart, 27, RevEnd, 31].Copy(Sheet.Cells[124, 27, 128, 31]);
            Sheet.Row(124).Height = 30;
            for (int i = table.Rows.Count; i < 5; i++)
            {
                ReplaceAllCellContent("<Revision>" + i.ToString(), "", Sheet);
                ReplaceAllCellContent("<Reason>" + i.ToString(), "", Sheet);
                ReplaceAllCellContent("<Date>" + i.ToString(), "", Sheet);
                ReplaceAllCellContent("<PrepairedBy>" + i.ToString(), "", Sheet);
                ReplaceAllCellContent("<ApprovedBy>" + i.ToString(), "", Sheet);
            }
        }

        //IPQC checksheet functions
        /// <summary>
        /// Add part info to IPQC checksheets and do some formatting
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddIPQCInfo(ExcelWorksheet Sheet, DataTable table)
        {
            ReplaceFirstCellContent("<MoldNo>", table.Rows[0]["MoldNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<PartName>", table.Rows[0]["PartName"].ToString(), Sheet);
            ReplaceFirstCellContent("<PartNo>", table.Rows[0]["PartNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<DieNo>", table.Rows[0]["DieNo"].ToString(), Sheet);
            ReplaceFirstCellContent("<DWGRev>", table.Rows[0]["DWGrev"].ToString(), Sheet);
            ReplaceFirstCellContent("<Material>", table.Rows[0]["Material"].ToString(), Sheet);
            ReplaceFirstCellContent("<His>", table.Rows[0]["His"].ToString(), Sheet);
            ReplaceFirstCellContent("<Cavity>", table.Rows[0]["NoOfCav"].ToString(), Sheet);
            ReplaceFirstCellContent("<Rev>", table.Rows[0]["RevNo"].ToString(), Sheet);
            string FilePath = table.Rows[0]["Illustration"].ToString();
            FileInfo ImageFile = new FileInfo(FilePath);
            if (ImageFile.Exists)
            {
                Image image = Image.FromFile(FilePath);
                int Width = 170 * image.Width / image.Height;
                if (Width > 340)
                {
                    Width = 340;
                }
                var Illus = Sheet.Drawings.AddPicture("Illus", image);
                Illus.SetPosition(140, 375);
                Illus.SetSize(Width, 170);
            }
            //Hide unused Cavity column
            for (int i = int.Parse(table.Rows[0]["NoOfCav"].ToString()) + 12; i < 44; i++)
            {
                if (i > 15)
                {
                    Sheet.Column(i).Hidden = true;
                }
            }
            //Dealing with special cavity designation
            if (table.Rows[0]["SpecialCavityLetter"].ToString() == "True")
            {
                String[] CavityLetter = table.Rows[0]["SpecialCavityList"].ToString().Split(',');
                int ActualCav = int.Parse(table.Rows[0]["NoOfCav"].ToString());
                if (int.Parse(table.Rows[0]["NoOfCav"].ToString()) > CavityLetter.Length)
                {
                    ActualCav = CavityLetter.Length;
                }

                for (int i = 0; i < ActualCav; i++)
                {
                    ReplaceFirstCellContent("<Cav>", "Cav " + CavityLetter[i], Sheet);
                }
            }
            else
            {
                for (int i = 1; i <= int.Parse(table.Rows[0]["NoOfCav"].ToString()); i++)
                {
                    ReplaceFirstCellContent("<Cav>", "Cav " + i.ToString(), Sheet);
                }
            }
            ReplaceAllCellContent("<Cav>", "", Sheet);
        }

        /// <summary>
        /// Add appearance checkpoints into IPQC checksheet
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddIPQCAppr(ExcelWorksheet Sheet, DataTable table)
        {
            //find rows and col address
            int FirstApprRow = FindCellRowAddress("<ApprItemNo>", Sheet);
            int ItemNoCol = FindCellColAddress("<ApprItemNo>", Sheet);
            int CheckContent = FindCellColAddress("<CheckContent>", Sheet);
            int Specifications = FindCellColAddress("<Specifications>", Sheet);
            int ApprTool = FindCellColAddress("<ApprTool>", Sheet);
            int ApprID = FindCellColAddress("<ApprID>", Sheet);
            int CurrentRow = FirstApprRow;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["IPQC"].ToString() == "True")
                {
                    Sheet.Cells[CurrentRow, ItemNoCol].Value = table.Rows[i]["ApprItemNo"].ToString();
                    Sheet.Cells[CurrentRow, CheckContent].Value = table.Rows[i]["CheckContent"].ToString();
                    Sheet.Cells[CurrentRow, Specifications].Value = table.Rows[i]["Specifications"].ToString();
                    Sheet.Cells[CurrentRow, ApprTool].Value = table.Rows[i]["ApprTool"].ToString();
                    Sheet.Cells[CurrentRow, ApprID].Value = "(ApprID)" + table.Rows[i]["ApprID"].ToString();
                    CurrentRow++;
                }
            }
            //hide unused apperance checkpoint rows
            for (int i = FirstApprRow + 29; i >= CurrentRow; i--)
            {
                Sheet.DeleteRow(i);
            }
        }

        /// <summary>
        /// Add Dimension checkpoints to IPQC checksheet
        /// </summary>
        /// <param name="Sheet"></param>
        /// <param name="table"></param>
        public void AddIPQCDim(ExcelWorksheet Sheet, DataTable table)
        {
            int FirstDimRow = FindCellRowAddress("<DimItemNo>", Sheet);
            int ItemNoCol = FindCellColAddress("<DimItemNo>", Sheet);
            int GeoSymbol = FindCellColAddress("<GeometricTolerance>", Sheet);
            int Specs = FindCellColAddress("<DimSpecifications>", Sheet);
            int Tolerance1 = FindCellColAddress("<Tolerance1>", Sheet);
            int Tolerance2 = FindCellColAddress("<Tolerance2>", Sheet);
            int Upper = FindCellColAddress("<Upper>", Sheet);
            int Lower = FindCellColAddress("<Lower>", Sheet);
            int FaMax = FindCellColAddress("<FaAcceptMax>", Sheet);
            int FaMin = FindCellColAddress("<FaAcceptMin>", Sheet);
            int Position = FindCellColAddress("<Position>", Sheet);
            int Tool = FindCellColAddress("<DimTool>", Sheet);
            int DimID = FindCellColAddress("<DimID>", Sheet);
            int CurrentRow = FirstDimRow;
            int ItemCount = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["IPQC"].ToString() == "True")
                {
                    ItemCount++;
                    Sheet.Cells[CurrentRow, ItemNoCol].Value = table.Rows[i]["DimItemNo"].ToString();
                    Sheet.Cells[CurrentRow, GeoSymbol].Value = table.Rows[i]["GeometricTolerance"].ToString();
                    Sheet.Cells[CurrentRow, Specs].Value = table.Rows[i]["Specifications"].ToString();
                    if (table.Rows[i]["Tolerance"].ToString().Contains("/"))
                    {
                        Sheet.Cells[CurrentRow, Tolerance1].Value = table.Rows[i]["Tolerance"].ToString().Split('/').First();
                        Sheet.Cells[CurrentRow, Tolerance2].Value = table.Rows[i]["Tolerance"].ToString().Split('/').Last();
                    }
                    else
                    {
                        Sheet.Cells[CurrentRow, Tolerance1].Value = table.Rows[i]["Tolerance"].ToString();
                        Sheet.Cells[CurrentRow, Tolerance2].Value = "";
                    }
                    Sheet.Cells[CurrentRow, Upper].Value = table.Rows[i]["Upper"].ToString();
                    Sheet.Cells[CurrentRow, Lower].Value = table.Rows[i]["Lower"].ToString();
                    Sheet.Cells[CurrentRow, FaMax].Value = table.Rows[i]["FaAcceptMax"].ToString();
                    Sheet.Cells[CurrentRow, FaMin].Value = table.Rows[i]["FaAcceptMin"].ToString();
                    Sheet.Cells[CurrentRow, Position].Value = table.Rows[i]["Position"].ToString();
                    Sheet.Cells[CurrentRow, Tool].Value = table.Rows[i]["DimTool"].ToString();
                    Sheet.Cells[CurrentRow, DimID].Value = "(DimID)" + table.Rows[i]["DimID"].ToString();
                    CurrentRow++;
                }
            }
            //Hide unused dimension checkpoint rows
            for (int i = FirstDimRow + 64; i >= CurrentRow; i--)
            {
                Sheet.DeleteRow(i);
            }
        }

        //Support functions
        /// <summary>
        /// Find and replace content in the fist match cell
        /// </summary>
        /// <param name="FindWhat"></param>
        /// <param name="ReplaceWith"></param>
        /// <param name="Sheet"></param>
        public void ReplaceFirstCellContent(String FindWhat, string ReplaceWith, ExcelWorksheet Sheet)
        {
            var query = from cell in Sheet.Cells["A1:AU130"]
                        where cell.Value?.ToString().Contains(FindWhat) == true
                        select cell;
            var FirstCell = query.FirstOrDefault();
            if (FirstCell != null)
            {
                FirstCell.Value = FirstCell.Value.ToString().Replace(FindWhat, ReplaceWith);
            }
        }

        /// <summary>
        /// Find and replace content in every cell in sheet
        /// </summary>
        /// <param name="FindWhat"></param>
        /// <param name="ReplaceWith"></param>
        /// <param name="Sheet"></param>
        public void ReplaceAllCellContent(String FindWhat, string ReplaceWith, ExcelWorksheet Sheet)
        {
            var query = from cell in Sheet.Cells["A1:AT130"]
                        where cell.Value?.ToString().Contains(FindWhat) == true
                        select cell;
            foreach (var Cell in query)
            {
                Cell.Value = Cell.Value.ToString().Replace(FindWhat, ReplaceWith);
            }
        }

        /// <summary>
        /// Find and replace in a single row
        /// </summary>
        /// <param name="Row"></param>
        /// <param name="FindWhat"></param>
        /// <param name="ReplaceWith"></param>
        /// <param name="Sheet"></param>
        public void ReplaceInRow(int Row, string FindWhat, string ReplaceWith, ExcelWorksheet Sheet)
        {
            var query = from cell in Sheet.Cells[Row, 1, Row, 52]
                        where cell.Value?.ToString().Contains(FindWhat) == true
                        select cell;
            var FirstCell = query.FirstOrDefault();
            if (FirstCell != null)
            {
                FirstCell.Value = FirstCell.Value.ToString().Replace(FindWhat, ReplaceWith);
            }
        }

        /// <summary>
        /// Find and return the row address of the cell contain string
        /// </summary>
        /// <param name="FindWhat"></param>
        /// <param name="ReplaceWith"></param>
        /// <param name="Sheet"></param>
        public int FindCellRowAddress(String FindWhat, ExcelWorksheet Sheet)
        {
            var query = from cell in Sheet.Cells["A1:AU130"]
                        where cell.Value?.ToString().Contains(FindWhat) == true
                        select cell.Start.Row;
            if (query != null && query.FirstOrDefault() > 0)
            {
                return query.FirstOrDefault();
            }
            return 121;//if not found
        }

        /// <summary>
        /// Find and return the column address of the cell contain string
        /// </summary>
        /// <param name="FindWhat"></param>
        /// <param name="Sheet"></param>
        /// <returns></returns>
        public int FindCellColAddress(String FindWhat, ExcelWorksheet Sheet)
        {
            var query = from cell in Sheet.Cells["A1:AU130"]
                        where cell.Value?.ToString().Contains(FindWhat) == true
                        select cell.Start.Column;
            if (query != null && query.FirstOrDefault() > 0)
            {
                return query.FirstOrDefault();
            }
            return 60;
        }

        /// <summary>
        /// Convert from True/False to Y/N
        /// </summary>
        /// <param name="Source"></param>
        /// <returns></returns>
        public string ConvertYN(string Source)
        {
            String Result = "";
            if (Source == "True")
            {
                Result = "Y";
            }
            if (Source == "False")
            {
                Result = "N";
            }
            return Result;
        }

        /// <summary>
        /// Convert from True/False to O/N
        /// </summary>
        /// <param name="Source"></param>
        /// <returns></returns>
        public string ConvertON(string Source)
        {
            if (Source == "True")
            {
                return "O";
            }
            return "N";
        }

        //Functions for measurement datatable
        /// <summary>
        /// Create measurement result table
        /// </summary>
        /// <param name="InfoTable"></param>
        /// <param name="ApprTable"></param>
        /// <param name="DimTable"></param>
        /// <param name="Batch"></param>
        /// <param name="InjectionDate"></param>
        /// <param name="MeasuredBy"></param>
        /// <returns></returns>
        public DataTable CreateBatchTable(DataTable InfoTable, DataTable ApprTable, DataTable DimTable, int Batch, string InjectionDate, string MeasuredBy, string Select)
        {
            //Column list
            string[,] ColumnsList = new string[,] {
                {"System.Int32", "MeasID"}, { "System.Int32", "ApprID" }, {"System.Int32","DimID" },
                {"System.String","MeasuredBy" }, {"System.String", "ValueList" },  {"System.String", "ToolID" },
                {"System.String","InjectionDate" }, { "System.Int32", "Batch" }, { "System.DateTime","MeasureDate"},
                {"System.Int32","ID" }, {"System.String", "Tool" }, {"System.String", "Shift" },
                {"System.String", "Temp" }, {"System.String", "Humid" },
                { "System.String", "MachineNo" }, {"System.String", "Note"}, {"System.Boolean", "Judge" }
            };

            //Add columns to table
            DataTable table = new DataTable();
            DataColumn column;
            DataRow row;
            for (int i = 0; i < ColumnsList.GetLength(0); i++)
            {
                column = new DataColumn
                {
                    DataType = Type.GetType(ColumnsList[i, 0]),
                    ColumnName = ColumnsList[i, 1]
                };
                table.Columns.Add(column);
            }
            //Limit length of data
            table.Columns["ToolID"].MaxLength = 10;
            table.Columns["InjectionDate"].MaxLength = 10;
            table.Columns["MeasuredBy"].MaxLength = 10;
            //Add  measurement data to table
            for (int i = 0; i < ApprTable.Rows.Count; i++)
            {
                row = table.NewRow();
                row["Batch"] = Batch;
                row["InjectionDate"] = InjectionDate;
                row["ApprID"] = ApprTable.Rows[i]["ApprID"].ToString();
                row["ID"] = InfoTable.Rows[0]["ID"];
                row["Tool"] = ApprTable.Rows[i]["ApprTool"].ToString();
                if (Select == "all" || (Select == "QA" && ApprTable.Rows[i]["QA"].ToString() == "True")
                    || (Select == "IPQC" && ApprTable.Rows[i]["IPQC"].ToString() == "True"))
                {
                    table.Rows.Add(row);
                }

            }
            for (int i = 0; i < DimTable.Rows.Count; i++)
            {
                row = table.NewRow();
                row["Batch"] = Batch;
                row["InjectionDate"] = InjectionDate;
                row["DimID"] = DimTable.Rows[i]["DimID"].ToString();
                row["ID"] = InfoTable.Rows[0]["ID"];
                row["Tool"] = DimTable.Rows[i]["DimTool"].ToString();
                if (Select == "all" || (Select == "QA" && DimTable.Rows[i]["QA"].ToString() == "True")
                    || (Select == "IPQC" && DimTable.Rows[i]["IPQC"].ToString() == "True"))
                {
                    table.Rows.Add(row);
                }
            }
            //Add info data to the first row of table
            table.Rows[0]["MeasureDate"] = DateTime.Now;
            table.Rows[0]["InjectionDate"] = InjectionDate;
            table.Rows[0]["MeasuredBy"] = MeasuredBy;
            //Add remaining info to columns
            return FillInfoToResult(InfoTable, ApprTable, DimTable, table);
        }

        /// <summary>
        /// Fill missing information into result table for display.
        /// </summary>
        /// <param name="InfoTable"></param>
        /// <param name="ApprTable"></param>
        /// <param name="DimTable"></param>
        /// <param name="ResultTable"></param>
        /// <returns></returns>
        public DataTable FillInfoToResult(DataTable InfoTable, DataTable ApprTable, DataTable DimTable, DataTable ResultTable)
        {
            DataTable table = ResultTable.Copy();
            DataColumn column;
            //Add missing column to table
            string[,] ColumnsList = new string[,] {
                {"System.String", "ItemNo" }, {"System.String","GeoSymbol"},
                {"System.String", "Specifications"}, {"System.String", "Tolerance"},
                {"System.String", "Range"}, {"System.String", "Position"},
                {"System.String", "FaAcceptMin"},{"System.String", "FaAcceptMax"}
                };
            for (int i = 0; i < ColumnsList.GetLength(0); i++)
            {
                column = new DataColumn
                {
                    DataType = Type.GetType(ColumnsList[i, 0]),
                    ColumnName = ColumnsList[i, 1]
                };
                table.Columns.Add(column);
            }
            //Add Columns for each cavity
            int NoOfCav = int.Parse(InfoTable.Rows[0]["NoOfCav"].ToString());
            string[] SpecialCavityList = InfoTable.Rows[0]["SpecialCavityList"].ToString().Split(',');
            if (SpecialCavityList.Length < NoOfCav && InfoTable.Rows[0]["SpecialCavityLetter"].ToString() == "True")
            {
                NoOfCav = SpecialCavityList.Length;
            }

            for (int i = 0; i < NoOfCav; i++)
            {
                column = new DataColumn
                {
                    DataType = Type.GetType("System.String")
                };
                if (InfoTable.Rows[0]["SpecialCavityLetter"].ToString() == "True")
                {
                    column.ColumnName = "Cavity " + SpecialCavityList[i];
                }
                else
                {
                    column.ColumnName = "Cavity " + (i + 1).ToString();
                }
                table.Columns.Add(column);
            }
            //Copy data from ValueList to Cavity column
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string[] ValueList = table.Rows[i]["ValueList"].ToString().Split(',');

                for (int j = 0; j < ValueList.Length; j++)
                {
                    table.Rows[i][j + table.Columns.Count - ValueList.Length] = ValueList[j];
                }
            }
            //Add info to columns
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (table.Rows[i]["DimID"].ToString() == "")
                {
                    for (int j = 0; j < ApprTable.Rows.Count; j++)
                    {
                        if (table.Rows[i]["ApprID"].ToString() == ApprTable.Rows[j]["ApprID"].ToString())
                        {
                            table.Rows[i]["Specifications"] = ApprTable.Rows[j]["CheckContent"];
                            table.Rows[i]["ItemNo"] = ApprTable.Rows[j]["ApprItemNo"].ToString().Trim();
                            table.Rows[i]["Tolerance"] = ApprTable.Rows[j]["Specifications"];
                        }
                    }
                }
                if (table.Rows[i]["ApprID"].ToString() == "")
                {
                    for (int j = 0; j < DimTable.Rows.Count; j++)
                    {
                        if (table.Rows[i]["DimID"].ToString() == DimTable.Rows[j]["DimID"].ToString())
                        {
                            table.Rows[i]["ItemNo"] = DimTable.Rows[j]["DimItemNo"];
                            table.Rows[i]["GeoSymbol"] = DimTable.Rows[j]["GeometricTolerance"];
                            table.Rows[i]["Specifications"] = DimTable.Rows[j]["GeometricTolerance"].ToString() + DimTable.Rows[j]["Specifications"];
                            table.Rows[i]["Tolerance"] = DimTable.Rows[j]["Tolerance"];
                            table.Rows[i]["Position"] = DimTable.Rows[j]["Position"];
                            table.Rows[i]["Range"] = DimTable.Rows[j]["Lower"] + "~" + DimTable.Rows[j]["Upper"];
                            table.Rows[i]["FaAcceptMin"] = DimTable.Rows[j]["FaAcceptMin"];
                            table.Rows[i]["FaAcceptMax"] = DimTable.Rows[j]["FaAcceptMax"];
                        }
                    }
                }
            }

            return table;
        }

    }
}
