using System.Data;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using InfinityMesh.Models;
using IronXL;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;


namespace InfinityMesh.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment Environment;
        public HomeController(IHostingEnvironment _environment)
        {
            Environment = _environment;
        }
        public IActionResult Index()
        {
            //------------------------ Otvara override
            WorkBook wbb = WorkBook.Load("override.xlsx");
            WorkSheet ws = wbb.GetWorkSheet("Sheet0");
            //------------------------- Tabele i podaci
            //------Naslovi iznad tabela
            ViewData["State1"] = ws["A1"].Value.ToString();
            ViewData["State2"] = ws["A2"].Value.ToString();
            ViewData["State3"] = ws["A3"].Value.ToString();
            //------ Tabela 1
            ViewData["T1Name1"] = ws["C1"].Value.ToString();
            ViewData["T1Name2"] = ws["E1"].Value.ToString();
            ViewData["T1Name3"] = ws["G1"].Value.ToString();
            ViewData["T1Name4"] = ws["I1"].Value.ToString();
            ViewData["T1Name5"] = ws["K1"].Value.ToString();
            ViewData["T1Votes1"] = ws["B1"].IntValue;
            ViewData["T1Votes2"] = ws["D1"].IntValue;
            ViewData["T1Votes3"] = ws["F1"].IntValue;
            ViewData["T1Votes4"] = ws["H1"].IntValue;
            ViewData["T1Votes5"] = ws["J1"].IntValue;
            //------ Tabela 2
            ViewData["T2Name1"] = ws["C2"].Value.ToString();
            ViewData["T2Name2"] = ws["E2"].Value.ToString();
            ViewData["T2Name3"] = ws["G2"].Value.ToString();
            ViewData["T2Name4"] = ws["I2"].Value.ToString();
            ViewData["T2Name5"] = ws["K2"].Value.ToString();
            ViewData["T2Votes1"] = ws["B2"].IntValue;
            ViewData["T2Votes2"] = ws["D2"].IntValue;
            ViewData["T2Votes3"] = ws["F2"].IntValue;
            ViewData["T2Votes4"] = ws["H2"].IntValue;
            ViewData["T2Votes5"] = ws["J2"].IntValue;
            //------ Tabela 3
            ViewData["T3Name1"] = ws["C3"].Value.ToString();
            ViewData["T3Name2"] = ws["E3"].Value.ToString();
            ViewData["T3Name3"] = ws["G3"].Value.ToString();
            ViewData["T3Name4"] = ws["I3"].Value.ToString();
            ViewData["T3Name5"] = ws["K3"].Value.ToString();
            ViewData["T3Votes1"] = ws["B3"].IntValue;
            ViewData["T3Votes2"] = ws["D3"].IntValue;
            ViewData["T3Votes3"] = ws["F3"].IntValue;
            ViewData["T3Votes4"] = ws["H3"].IntValue;
            ViewData["T3Votes5"] = ws["J3"].IntValue;
            return View();
        }

        [HttpPost]
        public IActionResult Index(IFormFile postedFile)
        {
            if (postedFile != null)
            {
                string path = Path.Combine(this.Environment.WebRootPath, "Uploads");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string fileName = Path.GetFileName(postedFile.FileName);
                string filePath = Path.Combine(path, fileName);
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    postedFile.CopyTo(stream);
                }
                WorkBook workbook = WorkBook.LoadCSV(filePath, fileFormat: ExcelFileFormat.XLSX, ListDelimiter: ",");
                workbook.SaveAs("CsvToExcelConversion.xlsx");
                WorkBook wb = WorkBook.Load("CsvToExcelConversion.xlsx");
                WorkSheet wss = wb.GetWorkSheet("Sheet0");
                //---------------------- Prvi red
                string[] name = new string[15];
                string[] state = new string[3];
                int[] votes = new int[15];
                state[0] = wss["A1"].Value.ToString();
                votes[0] = wss["B1"].IntValue;
                name[0] = wss["C1"].Value.ToString();
                votes[1] = wss["D1"].IntValue;
                name[1] = wss["E1"].Value.ToString();
                votes[2] = wss["F1"].IntValue;
                name[2] = wss["G1"].Value.ToString();
                votes[3] = wss["H1"].IntValue;
                name[3] = wss["I1"].Value.ToString();
                votes[4] = wss["J1"].IntValue;
                name[4] = wss["K1"].Value.ToString();
                ///------------------ Drugi red
                state[1] = wss["A2"].Value.ToString();
                votes[5] = wss["B2"].IntValue;
                name[5] = wss["C2"].Value.ToString();
                votes[6] = wss["D2"].IntValue;
                name[6] = wss["E2"].Value.ToString();
                votes[7] = wss["F2"].IntValue;
                name[7] = wss["G2"].Value.ToString();
                votes[8] = wss["H2"].IntValue;
                name[8] = wss["I2"].Value.ToString();
                votes[9] = wss["J2"].IntValue;
                name[9] = wss["K2"].Value.ToString();
                ///------------------ treci red
                state[2] = wss["A3"].Value.ToString();
                votes[10] = wss["B3"].IntValue;
                name[10] = wss["C3"].Value.ToString();
                votes[11] = wss["D3"].IntValue;
                name[11] = wss["E3"].Value.ToString();
                votes[12] = wss["F3"].IntValue;
                name[12] = wss["G3"].Value.ToString();
                votes[13] = wss["H3"].IntValue;
                name[13] = wss["I3"].Value.ToString();
                votes[14] = wss["J3"].IntValue;
                name[14] = wss["K3"].Value.ToString();
                //-------------------------Provjera imena
                for (int i = 0; i < 15; i++)
                {
                    if(name[i] == "DT") { name[i] = "Donald Trump"; }
                    else if (name[i] == "HC") { name[i] = "Hillary Clinton"; }
                    else if (name[i] == "JB") { name[i] = "Joe Biden"; }
                    else if (name[i] == "JFK") { name[i] = "John F. Kennedy"; }
                    else if (name[i] == "JR") { name[i] = "Jack Randall"; }
                }
                if (name[4] != "JR" || name[4] != "JFK" || name[4] != "JB" || name[4] != "HC") { name[i] = "Donald Trump"; }
                else if (name[9] != "JR" || name[9] != "JFK" || name[9] != "JB" || name[9] != "DT") { name[9] = "Hillary Clinton"; }
                else if (name[i] != "JR" || name[i] != "JFK" || name[i] != "DT" || name[i] != "HC") { name[i] = "Joe Biden"; }
                else if (name[i] != "JR" || name[i] != "DT" || name[i] != "JB" || name[i] != "HC") { name[i] = "John F. Kennedy"; }
                else if (name[i] != "DT" || name[i] != "JFK" || name[i] != "JB" || name[i] != "HC") { name[i] = "Jack Randall"; }
                //------------------------- Otvaranje override
                WorkBook wbb = WorkBook.Load("override.xlsx");
                WorkSheet ws = wbb.GetWorkSheet("Sheet0");
                //----prvi red
                ws["A1"].Value = state[0];
                ws["B1"].Value = votes[0];
                ws["C1"].Value = name[0];
                ws["D1"].Value = votes[1];
                ws["E1"].Value = name[1];
                ws["F1"].Value = votes[2];
                ws["G1"].Value = name[2];
                ws["H1"].Value = votes[3];
                ws["I1"].Value = name[3];
                ws["J1"].Value = votes[4];
                ws["K1"].Value = name[4];
                //----drugi red
                ws["A2"].Value = state[1];
                ws["B2"].Value = votes[5];
                ws["C2"].Value = name[5];
                ws["D2"].Value = votes[6];
                ws["E2"].Value = name[6];
                ws["F2"].Value = votes[7];
                ws["G2"].Value = name[7];
                ws["H2"].Value = votes[8];
                ws["I2"].Value = name[8];
                ws["J2"].Value = votes[9];
                ws["K2"].Value = name[9];
                //---- treci red
                ws["A3"].Value = state[2];
                ws["B3"].Value = votes[10];
                ws["C3"].Value = name[10];
                ws["D3"].Value = votes[11];
                ws["E3"].Value = name[11];
                ws["F3"].Value = votes[12];
                ws["G3"].Value = name[12];
                ws["H3"].Value = votes[13];
                ws["I3"].Value = name[13];
                ws["J3"].Value = votes[14];
                ws["K3"].Value = name[14];
                wbb.SaveAs("override.xlsx");
            }
            return View();
        }
    }
}