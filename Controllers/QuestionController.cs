using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using QuizApp.Models;

namespace QuizApp.Controllers
{
    public class QuestionController : Controller
    {
        private readonly Iquestion service;

        public List<Questions> questions { get; set; }
        public QuestionController(Iquestion service)
        {
            this.service = service;
        }


        public IActionResult Index()
        {

            questions = service.GetQuestions();
            return View(questions);
        }

        public IActionResult Details(int id)
        {

            Questions questions = service.GetQuestionsById(id);
            return View(questions);
        }

        [HttpPost]
        public IActionResult CheckAnswer(int questionId, int selectedOptionIndex,string options,string questionText)
        {
            
            int correctAnswerIndex = service.GetCorrectAnswerIndex(questionId);
            


            FileInfo existingFile = new FileInfo("C:\\Users\\anushree.gattigar\\Documents\\Book2.xlsx");

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {

                    int excelRowIndex = questionId + 1; // +1 to account for header row

                    //  the selected option is written in column 
                    worksheet.Cells[excelRowIndex, 3].Value = selectedOptionIndex;
                    worksheet.Cells[excelRowIndex, 1].Value = questionId;
                    worksheet.Cells[excelRowIndex, 4].Value = options;
                    worksheet.Cells[excelRowIndex, 2].Value = questionText;

                    package.Save();
                }
            }

            // Compare selectedOptionIndex with correctAnswerIndex and return the result
            if (selectedOptionIndex == correctAnswerIndex)
            {
                ViewBag.Result = "You got it Correct!";
            }
            else
            {
                ViewBag.Result = "Incorrect:(";
            }


            return View(questions);
            // return RedirectToAction("Index");
        }
    }
        }
 


                //[HttpPost]
                //public ActionResult SubmitAnswer(int questionIndex, int selectedOptionIndex)
                //{


                //    service.UpdateExcelFile(questionIndex, selectedOptionIndex);


                //    return RedirectToAction("Index"); 
                //}

           
