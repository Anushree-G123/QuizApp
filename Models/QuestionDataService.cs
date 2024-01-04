namespace QuizApp.Models

{
    using OfficeOpenXml;
    using System.Linq;

    public class QuestionDataService : Iquestion
    {
        private string path = "C:\\Users\\anushree.gattigar\\Documents\\Quiz1.xlsx";

        //public List<Questions> quizQuestion = new List<Questions>();



        public List<Questions> GetQuestions()
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                List<Questions> quizQuestion = new List<Questions>();


                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    Questions question = new Questions();
                    question.id = Convert.ToInt32(worksheet.Cells[row, 1].Value); //  Id is in column 1
                    question.QuestionText = worksheet.Cells[row, 2].Value.ToString(); //  QuestionText is in column 2

                    //Options are in columns 3 to 6
                    question.Options = new List<string>
                {
                    worksheet.Cells[row, 3].Value?.ToString(),
                    worksheet.Cells[row, 4].Value?.ToString(),
                    worksheet.Cells[row, 5].Value?.ToString(),
                    worksheet.Cells[row, 6].Value?.ToString()
                };

                    question.CorrectOption = worksheet.Cells[row, 7].Value?.ToString(); //  CorrectAnswer is in column 7
                    question.CorrectAnswerIndex = Convert.ToInt32(worksheet.Cells[row, 8].Value); //  CorrectAnswerIndex is in column 8

                    quizQuestion.Add(question);


                }

                return quizQuestion;


            }

        }

        public Questions GetQuestionsById(int id)

        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                List<Questions> quizQuestion = new List<Questions>();


                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    Questions question = new Questions();
                    question.id = Convert.ToInt32(worksheet.Cells[row, 1].Value); //  Id is in column 1
                    question.QuestionText = worksheet.Cells[row, 2].Value.ToString(); //  QuestionText is in column 2

                    //Options are in columns 3 to 6
                    question.Options = new List<string>
                {
                    worksheet.Cells[row, 3].Value?.ToString(),
                    worksheet.Cells[row, 4].Value?.ToString(),
                    worksheet.Cells[row, 5].Value?.ToString(),
                    worksheet.Cells[row, 6].Value?.ToString()
                };

                    question.CorrectOption = worksheet.Cells[row, 7].Value?.ToString(); //  CorrectAnswer is in column 7
                    question.CorrectAnswerIndex = Convert.ToInt32(worksheet.Cells[row, 8].Value); //  CorrectAnswerIndex is in column 8

                    quizQuestion.Add(question);


                }


                return quizQuestion.Where(p => p.id == id).FirstOrDefault();
            }
        }

        public int GetCorrectAnswerIndex(int questionId)
        {
            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                List<Questions> quizQuestion = new List<Questions>();


                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    Questions questions = new Questions();
                    questions.id = Convert.ToInt32(worksheet.Cells[row, 1].Value); //  Id is in column 1
                    questions.QuestionText = worksheet.Cells[row, 2].Value.ToString(); //  QuestionText is in column 2

                    //Options are in columns 3 to 6
                    questions.Options = new List<string>
                {
                    worksheet.Cells[row, 3].Value?.ToString(),
                    worksheet.Cells[row, 4].Value?.ToString(),
                    worksheet.Cells[row, 5].Value?.ToString(),
                    worksheet.Cells[row, 6].Value?.ToString()
                };

                    questions.CorrectOption = worksheet.Cells[row, 7].Value?.ToString(); //  CorrectAnswer is in column 7
                    questions.CorrectAnswerIndex = Convert.ToInt32(worksheet.Cells[row, 8].Value); //  CorrectAnswerIndex is in column 8

                    quizQuestion.Add(questions);


                }


                var question = quizQuestion.FirstOrDefault(q => q.id == questionId);

                //return question.CorrectAnswerIndex;

                if (question != null)
                {
                    return question.CorrectAnswerIndex;


                }

                return -1;
            }
        }

        public void UpdateExcelFile(int questionIndex, int selectedOptionIndex,string options,string questionText)
        {
            FileInfo existingFile = new FileInfo("C:\\Users\\anushree.gattigar\\Documents\\Quiz1.xlsx");

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet != null)
                {
                    
                    int excelRowIndex = questionIndex + 1; // +1 to account for header row

                    // Assuming the selected option is written in column 7 (index 6)
                    worksheet.Cells[excelRowIndex, 7].Value = selectedOptionIndex;

                    package.Save(); // Save changes back to the Excel file
                }
            }
        }
    }
}
        
  





    



















    //public class QuestionDataService : Iquestion
    //{


    //    static List<Questions> questions;
    //    public Questions Questions;
    //    static QuestionDataService()
    //    {
    //        questions = new List<Questions> {



    //                new Questions{
    //                id = 1,
    //                QuestionText = "What is the capital of India?",

    //                Options=new List<string> {"New Delhi","Bengaluru","Chennai","Dehadrun"},
    //                CorrectOption="New Delhi",CorrectAnswerIndex=0},
    //                new Questions
    //                {
    //                    id=2,
    //                    QuestionText="What is the capital of Karnataka?",
    //                    Options=new List<string> {"manglore","banglore","Mysore","Hubbali"},
    //                CorrectOption="Banglore",CorrectAnswerIndex=1},
    //                new Questions
    //                {
    //                    id=3,
    //                    QuestionText="What is the National Animal of India?",
    //                    Options=new List<string> {"Lion","Monkey","Cheetha","Tiger"} , CorrectOption = "Tiger", CorrectAnswerIndex = 3},




    //            };
    //    }




        
        //public List<Questions> GetQuestions() { return questions; }

    //    public Questions GetQuestionsById(int id)
    //    {
    //        return questions.Where(p=>p.id== id).FirstOrDefault();
    //    }

    //    public int GetCorrectAnswerIndex(int questionId)
    //    {
            
    //        var question = questions.FirstOrDefault(q => q.id == questionId);
    //        //return question.CorrectAnswerIndex;

    //        if (question != null)
    //        {
    //            return question.CorrectAnswerIndex;
                
    //        }

    //        return -1;
    //    }
    //}
    //}




