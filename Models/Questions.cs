
namespace QuizApp.Models
{
    public class Questions
    {
        public int id { get; set; }
        public string QuestionText { get; set; }
        public List<string> Options { get; set; }

        public string CorrectOption { get; set; }

        public int CorrectAnswerIndex { get; set; }

        
    }
}
