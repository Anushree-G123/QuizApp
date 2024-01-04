namespace QuizApp.Models
{
    public interface Iquestion
    {
        //public List<Questions> GetQuizFromExcel(string filepath);
        public List<Questions> GetQuestions();
        public Questions GetQuestionsById(int id);

        public int GetCorrectAnswerIndex(int questionId);

        public void UpdateExcelFile(int questionIndex, int selectedOptionIndex,string options, string questionText);
    }
}
