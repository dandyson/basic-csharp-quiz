using OfficeOpenXml;
using System;
using System.IO;

class MainClass
{
    static int score = 0;
    static DateTime startTime = DateTime.Now;
    static string[] questions = { "What is the capital of France?", "Who is the President of the United States?" };
    static string[] answers = { "Paris", "Joe Biden" };
    static string[] userAnswers = new string[questions.Length];

    public static void Main(string[] args)
    {
        DoQuestions(questions, answers);
        Console.WriteLine("Final Score: " + score);
        double secs = (DateTime.Now - startTime).TotalSeconds;
        Console.WriteLine("Time Taken: " + secs);
        StoreQuizResultsInExcel("QuizResults.xlsx", questions, userAnswers, answers, score, secs);
    }

    public static string[] DoQuestions(string[] q, string[] a)
    {
        for (int i = 0; i < q.Length; i++)
        {
            Console.WriteLine(q[i]);
            string userAnswer = Console.ReadLine().Trim().ToLower();
            userAnswers[i] = userAnswer; 

            if (userAnswer == a[i].Trim().ToLower())
            {
                Console.WriteLine("Correct");
                score += 10;
            }
            else
            {
                Console.WriteLine("Wrong. The answer is " + a[i]);
                score -= 5;
            }
        }

        return userAnswers; 
    }

    public static void StoreQuizResultsInExcel(string filePath, string[] questions, string[] userAnswers, string[] correctAnswers, int score, double timeTaken)
    {
        FileInfo file = new FileInfo(filePath);

        if (file.Exists)
        {
            file.Delete();
        }

        using (ExcelPackage package = new ExcelPackage(file))
        {
            // Create a new worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Quiz Results");

            // Define the column headers
            worksheet.Cells[1, 1].Value = "Question";
            worksheet.Cells[1, 2].Value = "User's Answer";
            worksheet.Cells[1, 3].Value = "Correct Answer";

            // Populate the worksheet with data
            for (int i = 0; i < questions.Length; i++)
            {
                worksheet.Cells[i + 2, 1].Value = questions[i];
                worksheet.Cells[i + 2, 2].Value = userAnswers[i];
                worksheet.Cells[i + 2, 3].Value = correctAnswers[i];
            }

            // Add time taken and final score at the bottom
            worksheet.Cells[questions.Length + 3, 1].Value = "Time Taken";
            worksheet.Cells[questions.Length + 3, 2].Value = timeTaken;
            worksheet.Cells[questions.Length + 4, 1].Value = "Final Score";
            worksheet.Cells[questions.Length + 4, 2].Value = score;

            // Save the Excel file
            package.Save();
        }
    }
}