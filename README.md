# personality-quiz-outcome-generator
This code looks into an excel file, looks at all the questions with the given points, and gives you an output of the chance that each category has of being the winner.

This was really useful for myself when I would creat quizzes for my wife's blog.
You can check out one such example here: https://millenniologyblog.com/wp_quiz/which-pregnancy-pillow-is-perfect-for-you/

I provided an example excel file where in the first sheet you enter all your possible results, and in the remaining sheets you enter all your questions.
You can add as many categories and as many questions as you like.
For each question, you enter all the answers and allocate points to each category as you like.

Once you have filled in the excel file, run the code.
It will first prompt you to enter the name of the excel sheet. Make sure the excel file is in the same directory as your python code.
Once you've entered the name, it will automatically calculate all the possible ways the quiz can be completed.
It will then give you an output of the absolute number of ways, and a percentage probability that each category has of being the winner.
