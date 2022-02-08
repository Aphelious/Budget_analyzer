##Budget Utility

---

###Objective:

This application is designed to show clear visualizations of personal 
finance data, accepting .xlsx or .csv files as input. First, input data is 
parsed, either using standard Python functionality or NLP techniques,
to convert the often obtuse transaction descriptions given by banking
software to something that makes ore sense to the user. The model will
learn to categorize the transaction descriptions into more useful and 
recognizable entries and this model will persist for the user to benefit
from in the future. A series of operations is then performed to analyze 
the transactions and generate the necessary descriptive statistics of 
the banking transactions. The application then calls on a suite of 
sophisticated visualization tools to produce clear and interesting graphs
that explain the user's personal financial behavior in ways that empower
them to make changes if they choose to do so. 

###Implementation:

At first the xlsx parser was pure Python, chaining several function calls 
together to break down each description line in the spreadsheet into
individual words or 'word-ish' sequences of letters and numbers. It then 
removes extraneous data, and combine it all in a list of lists; one list 
corresponding to one row in the spreadsheet. The script then presents the 
user with each 'word' and asks how to categorize it. The problem I ran 
into with this method was that I don't have the expertise to instruct 
Python to find words that occur together often. Thus, the parser sees 
the line 'John's Auto Mechanic,' which should be easy to categorize 
as 'Car repairs,' as separate entries: 'John's,' 'Auto,' and 'Mechanic.' 
Here the user must settle for throwing away 'John's' and 'Auto' as too 
general, relying on 'Mechanic' to be the only entry point into the 
'Car repair' category. This approach bothered me; what's the use of 
throwing away good data?

I opted for incorporating some basic NLP techniques to help tease out 
description terms in the spreadsheet that would be more informative if
kept together. Collocation is the best fit technique and I incorporated 
the library in a branch called 'NLP'. What I really like about this project 
is that the parser is basically a very rudimentary machine learning 
algorithm. 


####Update:
So after trying to wrangle Excel into formatting the dates the way I want
I'm not really satisfied. For one, openpyxl seems to create datetime objects 
whenever it finds a date, which is good to have that functionality. The
problem seems to be that when writing that data to the spreadsheet, Excel
parses the dates and converts them to their own date data type and then 
formats them differently, making it hard to control the final output of 
dates. So, yet again, I'm bumping into the need for a database and an
ORM to manage things. As such, I've decided to switch the implemetation to
a SQLite database, using SQLachemy's ORM. 