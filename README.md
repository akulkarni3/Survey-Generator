Survey-Generator
================

This program allows to generate surveys for 100 people consisting of 10 question each. 
The average ratings can be pre-decided and the algorithm generates random numbers such that the average remains the close to what is expected.
The survey is filled with 1 star rating to 5 star rating.
The code is optimized for giving maximum number of different values from 1 to 5.

    Features :
1. Generates surveys randomly.
2. The average rating of survey is pre-determined and maintained after execution.
3. The average ratings of each surveyee is determined.
4. The average ratings given to each question is also calculated.
5. The result is outputted on spreadsheet which is compatible with MS-Excel, Libreoffice, OpenOffice etc

    Limitation:
1. Survey is only generated for 100 people at a time and each person is given 10 question. This condition is mandatory as it is required to generate user given average keeping the numbers random.
2. File containing 100 random names is required. Sample names.txt is uploaded.

    Installation:
//Debian Linux
1. sudo apt-get install python pip
2. sudo pip install XlsxWriter

    Execution:
1. chmod +x survey-generator.py
2. ./survey-generator.py names.txt {OR python survey-generator.py names.txt}
3. Follow the on-screen instructions

    Output:
SurveyGenerator.xlsx is generated. Sample file is uploaded.
