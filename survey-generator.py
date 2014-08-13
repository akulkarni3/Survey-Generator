#!/usr/bin/python
from __future__ import division
import xlsxwriter
import random
from sys import argv

#Global variables
average                  = 0
lists                    = []
average_person_ratings   = []
average_question_ratings = []

#Random Number Generator
def myfunction(users,maxusers,minvalue,maxvalue):
    global average
    global lists
    while users < maxusers:
        questions = 0
        while questions < 10:
            x= random.randint(minvalue,maxvalue)
            lists.append(x)
	    questions+=1
        users +=1
   
if __name__=="__main__":
    avg       = 5.000
    c         = 0
    sets      = 1
    checklist = []
    sums      = 0
    count     = 0
    print "\t\t\t\t     Survey Generator    \n\tThis algorithm computes for 100 users & assuming 10 ques per user.\n\tThe survey is filled with 1 star rating to 5 star rating.\n\tThe code is optimized for giving maximum number of different values from 1 to 5.\n"
    while avg > .99:
        print "\tcase ",c," : ~ %.3f" % avg
 	checklist.append(avg)
 	c += 1
 	avg = avg - 0.125
    var = int(raw_input("\nType the Case Number :: Eg. Type '5' if you want average to be 4.375  >> "))
    print " The average you want close to is   ", checklist[var],"\n"
    a = [[5,5,5,5], 
         [5,5,5,4.5],
         [5,5,4.5,4.5],
         [5,4.5,4.5,4.5],
         [4.5,4.5,4.5,4.5],
         [4.5,4.5,4.5,4],
         [4.5,4.5,4,4],
         [4.5,4,4,4],
         [4,4,4,4],
         [4,4,4,3.5],
         [4,4,3.5,3.5],
         [4,3.5,3.5,3.5],
         [3.5,3.5,3.5,3.5],
         [3.5,3.5,3.5,3],
         [3.5,3.5,3,3],
         [3.5,3,3,3],
         [3,3,3,3],
         [3,3,3,2.5],
         [3,3,2.5,2.5],
         [3,2.5,2.5,2.5],
         [2.5,2.5,2.5,2.5],
         [2.5,2.5,2.5,2],
         [2.5,2.5,2,2],
         [2.5,2,2,2],
         [2,2,2,2],
         [2,2,2,1.5],
         [2,2,1.5,1.5],
         [2,1.5,1.5,1.5],
         [1.5,1.5,1.5,1.5],
         [1.5,1.5,1.5,1],
         [1.5,1.5,1,1],
         [1.5,1,1,1],
         [1,1,1,1]
]
    while sets < 5:
        users = 1
        maxusers = 26
        for col in a[var]:
           if col == 5:
               minvalue = 5
               maxvalue = 5
           elif col == 4.5:
               minvalue = 4
               maxvalue = 5
           elif col == 4:
               minvalue = 3
               maxvalue = 5
           elif col == 3.5:
               minvalue = 2
               maxvalue = 5
           elif col == 3:
               minvalue = 1
               maxvalue = 5
           elif col == 2.5:
               minvalue = 1
               maxvalue = 4
           elif col == 2:
               minvalue = 1
               maxvalue = 3
           elif col == 1.5:
               minvalue = 1
               maxvalue = 2
           elif col == 1:
               minvalue = 1
               maxvalue = 1

           myfunction(users,maxusers,minvalue,maxvalue)
           users    = users+25
           maxusers = maxusers+25
           sets += 1

    for element in lists:
         sums = sums + element
         count += 1
    average= sums/count

    #open spreadsheet page
    workbook = xlsxwriter.Workbook('SurveyGenerator.xlsx')
    worksheet = workbook.add_worksheet()

    #access Random names file via command line argument
    script , filename = argv
    txt = open(filename)
    print "Names of survey takers are given in file %r. " % filename
    logs = [line.strip() for line in open(filename)]

    #Header
    worksheet.write('H1','Survey Generator')
    worksheet.write('I2','--Created By: Ashish Kulkarni')

    #Decorating Spreadsheet
    worksheet.write('A4','SR No.')
    worksheet.write('B4','Names')
    i=1
    while i < 11:
        letter = chr(ord('D')+i)
        cellno = letter+ str(4)
        quesno = 'Ques. ' + str(i)
        worksheet.write(cellno,quesno)
        i     += 1
    worksheet.write(chr(ord('D')+i+1)+str(4), 'Avg')
    footer_alphabet= chr(ord('D')+i+3)

    i=1
    for log in logs:
        cellno = 'A'+str(i+5)
        worksheet.write(cellno,i)
        cellno = 'B'+str(i+5)
        worksheet.write(cellno,log)
        i+=1
    worksheet.write('B'+ str(i+6),'Avg. for each Ques.')
    footer_number = i+8


    #Assumptions
    worksheet.write(chr(ord(footer_alphabet)-8)+str(footer_number+1), 'Assumptions ::')
    worksheet.write(chr(ord(footer_alphabet)-7)+str(footer_number+2), '5 - Excellent')
    worksheet.write(chr(ord(footer_alphabet)-7)+str(footer_number+3), '4 - Good')
    worksheet.write(chr(ord(footer_alphabet)-7)+str(footer_number+4), '3 - Average')
    worksheet.write(chr(ord(footer_alphabet)-7)+str(footer_number+5), '2 - Below Average')
    worksheet.write(chr(ord(footer_alphabet)-7)+str(footer_number+6), '1 - Poor')

    #Footer
    worksheet.write(chr(ord(footer_alphabet)-4)+str(footer_number+1), 'Symbols ::')
    worksheet.write(chr(ord(footer_alphabet)-3)+str(footer_number+2), '* - Max. Average Ratings')
    worksheet.write(chr(ord(footer_alphabet)-3)+str(footer_number+3), '# - Min. Average Ratings')
    worksheet.write(chr(ord(footer_alphabet)-4)+str(footer_number+5), 'Overall Average Ratings ::')
    worksheet.write(chr(ord(footer_alphabet)-3)+str(footer_number+6), str(average))

    #Data Simulation //dynamic-values
    row_counter = 0
    no = 6
    counter =0
    while row_counter < 100:
        col_counter    = 0
        person_ratings = 0
        while col_counter < 10:
            worksheet.write(chr(ord('E')+col_counter)+str(no+row_counter),lists[counter])
            person_ratings = person_ratings + lists[counter]
            counter += 1
            col_counter += 1
        average_person_ratings.append(person_ratings/10)
        worksheet.write(chr(ord('E')+col_counter+1)+str(no+row_counter),person_ratings/10)
        row_counter += 1

    #Max/Min Average Value
    max_index = average_person_ratings.index(max(average_person_ratings)) #gets index of max value 
    min_index = average_person_ratings.index(min(average_person_ratings)) #gets index of min value
    worksheet.write('Q'+ str(6+max_index),'*')
    worksheet.write('Q'+ str(6+min_index),'#')
    
    #Average Question Ratings   
    i         = 0
    counter   = 0
    while i < 10:
        summation = 0
        while counter < 1000:
            summation = summation + lists[counter]
            counter   = counter + 10
        average_question_ratings.append(summation/100)
        i      += 1
        counter = i
    incrementor   = 0
    for avg_ques in average_question_ratings:
        worksheet.write(chr(ord('E')+incrementor)+ str(107),avg_ques)
        incrementor += 1

    #Best/Worst Rated Question
    max_index = average_question_ratings.index(max(average_question_ratings)) #gets index of max value 
    min_index = average_question_ratings.index(min(average_question_ratings)) #gets index of min value
    worksheet.write(chr(ord('E')+max_index)+ str(108),'*')
    worksheet.write(chr(ord('E')+min_index)+ str(108),'#')
    best_string  = 'Question No. '+ str(max_index+1) +' is the best rated question.' 
    worst_string = 'Question No. '+ str(min_index+1) +' is the worst rated question.'

    #Conclusion
    worksheet.write('C'+str(footer_number+8),'Conclusion ::')
    worksheet.write('D'+str(footer_number+10),str(best_string))
    worksheet.write('D'+str(footer_number+12),str(worst_string))

    #Closing the spreadsheet
    workbook.close()
