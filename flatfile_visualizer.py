#!usr/bin/env python


#Matplotlib is utilized in this script. Here is the provided citing:
#@Article{Hunter:2007,
#  Author    = {Hunter, J. D.},
#  Title     = {Matplotlib: A 2D graphics environment},
#  Journal   = {Computing In Science \& Engineering},
#  Volume    = {9},
#  Number    = {3},
#  Pages     = {90--95},
#  abstract  = {Matplotlib is a 2D graphics package used for Python
#  for application development, interactive scripting, and
#  publication-quality image generation across user
#  interfaces and operating systems.},
#  publisher = {IEEE COMPUTER SOC},
#  year      = 2007
#}

import re
import xlrd
import csv
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


labels= []
x=[]
y=[]
plot_titles = []

def stepplot(x, y, labels, plot_titles):

        """Generates Correlation Graph. 
        With the x,y coordinates, labels, and plot titles
        established, the step-plots can be generated. Output 
        is PDF file format"""

        plt.figure()      #makes new image for each plot  
        
        #plot x & y stemplot. format plot points
        plt.stem(x, y, linefmt='k--', markerfmt='ro', basefmt='k-')               
        
        #set x-axis labels and set them vertical. size 10 font.
        plt.xticks(x, labels, rotation='vertical', fontsize = 10)
              
        #set titles for graph and axes
        plt.title(plot_titles[0])
        plt.xlabel("Biomarkers")
        plt.ylabel("Correlation Values")
        
        # slightly move axis away from plot. prevents clipping the labels 
        plt.margins(0.2)
        
        # Tweak spacing to prevent clipping of tick-labels
        plt.subplots_adjust(bottom=0.15)
        
        plt.tight_layout()   #prevents labels from being clipped
                  

        with PdfPages(plot_titles[0]+'.pdf') as pdf: #creates new file for each figure
            pdf.savefig()


def heatmap(data):

    """This function will generate a heatmap of the user's metagenomic 
    organisms vs clinical biomarker correlations. The 
    output is PDF file format. I have adopted the code provided in a 
    stackoverflow.com response by J. Sundram and joelotz to produce the 
    heatmap. URL: http://stackoverflow.com/a/16124677/2900840"""
       
    
    # open plot area. plot data. set blue scheme and legend
    fig, ax = plt.subplots()
    heatmap = ax.pcolor(data, cmap=plt.cm.seismic, alpha=0.8) #set color scheme
    plt.colorbar(heatmap) #create legend for the plot
   
    #set the graph format
    fig = plt.gcf()
    fig.set_size_inches(8,11)

    # turn off the frame
    ax.set_frame_on(False)

    # put the major ticks at the middle of each cell
    ax.set_yticks(np.arange(data.shape[0])+0.5, minor=False)
    ax.set_xticks(np.arange(data.shape[1])+0.5, minor=False)

    ax.invert_yaxis() #flip y-axis from top to bottom
    ax.xaxis.tick_top()  #place x-axis on top of graph

    # Set the labels and format them
    labels = data.columns # make column headings the x-axis labels 
    ax.xaxis.set_label_position('top') #x-axis labels above graph
    ax.set_xticklabels(labels, minor=False, fontsize=8) #no x-axis minor ticks
    ax.set_yticklabels(data.index, minor=False, fontsize=8) #no y-axis minor ticks
    
    #set labels for graph, x-axis, and y-axis


    # rotate the x-axis labels. prevents overlap
    plt.xticks(rotation=90)

    ax.grid(False)

    # Turn off all the axis ticks
    ax = plt.gca()
    for t in ax.xaxis.get_major_ticks(): 
        t.tick1On = False 
        t.tick2On = False 
    for t in ax.yaxis.get_major_ticks(): 
        t.tick1On = False 
        t.tick2On = False
        
    #allow autoformat to prevent overlapping
    ax.axis('auto')
    plt.tight_layout()
    
    #creates new file for each figure
    with PdfPages('heatmap_graph.pdf') as pdf: 
        pdf.savefig()



#Prompt user to enter file
print "Enter the file path (must be: .xls  .xlsx  .csv): "
filepath = raw_input()

#determine the user's file type
if re.search(".xls", filepath) or re.search(".xlsx", filepath):
    print "EXCEL format detected."  #if the file is an excel format
    file_type ="excel"
    
    #prompt user to choose what worksheet he/she wants to create visualization from
    print "What worksheet do you want to process? (ENTER '0' for Sheet 1, '1' for Sheet 2, etc...)"
    sheet_num = int(raw_input()) #make input an integer & set as variable
    
elif re.search(".csv", filepath):
    print "CSV format detected."  #if the file is a csv format
    file_type = "csv"
else:
    # prompts user incorrect file type
    print "ERROR:  Check your file format. The file is not in the correct format."



#Prompt user to chose whether heatmap or step-plots are desired
print """
You can either generate a heatmap of the correlations of
all organisms vs biomarkers OR generate step-plots for each
organism vs biomarkers? Enter 'h' for HEATMAP or 's' for STEP-PLOT. 
"""
userplot = raw_input()

#determine user's choice of plot type
if userplot == 'h' or userplot == 'H':
    print "Heatmap was choosen."
    plot_type ="heatmap"
elif userplot == 's' or userplot == 'S':
    print "Step-plot was choosen."
    plot_type = "stepplot"
else:
    print "ERROR! An 'h' or 's' was not submitted. Please try again."




###EXCEL files are processed here...
if file_type == "excel":    
    ### Convert the .xls or .xlsx file into a csv file
    
    #open user's .xls/.xlsx file
    wb = xlrd.open_workbook(filepath) 
    wb.sheet_names()
    
    #set user's choice of worksheet as activate sheet to be converted into csv
    sheet = wb.sheet_by_index(sheet_num)
    
    # Create new .csv file and populate it with the excel file's data.
    # This block is adopted from a response by M. Pieters and Cacheing
    # on stackoverflow.
    # URL:  http://stackoverflow.com/a/18113843/2900840
    fp = open('result.csv', 'wb')  
    wr = csv.writer(fp, quoting=csv.QUOTE_ALL)
    for rownum in xrange(sheet.nrows):  
        wr.writerow([unicode(val).encode('utf8') for val in sheet.row_values(rownum)])
    fp.close()
    

    ###user selected HEATMAP...
    if plot_type == 'heatmap':
        #read the .csv file just created from the conversion
        data = pd.read_csv('result.csv', index_col=0, header=1, skiprows=1, delimiter=',')
        
        #fills empty cells with a value of '0' (zero)
        data = data.fillna(0)
        
        heatmap(data)  #call function to generate heatmap


    ###user selected STEP-PLOT...
    if plot_type == 'stepplot':
        #read the .csv file just created from the conversion
        with open('result.csv') as csv_file:
            for row in csv.reader(csv_file, delimiter = ','):
                if row[1] == '':  #skips over file titling 
                    continue
                plot_titles.append(row.pop(0)) #takes 1st of element of
                                           #each row & puts it in titles list
            
                row_len = len(row) #counts number items current row        

                if re.search("[A-Z|a-z]", row[0]):  #row containing column names
                    labels = row                    # put into list; x-axis ticks

                else:     
                    y = row    #places remaining elements in y-axis list

                    plot_titles.pop(0)  #1st element of current row stored as plot title 
                    
                    x = [i for i in range(1,(row_len+1))] #creates x-axis length of row
                        
                    # convert empty data cells to 0's (zero)
                    # credit goes to user "umop episdn" for the following line
                    # URL: http://stackoverflow.com/a/15328851/2900840
                    # for each element in list y, y equals element or 0.00
                    y = [element or '0.00' for element in y]     
                                    
                    #call function to generate step-plots
                    stepplot(x, y, labels, plot_titles) 
                    



###CSV file process here...
if file_type == "csv":  
    

    ##user selected HEATMAP...
    if plot_type == 'heatmap':
        #read the file .csv file provided as input 
        data = pd.read_csv(filepath, index_col=0, header=1, skiprows=1, delimiter=',')
        
        #fills empty cells with a value of '0' (zero)
        data = data.fillna(0)

        heatmap(data)  #call function to generate heatmap
        

    ##user selected STEP-PLOT...
    if plot_type == 'stepplot':
        with open(filepath) as csv_file:
            for row in csv.reader(csv_file, delimiter = ','):
               
                if row[1] == '':  #skips over file titling 
                    continue
                plot_titles.append(row.pop(0)) #takes 1st of element of
                                           #each row & puts it in titles list
            
                row_len = len(row) #counts number items current row        

                if re.search("[A-Z|a-z]", row[0]):  #row containing column names
                    labels = row                    # put into list; x-axis ticks

                else:     
                    y = row    #places remaining elements in y-axis list

                    plot_titles.pop(0)  #1st element of current row stored as plot title 
         
                    x = [i for i in range(1,(row_len+1))] #creates x-axis length of row
 
                    # convert empty data cells to 0's (zero)
                    # credit goes to user "umop episdn" for the following line
                    # URL: http://stackoverflow.com/a/15328851/2900840
                    # for each element in list y, y equals element or 0.00
                    y = [element or '0.00' for element in y]  
       
                    #call function to generate step-plots
                    stepplot(x, y, labels, plot_titles)  
