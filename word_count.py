# This script uses libraries (see the import statements below)
#   If you have trouble, you might need to run the following commands in
#   a command window to install the packages
#       pip install python-docx
#       pip install pandas
#       pip install tk


from docx import Document
import re
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import ast
import argparse

class Socialization_of_Emotion(object):
    """
    A class used to handle the data collection from .docx files as well as 
    key word searches
    
    Examples
    --------
    >>> #Default Parameters
    >>> soe = Socialization_of_Emotion()
    >>> soe.search_keys()    
    >>> #Rerun without rebuilding
    >>> soe = Socialization_of_Emotion(setUpFile="results/word_count_final.csv")
    >>> soe.search_keys()
    
    >>> #Use different .docs folder
    >>> soe = Socialization_of_Emotion(originDirectory="2020 Combined Cleaned Transcriptions",
                                       saveDirectory="results",
                                       saveFile="word_counts_final2.csv",
                                       keysFile="NEW keywords2.csv")
    >>> soe.search_keys()
    >>> #Rerun without rebuilding
    >>> soe = Socialization_of_Emotion(setUpFile="results/word_count_final2.csv")
    >>> soe.search_keys()
    """
    def __init__(self, 
                 setUpFile=None,
                 originDirectory="2019 Combined Cleaned Transcriptions",
                 saveDirectory="results",
                 saveFile="word_counts_final.csv",
                 keysFile="NEW keywords.csv",     
                 verbose=True
                ):
        """
        Parameters
        ----------
            setupFile : str or None, optional
                Path to the a set up file (saved from earlier). If None, then will
                use the files in originDirectory to create a set up file, saveFile,
                in saveDirectory. Default is None
                
            originDirectory : str, optional
                Directory storing the .docx files; only matters when setUpFile is None.
                Default is '2019 Combined Cleaned Transactions'
                
            saveDirectory : str, optional
                Directory to save saveFile to; only matters when setUpFile is None.
                Default is 'results'
                
            saveFile : str, optional
                CSV file to save the results to; default is 'word_counts_final.csv'
                
            keysFileName : str, optional
                CSV file with search queries; default is 'NEW keywords.csv'     
                
            verbose : bool, optional
                Whether to print out helpful statements to show progress; default is True
        """
        #Check input
        if type(saveFile) != str or saveFile[-4:] != '.csv':
            raise ValueError("saveFile must be a .csv file; got {}".format(saveFile))
        if type(keysFile) != str or keysFile[-4:] != '.csv':
            raise ValueError("keysFileName must be a .csv file; got {}".format(keysFile))
            
        self.originDirectory = originDirectory
        self.saveDirectory = saveDirectory
        self.saveFile = saveFile
        self.keysFile = keysFile
        self.verbose=verbose
        if setUpFile is not None:
            try:
                self.df_counts = pd.read_csv(setUpFile)
                self.setUpFile = setUpFile
            except:
                raise ValueError("Could not find setUpFile at {}".format(setUpFile))
        else:
            self._setUp()            
        
    def _setUp(self):
        directory = os.fsencode(self.originDirectory)
        columnNames = ['ID', 'PageNum', 'P_transcript', 'C_transcript', 'P_WordCount', 'C_WordCount']
        self.df_counts = pd.DataFrame(columns=columnNames)
        for file in os.listdir(directory):  # Grab each file, one at a time
            filename = os.fsdecode(file)  # Take the code from computer language to a path
            if filename.split('.')[-1] != 'docx':
                continue #skip non .docx files
            openName = self.originDirectory + os.sep + filename # I want to leave the original files alone, so put edited files in a different location
            self._updateWordCount(openName)
            self.setUpFile = self.saveDirectory + os.sep + self.saveFile
            self.df_counts.to_csv(self.setUpFile, index=True, index_label='Participant_ID')
        
    def _updateWordCount(self, filename):
        """
        Iterates through a .docx file and adds data for each page number
        
        Parameters
        ----------
            filename : str
                Path to a .docx file
        """
        currDoc = filename.split(os.sep)[-1][:-5]
        if self.verbose:
            print('Working to analyze ' + currDoc)  # User-friendly message
        currIndex = None
        currPage = None
        currSpeaker = None
        doc = Document(filename)  # Open the current document
        for paragraph in doc.paragraphs:
            if paragraph.text == '':
                continue

            # Update from a speaker change
            if paragraph.text[:5].lower() in ['paren', 'child']:
                currSpeaker = paragraph.text[:5].lower()

            # Update from a page change
            page_in_para = re.search('\[([pP]age )(\d+-*\w*)\]', paragraph.text)
            if page_in_para:
                currPage = page_in_para[2]
                currIndex = currDoc + '-Page-' + currPage
                if (currPage.lower == 'stop'):
                    if self.verbose:
                        print('Stop Page Detected')
                else:
                    self.df_counts.loc[currIndex, 'ID'] = currDoc
                    self.df_counts.loc[currIndex, 'PageNum'] = currPage

            if currIndex is not None:
                # Save the transcript
                if currSpeaker == 'paren':
                    if pd.isna(self.df_counts.loc[currIndex, 'P_transcript']):
                        self.df_counts.loc[currIndex, 'P_transcript'] = ""
                    self.df_counts.loc[currIndex, 'P_transcript'] += ' ' + paragraph.text
                elif currSpeaker == 'child':
                    if pd.isna(self.df_counts.loc[currIndex, 'C_transcript']):
                        self.df_counts.loc[currIndex, 'C_transcript'] = ""
                    self.df_counts.loc[currIndex, 'C_transcript'] += ' ' + paragraph.text

                # Handle the initialization of a new cell
                if pd.isna(self.df_counts.loc[currIndex, 'P_WordCount']):
                    self.df_counts.loc[currIndex, 'P_WordCount'] = 0
                if pd.isna(self.df_counts.loc[currIndex, 'C_WordCount']):
                    self.df_counts.loc[currIndex, 'C_WordCount'] = 0

                # Update the Word Count
                wordcount = len(re.findall("(\S+)", paragraph.text))  # Count the total number of words
                pageNum = len(re.findall(r'\[([pP]age )(\d+-*\w*)\]',
                                         paragraph.text))  # Find all the [page ##] to subtract from word count
                wordCorrect = len(re.findall(r'\[\[\w+\]\]',
                                             paragraph.text))  # Find all the instances of 'nake [[snake]] where word is corrected
                name = len(re.findall(r'\[[a-zA-Z\']+ name\]',
                                      paragraph.text))  # Find the instances where name has been de-identified
                words = wordcount - 2*pageNum - wordCorrect - name  # Each [page ##] has two words, each [[corrected]] has one, and count [child's name] as 1 word instead of two
                if (currSpeaker == 'paren'):
                    self.df_counts.loc[currIndex, 'P_WordCount'] += words
                elif (currSpeaker == 'child'):
                    self.df_counts.loc[currIndex, 'C_WordCount'] += words
                    
    def search_keys(self):
        """
        Iterates over keysFile and searches each transcript for a match. 
        Creates a parent column and child column for each query
        """
        if 'P_transcript' not in self.df_counts.columns:
            raise Exception("Invalid setUpFile at {}: missing column P_transcript".format(self.setUpFile))
        if 'C_transcript' not in self.df_counts.columns:
            raise Exception("Invalid setUpFile at {}: missing column C_transcript".format(self.setUpFile))
            
        keywords = pd.read_csv(self.keysFile)
        if self.verbose:
            print('Analyzing Key Words')
        for index, row in keywords.iterrows():
            for query in row:
                if isinstance(query, str):
                    self.df_counts.loc[:, 'P_' + query] = self.df_counts.P_transcript.str.lower().str.count(query)
                    self.df_counts.loc[:, 'C_' + query] = self.df_counts.C_transcript.str.lower().str.count(query)
        self.df_counts.fillna(0, inplace=True)
        savePath = self.saveDirectory + os.sep + self.saveFile
        self.df_counts.to_csv(savePath, index=True, index_label='Participant_ID')
        if self.verbose:
            print("Results saved to {}".format(savePath))
            
def main(args):
    parser = argparse.ArgumentParser(description="Argument parser for word_count.py")
    parser.add_argument('-o', type=ast.literal_eval, default=True, help="Whether to prompt the user for an origin directory (where to search for the .docx files).")
    parser.add_argument('-s', type=ast.literal_eval, default=True, help="Whether to prompt the user for a save directory (where the results will be saved to).")
    parser.add_argument('-k', type=ast.literal_eval, default=True, help="Whether to prompt the user for a keys file (csv with search queries).")
    parser.add_argument('-r', type=ast.literal_eval, default=False, help="Whether to prompt the user for a setUpFile (csv with the previous results)")
    parser.add_argument('--setUpFile', type=str, default='results' + os.sep + 'word_counts_final.csv', help="File with previous results.")
    parser.add_argument('--originDirectory', type=str, default='2019 Combined Cleaned Transcriptions', help="Directory to search for the .docx files.")
    parser.add_argument('--saveDirectory', type=str, default='results', help="Directory the results will be saved to.")
    parser.add_argument('--keysFile', type=str, default="New keywords.csv", help="csv file with search queries.")
    parser.add_argument('--saveFile', type=str, default='word_counts_final.csv', help="csv file to save the results to.")
    parser.add_argument('--verbose', type=ast.literal_eval, default=True, help="Whether to print out progress messages.")       
    opts = parser.parse_args(args)
    
    root = tk.Tk()
    root.withdraw()
    
    if opts.o and not opts.r:
        originDirectory = filedialog.askdirectory(title="Select the folder with your .docx files")
    else:
        originDirectory = opts.originDirectory
    if opts.s and not opts.r:
        saveDirectory = filedialog.askdirectory(title="Select the folder to save results to")
    else:
        saveDirectory = opts.saveDirectory
    if opts.k:        
        keysFile = filedialog.askopenfilename(title="Select a csv file containing your keywords",
                                           filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
    else:
        keysFile = opts.keysFile
    if opts.r:        
        setUpFile = filedialog.askopenfilename(title="Select a csv file containing previous results",
                                           filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*")))
    else:
        setUpFile = None    
    
    soe = Socialization_of_Emotion(setUpFile=setUpFile,
                                   originDirectory=originDirectory,
                                   saveDirectory=saveDirectory,                                   
                                   saveFile=opts.saveFile,
                                   keysFile=keysFile,
                                   verbose=opts.verbose)
    soe.search_keys()
            
if __name__ == "__main__":
    main(sys.argv[1:])
