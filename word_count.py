# resume counting
# take some text and count the right number of words
# Include some way to exclude common words that are not relevant
import os
import docx2txt
import pandas as pd


def wordlistToFreq(wordlist):
    resume_count = [wordlist.count(w) for w in wordlist]
    return dict(list(zip(wordlist, resume_count)))


def sortFreq(freqdict):
    sorteddict = [(freqdict[key], key) for key in freqdict]
    sorteddict.sort()
    sorteddict.reverse()
    return sorteddict


def removeCommonWords(cleanedwordlist, stopwords):
    return [w for w in cleanedwordlist if w not in stopwords]


stopwordlist = ['a', 'the', 'to', 'was', 'of', 'in', 'be', 'and', 'an']
stopwordlist += ['that', 'this', 'it', 'all', 'as', 'at', 'by', 'can']
stopwordlist += ['do', 'did', 'from', 'had', 'has', 'hasnt', 'he', 'it', 'its']
stopwordlist += ['or', 'than', 'then', 'were', 'was', 'you', 'your', 'i', 'is']
stopwordlist += ['will', 'with', 'into', 'if', 'on', 'we', 'our', 'for']
stopwordlist += ['have', 'these', 'my']

# run the functions in order
# read files
file_location = input("Please input where the file location is: ")
os.chdir(file_location)

descrip_file = input("Please tell me the description file name: ")
descrip_txt = docx2txt.process(descrip_file)
descrip_txt_lower = descrip_txt.lower()
clean_descrip_txt = descrip_txt_lower.split()
clean_descrip_txt = filter(str.isalpha, clean_descrip_txt)

resume_file = input("Please tell me the resume file: ")
resume_txt = docx2txt.process(resume_file)
resume_txt_lower = resume_txt.lower()
clean_resume_txt = resume_txt_lower.split()
clean_resume_txt = filter(str.isalpha, clean_resume_txt)

output_name = input("Output File name with xlsx: ")


# run functions for both description and resume file
resume_rmv_stop = removeCommonWords(clean_resume_txt, stopwordlist)
resume_freq = wordlistToFreq(resume_rmv_stop)
resume_freq_sort = sortFreq(resume_freq)

descrip_rmv_stop = removeCommonWords(clean_descrip_txt, stopwordlist)
descrip_freq = wordlistToFreq(descrip_rmv_stop)
descrip_freq_sort = sortFreq(descrip_freq)


# move the to Pandas DataFrame to export to Excel
resume_df = pd.DataFrame(resume_freq_sort)
descrip_df = pd.DataFrame(descrip_freq_sort)
combined_df = pd.concat([resume_df, descrip_df], axis=1)

# clean the dataframe for export
combined_df = combined_df.set_axis(['ResWord#', 'Word', 'DescWord#', 'Descword'
], axis=1, inplace=False)

combined_df.to_excel(output_name, sheet_name='Score', index=False)
