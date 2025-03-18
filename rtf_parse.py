import pandas as pd
from striprtf.striprtf import rtf_to_text

from tkinter import filedialog

file_path = filedialog.askopenfilename(title="Select a File", filetypes=[("RTF files", "*.rtf"), ("All files", "*.*")])

with open(file_path) as infile:
    content = infile.read()
    text = rtf_to_text(content).split('.\n')

dict_data = {}

word_mansi = []
part_speech = []
meaning1 = []
meaning2 = []
example1 = []
example2 = []

for i in range(len(text)):
    text_line = text[i].split(';')
    text_line[0] = text_line[0].replace('см.','см,')

    if len(text_line) == 4 and text_line[1].find('–') == -1:

        parts_line1 = ''.join(text_line[0:1])
        parts_line1 = parts_line1.split('. ')
        mword_ps_mean = parts_line1[0].split(' ')
        mean_val2 = text_line[1]

        if mean_val2.find(')'):
            parts_line1_2 = mean_val2.split(') ')
            mean_val2 = parts_line1_2[-1]

        if mean_val2.find('–') != -1:
            meaning2.append(" ")
            ex_val1 = text_line[-2]
            example1.append(ex_val1 + mean_val2)
        else:
            if mean_val2.find('2.') > -1:
                mean_val2 = mean_val2.split('2.')
            elif mean_val2.find('1.') > -1:
                mean_val2 = mean_val2.split('1.')
            else:
                mean_val2 = mean_val2.split('2)')
            meaning2.append(mean_val2[-1].strip())
            ex_val1 = text_line[-2]
            example1.append(text_line[-2].strip())

        ex_val2 = text_line[-1]
        example2.append(text_line[-1])
        mean_val1 = text_line[0].split('.')
        meaning1.append(mean_val1[1].replace('1)','').strip())
    else:
        meaning2.append(' ')
        ex_val1 = text_line[-1]
        example1.append(text_line[-1])
        ex_val2 = ' '
        example2.append(' ')

        parts_line1 = text_line[0].split('. ')

        if len(parts_line1) > 2:
            if parts_line1[1].find('(') != -1:
                parts_line1[-1] = '.'.join(parts_line1[-2:])
            else:
                parts_line1[-1] = parts_line1[1]

        mword_ps_mean = parts_line1[0].split(' ')
        meaning1.append(parts_line1[-1])

    word_mansi.append(' '.join(mword_ps_mean[:-1]))
    part_speech.append(mword_ps_mean[-1])

dict_data['word'] = word_mansi
dict_data['part_speech'] = part_speech
dict_data['meaning1'] = meaning1
dict_data['meaning2'] = meaning2
dict_data['example1'] = example1
dict_data['example2'] = example2

data_to_file = pd.DataFrame(dict_data)

data_to_file.to_excel('out.xlsx')
data_to_file.to_csv('out.csv')