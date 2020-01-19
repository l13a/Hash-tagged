import xlrd
import xlwt
from xlwt import Workbook
import csv
from typing import List, Dict, TextIO

loc = ("MentalHealth.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# Stores ALL words in list
list_words = []
for i in range(sheet.nrows):
    list_words.extend(sheet.cell_value(i, 3).split(' '))

i = 2
while i < len(list_words):
    if list_words[i].startswith('@') or list_words[i].startswith('#'):
        list_words[i] = ''
    else:
        if not list_words[i].isalpha():
            temp = ''
            for char in list_words[i]:
                if char.isalpha():
                    temp += char
            list_words[i] = temp.lower()
    i += 1

list_words.sort()

# Step 0: List of determiner words
file = open('Determiners.txt','r')
determiners = file.readlines()
for i in range(len(determiners)):
    determiners[i] = determiners[i].strip()

# Step 1: Create Dict of word and frequency
word_to_frequency = {}
j = 0
while j < len(list_words):
    word = list_words[j].lower()
    # Making sure determiners are not counted
    if word not in determiners:
        if word not in word_to_frequency:
            word_to_frequency[word] = 1
        elif list_words[j] in word_to_frequency:
            word_to_frequency[word] += 1
    j += 1

#Step 2: Sort words by frequency
frequencies = []
frequency_to_word = {}

for key in word_to_frequency:
    value = word_to_frequency[key]
    # Add frequency in frequencies list
    frequencies.append(value)
    # Make reverse dictionary
    if value in frequency_to_word:
        frequency_to_word[value].append(key)
    else:
        frequency_to_word[value] = [key]

frequencies.sort(reverse=True)
unique_frequencies = []

for value in frequencies:
    if value not in unique_frequencies:
        unique_frequencies.append(value)

# Write top 50 results onto excel

# Workbook is created

# wb = Workbook()
# sheet1 = wb.add_sheet('Sheet 1')
#
# counter = 0
#
# for value in unique_frequencies:
#     for word in frequency_to_word[value]:
#         if counter < 50:
#             sheet1.write(counter, 0, word)
#             sheet1.write(counter, 1,value)
#             counter += 1
# wb.save('Results.csv')

with open('Results.csv', mode='w') as results_file:
    results_writer = csv.writer(results_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    counter = 0

    results_writer.writerow(['Word', 'Frequency'])
    for value in unique_frequencies:
        for word in frequency_to_word[value]:
            if counter < 30:
                results_writer.writerow([word, value])
                counter += 1
