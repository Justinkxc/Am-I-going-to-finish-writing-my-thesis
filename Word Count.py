
'''
import os

cwd = os.getcwd()
files = os.listdir(cwd)
print("Files in %r: %s" % (cwd, files))
'''

import docx2txt

# dictionary to store the word counts
word_counts = {}

# define a function to check for proper words (letters only)
import string
def processWord(word):
    # remove starting punctuation
    #if word.startswith("\"") or word.startswith("\'"):
       # word = word[1:]
    # remove trailing punctuation
    if word.endswith(".") or word.endswith("?") or word.endswith("!") or word.endswith(","):
        word = word[:-1]
    # check for non letters
    for l in list(word):
        if l not in string.ascii_lowercase:
            # ascii_lowercase brings out all the lowercase letters of the alphabet
            return ""
    return word

# read all the text and count the words
infile = docx2txt.process("Ms Thesis Draft 1.docx")
lines = infile.split("\n")
for line in lines:
    words = line.rstrip().lower().split()
    for word in words:
        word = processWord(word)
        #print(word)
        if word != "":
            if word not in word_counts:
                word_counts[word] = 0

            word_counts[word] = word_counts[word] + 1

# get the total word count
total = sum(word_counts.values())
print("Total word count :", total)
print()

# get a count list for sorting
word_list = []
for word in word_counts:
    word_list.append((word_counts[word] / total, word))

# sort the word list
word_list.sort(reverse=True)

# print top 20
print("Top 20 words")
for (freq, word) in word_list[:20]:
    percent = freq * 100
    print("{0:<12s}  {1:0.3f}".format(word, percent))

print()  # blank line

# bottom 20
print("Bottom 20 words")
for (freq, word) in word_list[-20:]:
    percent = freq * 100
    print("{0:<12s}  {1:0.3f}".format(word, percent))
