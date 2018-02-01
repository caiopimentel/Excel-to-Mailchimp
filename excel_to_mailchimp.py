import sys
import os
import csv
import collections
import pyperclip

ARG_NAMES = ['script', 'inFile', 'outputType']
ArgList = collections.namedtuple('ArgList', ARG_NAMES)
args = dict(zip(ARG_NAMES, sys.argv))
args = ArgList(*(args.get(arg, 0) for arg in ARG_NAMES))

if args.inFile != 0:
    IN_FILE = open(args.inFile, 'r')
    READER = csv.reader(IN_FILE)
    outData = ''

    for row in READER:
        words = row[0].split(" ")
        row[0] = ''
        newRow = ''

        for i, w in enumerate(words):
            if w.lower() == 'de' or w.lower() == 'do' or w.lower() == 'da':
                w = w.lower()
            else:
                w = w.capitalize()

            if i == 0:
                row[0] += w + ','
            elif i != len(words) - 1:
                row[0] += w + ' '
            else:
                row[0] += w

        for i, r in enumerate(row):
            if i != len(row) - 1:
                newRow += row[i] + ','
            else:
                newRow += row[i]

        outData += newRow + '\r\n'

    IN_FILE.close()

    if args.outputType == '-o':
        OUTFILE = open(os.getcwd() + args.inFile.split('.')[1] + '_OUT.txt', "a+") 
        OUTFILE.write(outData)
        OUTFILE.close()
        print('\nData ready to MailChimp saved on file: ' + os.getcwd() + args.inFile.split('.')[1] + '_OUT.txt\n')
    else:
        pyperclip.copy(outData)
        print('\nData ready to MailChimp copied to clipboard.\n')

else:
    print('\nUsage:\n  excel_to_mailchimp.py inputFile.csv [options]\n')
    print('\nOptions:\n  -o\t\tOutputs a txt file containing the formated data for Mailchimp.\n')
    print('    \t\tWhen no option is used the formated data will be copied to clipboard.\n')
