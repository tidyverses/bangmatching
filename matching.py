import numpy as np
from operator import itemgetter
import pandas as pd
import random
import sys

# numbers not assigned to fics
NULL_FICS = [1, 14, 17, 18, 22, 23, 25, 28, 31, 34, 37, 43, 44]
# names of fields on input Excel spreadsheets
FIELDS_OF_INTEREST = [
                      "Preferred Name",    # artist name
                      "First Choice Fic",  # fic ranked 1
                      "Second Choice Fic", # fic ranked 2
                      "Third Choice Fic",  # fic ranked 3
                      "Fourth Choice Fic", # fic ranked 4
                      "Fifth Choice Fic"]  # fic ranked 5

def import_input_file(filename, headings, numfics, matches):
    finalmatches = matches
    people = {}
    fics = {}
    df = pd.read_excel(filename)
    names = df.loc[:,headings[0]].tolist()
    c1 = df.loc[:,headings[1]].tolist()
    c1 = [int(i.split('-')[0]) for i in c1] 
    c2 = df.loc[:,headings[2]].tolist()
    c2 = [int(i.split('-')[0]) for i in c2] 
    c3 = df.loc[:,headings[3]].tolist()
    c3 = [int(i.split('-')[0]) for i in c3] 
    c4 = df.loc[:,headings[4]].tolist()
    c4 = [int(i.split('-')[0]) for i in c4] 
    c5 = df.loc[:,headings[5]].tolist()
    c5 = [int(i.split('-')[0]) for i in c5] 

    for x in range(len(names)):
        people[names[x]] = [c1[x], c2[x], c3[x], c4[x], c5[x]]
    
    for x in range(numfics):
        finalmatches[x + 1] = []

    fics = init_fics(numfics, people)

    for person in people.keys():
        for fic in people[person]:
            fics[fic].append([person, people[person].index(fic) + 1])
    
    return people, fics, finalmatches


def init_fics(numfics, people):
    fics = {}

    for x in range(numfics):
        fics[x + 1] = []
    for person in people.keys():
        for fic in people[person]:
            fics[fic].append([person, people[person].index(fic) + 1])
    return fics


def round(rank, ficarray, matches, people):
    finalmatches = matches
    for fic in ficarray.keys():
        random.shuffle(ficarray[fic])
        for person in ficarray[fic]:
            if person[1] == rank and finalmatches[fic] == [] and person[0] not in finalmatches.values():
                finalmatches[fic] = person[0]
        newlist = [x for x in ficarray[fic] if not (x[1] == rank or x[0] in finalmatches.values())]
        ficarray[fic] = newlist
        newnames = [x for x in people.keys() if not x in finalmatches.values()]
        if len(newnames) == 0:
            break

    return finalmatches


def main_algorithm(people, fics, matches, extra):
    finalmatches = matches
    newnames = [x for x in people.keys() if not x in finalmatches.values()]

    if extra == False:
        # before matching, if any fic has only 1 person who ranked, give them that fic
        for fic in fics.keys():
            if len(fics[fic]) == 1 and fics[fic][0][0] not in finalmatches.values():
                finalmatches[fic] = fics[fic][0][0]
            newnames = [x for x in people.keys() if not x in finalmatches.values()]
            if len(newnames) == 0:
                return finalmatches
    # go through each fics - if one or more people ranked the fic #1, 
    # choose one of the #1 rankers at random and assign them the fic.
    # then go through the unassigned fics - if one or more people ranked the fic #2,
    # choose one of the #2 rankers at random and assign them the fic.
    # repeat for ranks #3, #4, and #5 with all fics still unassigned
    originalfics = fics.copy()
    ranksleft = 5

    while ranksleft > 0:
        for x in range(ranksleft):
            print(x + 6-ranksleft)
            finalmatches = round(x+(ranksleft-4), fics, finalmatches, people)
            newnames = [x for x in people.keys() if not x in finalmatches.values()]
            if len(newnames) == 0:
                break
        fics = originalfics.copy()
        ranksleft -= 1
        if len(newnames) == 0:
                break
        
    if not len(newnames) == 0:
        finalmatches = round(5, fics, finalmatches, people)

    # try to give unassigned fics to people without fics who ranked the unassigned fic in their top 5
    newnames = [x for x in people.keys() if not x in finalmatches.values()]
    for fic in finalmatches.keys():
        if finalmatches[fic] == []:
            for person in originalfics[fic]:
                persons_old_fic = list(finalmatches.keys())[list(finalmatches.values()).index(person[0])]
                unused_list_takers = list(set(newnames) & set(list(map(itemgetter(0), originalfics[persons_old_fic]))))
                if len(unused_list_takers) > 0:
                    finalmatches[fic] = person[0]
                    finalmatches[persons_old_fic] = unused_list_takers[0]
                    break
                newnames = [x for x in people.keys() if not x in finalmatches.values()]

    return finalmatches


def export_matches(people, matches, extramatches):
    finalmatches = matches
    dfoutput = {'Fic': list(finalmatches.keys()), 
                'Artist 1': list(finalmatches.values()), 
                'Artist 1 Ranking':[0]*len(finalmatches.keys()),
                'Artist 2': list(extramatches.values()), 
                'Artist 2 Ranking':[0]*len(finalmatches.keys()),}
    for fic in finalmatches.keys():
        if fic in NULL_FICS:
            dfoutput['Artist 1'][fic - 1] = '---'
            dfoutput['Artist 1 Ranking'][fic - 1] = "---"
            dfoutput['Artist 2'][fic - 1] = '---'
            dfoutput['Artist 2 Ranking'][fic - 1] = "---"
        else:
            name = finalmatches[fic]
            if name != []:
                dfoutput['Artist 1 Ranking'][fic - 1] = people[name].index(fic) + 1
            else:
                dfoutput['Artist 1 Ranking'][fic - 1] = ''
                dfoutput['Artist 1'][fic - 1] = ''
            name2 = extramatches[fic]
            if name2 != []:
                dfoutput['Artist 2 Ranking'][fic - 1] = people[name2].index(fic) + 1
            else:
                dfoutput['Artist 2 Ranking'][fic - 1] = ''
                dfoutput['Artist 2'][fic - 1] = ''
    df2 = pd.DataFrame(data=dfoutput)
    df2.to_excel("matches.xlsx") 


if __name__ == '__main__':
    originalmatches = {}
    originalpeople, originalfics, originalmatches = import_input_file(sys.argv[1], FIELDS_OF_INTEREST, 
                                                    int(sys.argv[2]), originalmatches)
    
    for fic in originalfics:
        print(str(fic) + " " + str(originalfics[fic][:int(len(originalfics[fic]) / 2)]))
        originalfics[fic] = originalfics[fic][:int(len(originalfics[fic]) / 2)]
    
    fics = originalfics.copy()
    extramatches = originalmatches.copy()
    finalmatches = originalmatches.copy()
    # generates first-artist matches until it comes up with a result where all fics have 1 artist
    while list(finalmatches.values()).count([]) > len(NULL_FICS):
        fics = originalfics.copy()
        people = originalpeople.copy()   
        extramatches = originalmatches.copy()
        finalmatches = originalmatches.copy()
        finalmatches = main_algorithm(people, fics, finalmatches, False)
    # removes matched artists from list of artists left to match    
    newpeople = originalpeople.copy()
    for k in list(finalmatches.values()):
        if k != []:
            newpeople.pop(k, None)
    # runs matching again with remaining artists to generate second-artist matches
    extrafics = init_fics(int(sys.argv[2]), newpeople)
    extramatches = main_algorithm(newpeople, extrafics, extramatches, True)
    # put matches in excel spreadsheet
    export_matches(originalpeople, finalmatches, extramatches)
    
    
    
        
