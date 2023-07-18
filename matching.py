import numpy as np
from operator import itemgetter
import pandas as pd
import random
import sys

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


def samplenames(numfics, matches):
    finalmatches = matches
    people = {}
    names = ['Alice', 'Arthur', 'Brenda', 'Bob', 'Christine', 'Charlie', 'Dana', 'Donald', 'Ella', 'Ethan',
          'Farrah', 'Frank', 'Greta', 'George', 'Helen', 'Harry', 'Irma', 'Isaac', 'Joan', 'Jacob',
          'Kathy', 'Kyle', 'Lauren', 'Louis', 'Megan', 'Miles', 'Nina', 'Nathan', 'Olivia', 'Oliver',
          'Patty', 'Peter', 'Quinn', 'Quentin', 'Ruth', 'Ryan', 'Sarah', 'Sam', 'Tina', 'Tyler', 'Ursula',
          'Ulysses', 'Virginia', 'Victor', 'Winona', 'Wyatt', 'Ximena', 'Xavier', 'Yvonne', 'Yosef',
          'Zinnia', 'Zander']
    names = names[:45]
    for x in range(len(names)):
        people[names[x]] = np.random.choice(np.arange(1, numfics+1), replace=False, size=(5)).tolist()

    for x in range(numfics):
        fics[x + 1] = []
        finalmatches[x + 1] = []

    for person in people.keys():
        for fic in people[person]:
            fics[fic].append([person, people[person].index(fic) + 1])

    return people, fics, finalmatches

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

def main_algorithm(people, fics, matches):
    finalmatches = matches
    newnames = [x for x in people.keys() if not x in finalmatches.values()]

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
        name = finalmatches[fic]
        if name != []:
            dfoutput['Artist 1 Ranking'][fic - 1] = people[name].index(fic) + 1
        else:
            dfoutput['Artist 1 Ranking'][fic - 1] = ''
            dfoutput['Artist 1'][fic - 1] = ''
        name2 = extramatches[fic]
        if name2 != []:
            dfoutput['Artist 2 Ranking'][fic - 1] = people[name].index(fic) + 1
        else:
            dfoutput['Artist 2 Ranking'][fic - 1] = ''
            dfoutput['Artist 2'][fic - 1] = ''
    df2 = pd.DataFrame(data=dfoutput)
    df2.to_excel("matches.xlsx") 


if __name__ == '__main__':
    originalmatches = {}
    originalpeople, fics, originalmatches = import_input_file(sys.argv[1], 
                                                ["Preferred Name", "First Choice Fic", 
                                                    "Second Choice Fic", "Third Choice Fic", 
                                                    "Fourth Choice Fic", "Fifth Choice Fic"], 
                                                    int(sys.argv[2]), originalmatches)
    
    #originalpeople, fics, originalmatches = samplenames(38, originalmatches)
    extramatches = originalmatches.copy()
    finalmatches = originalmatches.copy()
    finalmatches = main_algorithm(originalpeople, fics, finalmatches)
    newpeople = originalpeople.copy()
    for k in list(finalmatches.values()):
        if k != []:
            newpeople.pop(k, None)
    extrafics = init_fics(38, newpeople)
    extramatches = main_algorithm(newpeople, extrafics, extramatches)
    export_matches(originalpeople, finalmatches, extramatches)
    
    