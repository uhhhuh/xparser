#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import cPickle
import json
import csv


#import xparse
#ws = xparse.load_file('tests/test_data/test_book.xlsx')
#data_all = xparse.parse_person('A2:A787')
#blocks_by_p = xparse.make_blocks(data_all)

def pick_load(file_name):
    with open(file_name, 'rb') as fin:
        return cPickle.load(fin)

def pick_save(_from, file_name):
    with open(file_name, 'wb') as fout:
        cPickle.dump(_from, fout)


'''Minimum edit distance
medd("intention", "execution") => 8
medd("fish", "fihs") => 1
med("fish", "fihs") => 2
'''

def medd(target, source):
    '''Minimum edit distance with transposition (Damerau-Levenshtein)'''

    n = len(target) # row
    m = len(source) # column

    ins_cost = lambda x: 1
    del_cost = lambda x: 1
    sub_cost = lambda x, y: 0 if (x == y) else 2
    transp_cost = 1

    distance = [[0 for j in range(m+1)] for i in range(n+1)]

    for i in range(1, n+1):
        distance[i][0] = distance[i-1][0] + ins_cost(target[i-1])

    for j in range(1, m+1):
        distance[0][j] = distance[0][j-1] + del_cost(source[j-1])

    for i in range(1, n+1):
        for j in range(1, m+1):
            insert = distance[i-1][j] + ins_cost(target[i-1])
            subst = distance[i-1][j-1] + sub_cost(source[j-1],target[i-1])
            delete = distance[i][j-1] + del_cost(source[j-1])
            #transp = distance[i-2][j-2] + transp_cost(source[j-2],target[i-2])
            distance[i][j] = min(insert, subst, delete)

            if i and j and target[i-1] == source[j-2] and target[i-2] == source[j-1]:
                distance[i][j] = min(distance[i][j], distance[i-2][j-2] + transp_cost)
                
    return distance[n][m]

def med(target, source):
    '''Minimum edit distance (Levenshtein)'''
    
    n = len(target) # row
    m = len(source) # column

    ins_cost = lambda x: 1
    del_cost = lambda x: 1
    sub_cost = lambda x, y: 0 if (x == y) else 2

    
    distance = [[0 for j in range(m+1)] for i in range(n+1)]

    for i in range(1, n+1):
        distance[i][0] = distance[i-1][0] + ins_cost(target[i-1])

    for j in range(1, m+1):
        distance[0][j] = distance[0][j-1] + del_cost(source[j-1])

    for i in range(1, n+1):
        for j in range(1, m+1):
            insert = distance[i-1][j] + ins_cost(target[i-1])
            subst = distance[i-1][j-1] + sub_cost(source[j-1],target[i-1])
            delete = distance[i][j-1] + del_cost(source[j-1])
            distance[i][j] = min(insert, subst, delete)

    return distance[n][m]

def best_match(string, dictionary):
    candidates = {}
    for entry in dictionaries[dictionary].keys():
        candidates.update({entry: utils.medd(entry, string)})
        
    return min(candidates.iterkeys(), key=lambda k: candidates[k])


def load_dictionaries(dicts_dir):
    '''Load dictionaries from CSV-files and store in a dict.'''

    def key_to_int(value, conv=True):
        '''FIXME: ugly; convert all keys to ints'''
        if conv:
            try:
                return int(value)
            except Exception, err:
                #logger.error('{}. Couldn\'t convert a csv value [value] to integer'.format(value, err))
                return value
        else:
            return value

    def _load_csv(dict_name, convert=True):
        with open(dicts_dir + os.sep + dict_name + '.csv') as csvfile:
            reader = csv.DictReader(csvfile)
            mydict = [{v: key_to_int(k, conv=convert) for k,v in row.items()} for row in reader]
            for entry in mydict:
                mydicts[dict_name].update(entry)


    def load_csv(dict_name, convert=True):
        with open(dicts_dir + os.sep + dict_name + '.csv') as csvfile:
            reader = csv.DictReader(csvfile)
            mydict = [{v: key_to_int(k, conv=convert) for k,v in row.items() if v} for row in reader]
            for entry in mydict:
                mydicts[dict_name].update(entry)

    mydicts = {
        'relationType':{},
        'objectType': {},
        'ownershipType': {},
        'country': {}
        }
    
    for dict_name in mydicts.keys():
        load_csv(dict_name)

    mydicts.update({'none_values': {}})
    load_csv('none_values', convert=False)

    return mydicts
