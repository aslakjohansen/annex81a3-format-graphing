#!/usr/bin/env python3

from sys import exit
from openpyxl import load_workbook
from graphviz import Graph

filename = 'High-Level Format Compatibility.xlsx'
xtab = [
    'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
    'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
    'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH'
]
ytab = list(range(5, 37))
nametable = {
    'ASHRAE 223P': 'A223p',
    'Control Description Language': 'CDL',
    'Linked Building Data': 'LBD',
}

def die (message):
    print(message)
    exit(1);

def index (x, y, safe=True):
    if safe and x>len(xtab): die('Err: X index of %u out of bounds' % x)
    if safe and y>len(ytab): die('Err: Y index of %u out of bounds' % y)
    
    if type(x)==int: x = str(xtab[x])
    if type(y)==int: y = str(ytab[y])
    
    return sheet['%s%s' % (x, y)]

#####################################################################
################################################################ main

wb = load_workbook(filename = filename)
sheet = wb['Ark1']

names = list(map(lambda x: index(x, '4', safe=False).value, xtab))

g = Graph(engine='neato', graph_attr={'overlap': 'false'})

# nodes for formats
for name in names:
    label = nametable[name] if name in nametable else name
    g.node(name, label)

# edges for integrations
for x in range(len(xtab)):
    for y in range(len(ytab)):
        value = index(x, y).value
        if value!=None and value.strip()!='':
            g.edge(names[x],
                   names[y])

# base format nodes
baseformats = []
for entries in filter(lambda value: value!=None and value.strip()!='', map(lambda x: index(x, '3', safe=False).value, range(len(xtab)))):
    for entry in entries.split(','):
        if not entry in baseformats:
            baseformats.append(entry)
for name in baseformats:
    label = name
    g.node(name, label, fillcolor='lightblue2', style='filled')

# edges for base formats
for x in range(len(names)):
    name = names[x]
    formats = index(x, '3', safe=False).value
    if formats==None or formats.strip()=='': continue
    for f in formats.split(','):
        print(name, f)
        g.edge(name, f)

# store
g.render('list2dot.gv')

