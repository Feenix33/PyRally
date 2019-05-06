from pyral import Rally, rallySettings, rallyWorkset
import sys
import argparse
import csv

"""
getPGM.py
From the list of the inputs, get the parents until you get to the PGM level or a dead end
Structure in Rally
"""

#Globals
errout = sys.stderr.write
gRallyInst = None
glbOutName = 'out.csv'
glbInName = 'in.txt'

gUSFields = [ 
        "TeamFeature.FormattedID",
        "TeamFeature.Name",
        ]

gGenFields = [ 
        "Parent.FormattedID",
        "Parent.Name",
        ]


def returnAttrib(item, attr, default=""):
    locAttr = attr.split('.')
    if len(locAttr) == 1:
        return getattr(item, locAttr[0], default)
    else:
        return getattr(getattr(item, locAttr[0], ""), locAttr[1], default)

def printHelp():
    print ("-h         Help")
    print ("-iFile     Use File as the input instead of command line arguments")
    print ("-oFile     Write to File instead of out.csv")
            

def getTokens(infile):
    with open(infile) as f:
        lines = [line.rstrip('\n') for line in f]
    return lines

def queryParent(id):
    if id[0] == "U": entityName = 'HierarchicalRequirement'
    elif id[0] == "T": entityName = 'PortfolioItem'
    elif id[0] == "F": entityName = 'PortfolioItem'
    elif id[0] == "P": entityName = 'PortfolioItem'
    else: return []

    queryString = 'FormattedID = "%s"' % id
    #print (queryString)
    response = gRallyInst.get(entityName, fetch=True, projectScopeDown=True, query=queryString)
    if response.resultCount > 0:
        for item in response:
            if id[0] == "U": 
                fieldList = gUSFields
            else:
                fieldList = gGenFields
            outrow = [item.FormattedID,] + [returnAttrib(item, field, default="") for field in fieldList]
            if id == item.FormattedID:
                return outrow

def main(args):
    global gRallyInst
    global glbOutName
    global glbInName

    #defaults
    glbOutName = 'out.csv'

    # Parse command line
    options = [opt for opt in args if opt.startswith('-')]
    args    = [arg for arg in args if arg not in options]
    #if len(args) < 1:
    #    errout('ERROR: Wrong number of arguments\n')
    #    printHelp()
    #    sys.exit(3)

    tokens = args # overwrite if -i

    for opt in options:
        if opt == '-h':
            printHelp()
            sys.exit(3)
        elif opt[:2] == '-i':
            if len(opt) > 2:
                glbInName = opt[2:]
                tokens = getTokens(glbInName)
        elif opt[:2] == '-o':
            if len(opt) > 2:
                glbOutName = opt[2:]


    # Rally setup
    #print ("Logging in")
    gRallyInst = Rally(
        server = 'rally1.rallydev.com',
        apikey = '_LhzUHJ1GQJQWkEYepqIJV9NO96FkErDpQvmHG4WQ',
        workspace = 'Sabre Production Workspace',
        project = 'Sabre')

    #print ("Start loop")

    with open (glbOutName, 'w', newline='', encoding='utf-8') as csvfile:
        outfile = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        outfile.writerow(['FormattedID', 'Parent.FormattedID','Parent.Name'])
        outrow = ['', '', '']
        for id in tokens:
            outrow[1] = id.upper()
            while len(outrow[1]) != 0  and outrow[1][:2] != "PG":
                outrow = queryParent(outrow[1])
                outfile.writerow(outrow)

if __name__ == '__main__':
    main(sys.argv[1:])
